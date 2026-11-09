"""엑셀 통계표 → 한글(hwpx) 보고서 페이지 자동 생성.

양식 hwpx 안에 완성된 예시 표 1개 페이지(--sample)를 찾아, 그 페이지의 문단·그래프 박스·통계표를
틀로 삼아 나머지 표들을 같은 모양으로 이어 붙인다. 표 테두리·글자 모양·열 너비도 예시 표에서 읽는다.

  python3 build_report.py --template 양식.hwpx --excel 통계표.xlsx --list
  python3 build_report.py --template 양식.hwpx --excel 통계표.xlsx --sample 2-1 \
      --groups "비자유형,국적,연령대,성별" --out 결과.hwpx
"""
import argparse, copy, os, re, sys, tempfile
from xml.sax.saxutils import escape
from lxml import etree
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hwtx_common as H
from hwtx_common import HP, NS, wrap

ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
ap.add_argument('--template', required=True, help='양식 hwpx (예시 표 1개 페이지가 완성된 파일)')
ap.add_argument('--excel', required=True, help='통계표 xlsx')
ap.add_argument('--out', help='결과 hwpx')
ap.add_argument('--sample', help='양식 예시 표 번호 (예: 2-1)')
ap.add_argument('--start', help='생성 시작 표 번호 (기본: sample 다음 표)')
ap.add_argument('--end', help='생성 마지막 표 번호 (기본: 끝)')
ap.add_argument('--groups', help='특성별 결과 구분, 쉼표 구분·적은 순서대로 (예: 비자유형,국적,연령대,성별)')
ap.add_argument('--min-n', type=int, default=10, help='특성별 문구에 넣을 최소 사례수 (기본 10)')
ap.add_argument('--abbr', action='append', default=[], help='문장 속 보기 축약 "원문=축약" (여러 번 가능)')
ap.add_argument('--title', action='append', default=[], help='표 제목 바꾸기 "표번호 또는 엑셀제목=새 제목" (같은 번호가 여러 개면 "3-1|Base원문=새 제목")')
ap.add_argument('--base', action='append', default=[], help='Base 문구 풀어쓰기 "엑셀 BASE 원문=쉬운 말" (설문지 참고, 여러 번 가능)')
ap.add_argument('--fill-sample', action='store_true', help='예시 페이지의 문구·Base도 새로 작성')
ap.add_argument('--scale-full', action='store_true', help='척도 세부표(x-y-z)도 문구+그래프 페이지 포함')
ap.add_argument('--list', action='store_true', help='표 목록·특성 구분만 출력하고 종료')
ap.add_argument('--texts', help='[웹앱] 표 순번별 분석문구 JSON {"순번": {"overall": [...], "group": [...]}} — 있으면 규칙 문구 대신 사용')
A = ap.parse_args()
for a in A.abbr:
    k, v = a.split('=', 1); H.ABBR[k] = v
for a in A.base:
    k, v = a.split('=', 1); H.BASE_MAP[k.strip()] = v.strip()
for a in A.title:
    k, v = a.split('=', 1); H.TITLE_MAP[k.strip()] = v.strip()

tables = H.classify(H.parse_excel(A.excel), scale_full=A.scale_full)
for _i, _t in enumerate(tables): _t['_idx'] = str(_i)
import json as _json
TEXTS = _json.load(open(A.texts, encoding='utf-8')) if A.texts else {}
def OV(t): return (TEXTS.get(t['_idx']) or {}).get('overall') or H.overall_paras(t)
def GR(t): return (TEXTS.get(t['_idx']) or {}).get('group') or H.group_paras(t, GROUPS, A.min_n)
if A.list:
    for t in tables:
        print(f"<표{t['num']}> {t['name']} | Base: {t['base']} | 보기 {len(t['item_idx'])}개 | {t['type']}"
              + (' | 중복응답' if t['multi'] else '') + (f" | 평균({t['unit']})" if t['mean_idx'] else ''))
    grp = {}
    for t in tables:
        for r in t['rows'][1:]: grp.setdefault(r['grp'], {}).setdefault(r['label'], r['n'])
    for g, d in grp.items(): print(f"특성 구분: {g} → " + ', '.join(f'{k}({v})' for k, v in d.items()))
    sys.exit(0)
if not A.sample or not A.groups:
    sys.exit('--sample 과 --groups 가 필요합니다 (특성별 분석 구분은 사용자에게 먼저 물어볼 것)')
GROUPS = [g.strip() for g in A.groups.split(',') if g.strip()]
OUT = A.out or '/mnt/user-data/outputs/' + os.path.splitext(os.path.basename(A.template))[0] + '_완성.hwpx'

i_s = next(i for i, t in enumerate(tables) if t['num'] == A.sample)
sample = tables[i_s]
i0 = next(i for i, t in enumerate(tables) if t['num'] == A.start) if A.start else i_s + 1
i1 = max(i for i, t in enumerate(tables) if t['num'] == A.end) + 1 if A.end else len(tables)
todo = tables[i0:i1]

# ---------------------------------------------------------------- 양식 페이지 찾기
D = H.Doc(A.template, tempfile.mkdtemp())
P = D.paras(); txt = D.text
NAME = sample['name']
cand = [i for i, p in enumerate(P) if txt(p) in (NAME, sample['title']) and not D.has_tbl(p)]
if not cand: sys.exit(f"양식에서 예시 표 제목 '{NAME}'을 찾지 못했습니다. 양식의 소제목을 엑셀 표 제목과 같게 맞춰 주세요.")
i_sub = cand[0]
i_all = next(i for i in range(i_sub + 1, len(P)) if txt(P[i]).endswith('전체 결과'))
i_cap = next(i for i in range(i_all + 1, len(P)) if txt(P[i]) in (NAME, sample['title']))
i_box = next(i for i in range(i_cap + 1, len(P)) if D.has_tbl(P[i]))
i_grp = next(i for i in range(i_box + 1, len(P)) if txt(P[i]).endswith('특성별 결과'))
i_tcap = next(i for i in range(i_grp + 1, len(P)) if txt(P[i]) in (NAME, sample['title']))
i_tbl = next(i for i in range(i_tcap + 1, len(P)) if D.has_tbl(P[i]))
g_txt = [i for i in range(i_grp + 1, i_tcap - 1) if txt(P[i])]
T = dict(pb1=P[i_sub - 1], sub=P[i_sub], all=P[i_all], txt=P[i_all + 1], blank1=P[i_all + 2], cap=P[i_cap],
         box=P[i_box], blank2=P[i_box + 1], grp=P[i_grp], gtxt=P[g_txt[0]], pb2=P[i_tcap - 1], tcap=P[i_tcap],
         unit=P[i_tcap + 1], tbl=P[i_tbl])
G_BLANKS = [P[i] for i in range(g_txt[-1] + 1, i_tcap - 1)]
T_TXT = [copy.deepcopy(P[i]) for i in range(i_all + 1, i_cap - 1) if txt(P[i])] or [copy.deepcopy(T['txt'])]

# ---------------------------------------------------------------- 예시 통계표에서 서식 읽기
src_tbl = T['tbl'].find('.//' + HP + 'tbl')
cells = {}
for tc in src_tbl.iter(HP + 'tc'):
    a = tc.find(HP + 'cellAddr'); cells[(int(a.get('rowAddr')), int(a.get('colAddr')))] = tc
def info(tc):
    p = tc.find('.//' + HP + 'p'); run = p.find(HP + 'run'); seg = p.find('.//' + HP + 'lineseg')
    sp = tc.find(HP + 'cellSpan'); sz = tc.find(HP + 'cellSz')
    return dict(bf=tc.get('borderFillIDRef'), pp=p.get('paraPrIDRef'), st=p.get('styleIDRef'), cp=run.get('charPrIDRef'),
                vs=seg.get('vertsize'), sp=seg.get('spacing'), bl=seg.get('baseline'),
                rs=int(sp.get('rowSpan')), cs=int(sp.get('colSpan')), w=int(sz.get('width')), h=int(sz.get('height')),
                margin=etree.tostring(tc.find(HP + 'cellMargin')).decode())
hr = next(r for r in range(1, 6) if (r, 0) in cells and txt(cells[(r, 0)]) == '전체')       # 헤더 행 수
S = dict(corner=info(cells[(0, 0)]), head=info(cells[(0, 2)]), tlab=info(cells[(hr, 0)]), tn=info(cells[(hr, 2)]),
         tval=info(cells[(hr, 3)]), grp=info(cells[(hr + 1, 0)]), lab=info(cells[(hr + 1, 1)]), n=info(cells[(hr + 1, 2)]),
         val=info(cells[(hr + 1, 3)]))
g_rows = [r for (r, c) in cells if c == 0 and r > hr]
big = next((r for r in sorted(g_rows) if info(cells[(r, 0)])['rs'] >= 3), None)
if big is not None:
    span = info(cells[(big, 0)])['rs']
    BF_FIRST, BF_MID, BF_LAST = info(cells[(big, 1)])['bf'], info(cells[(big + 1, 1)])['bf'], info(cells[(big + span - 1, 1)])['bf']
else:
    BF_FIRST = BF_MID = BF_LAST = S['lab']['bf']
TOTAL_W = int(src_tbl.find(HP + 'sz').get('width'))
W0, W1, WN = S['grp']['w'], S['lab']['w'], S['n']['w']
RH = S['val']['h']
EM_H = D.em.get(S['head']['cp'], 857); EM_D = D.em.get(S['val']['cp'], 767)

def tc_xml(r, c, rs, cs, w, h, role, text, header=False):
    s = S[role]; em = EM_H if header else EM_D
    lines = wrap(text, max(w - 400, 500), em) if text else [0]
    vs, sp, bl = int(s['vs']), int(s['sp']), s['bl']
    segs = ''.join(f'<hp:lineseg textpos="{st}" vertpos="{j*(vs+sp)}" vertsize="{vs}" textheight="{vs}" baseline="{bl}" '
                   f'spacing="{sp}" horzpos="0" horzsize="{w-3}" flags="393216"/>' for j, st in enumerate(lines))
    run = f'<hp:run charPrIDRef="{s["cp"]}"><hp:t>{escape(text)}</hp:t></hp:run>'
    return (f'<hp:tc name="" header="{1 if header else 0}" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="{s["bf"]}">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
            f'<hp:p id="2147483648" paraPrIDRef="{s["pp"]}" styleIDRef="{s["st"]}" pageBreak="0" columnBreak="0" merged="0">{run}<hp:linesegarray>{segs}</hp:linesegarray></hp:p></hp:subList>'
            f'<hp:cellAddr colAddr="{c}" rowAddr="{r}"/><hp:cellSpan colSpan="{cs}" rowSpan="{rs}"/><hp:cellSz width="{w}" height="{h}"/>'
            + s['margin'].replace(' xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"', '') + '</hp:tc>')

def with_bf(role, bf):
    S[role + '_x'] = dict(S[role], bf=bf); return role + '_x'

TBL_ID = [1209800000]
fmt = H.xfmt       # 엑셀 화면 값과 같은 반올림

def build_table(t):
    ncol = len(t['cols'])
    w0, w1 = (W0, W1) if ncol <= 6 else (int(W0 * 0.75), int(W1 * 0.95))
    wd_total = TOTAL_W - w0 - w1 - WN
    wd = [wd_total // ncol] * ncol; wd[-1] += wd_total - sum(wd)
    two = t['two']; hdr_rows = 2 if two else 1
    line = lambda n: 566 + int(S['head']['vs']) + (n - 1) * (int(S['head']['vs']) + int(S['head']['sp']))
    hl = lambda text, w: len(wrap(text, max(w - 400, 500), EM_H))
    if not two:
        hh = [line(max([hl(c['top'], wd[k]) for k, c in enumerate(t['cols'])] + [1]))]
    else:
        h2 = line(max([hl(c['sub'], wd[k]) for k, c in enumerate(t['cols']) if c['sub']] + [1]))
        h1 = line(max([hl(c['top'], wd[k]) for k, c in enumerate(t['cols']) if not c['sub']] + [1]))
        tot = max(h1, h2 + S['corner']['h']); hh = [tot - h2, h2]
    HH = sum(hh)
    r0 = [tc_xml(0, 0, hdr_rows, 2, w0 + w1, HH, 'corner', '전체', True),
          tc_xml(0, 2, hdr_rows, 1, WN, HH, 'head', '사례수', True)]
    r1 = []; k = 0
    while k < ncol:
        c = t['cols'][k]
        if two and c['sub']:
            span = 1
            while k + span < ncol and t['cols'][k + span]['sub'] and t['cols'][k + span]['top'] == c['top']: span += 1
            r0.append(tc_xml(0, 3 + k, 1, span, sum(wd[k:k + span]), hh[0], 'head', c['top'], True))
            for j in range(span): r1.append(tc_xml(1, 3 + k + j, 1, 1, wd[k + j], hh[1], 'head', t['cols'][k + j]['sub'], True))
            k += span
        else:
            r0.append(tc_xml(0, 3 + k, hdr_rows, 1, wd[k], HH, 'head', c['top'], True)); k += 1
    rows = ['<hp:tr>' + ''.join(r0) + '</hp:tr>'] + (['<hp:tr>' + ''.join(r1) + '</hp:tr>'] if two else [])
    ri = hdr_rows; total_h = HH; tot = t['rows'][0]
    cl = [tc_xml(ri, 0, 1, 2, w0 + w1, RH, 'tlab', '전체'), tc_xml(ri, 2, 1, 1, WN, RH, 'tn', f"{tot['n']:,}")]
    cl += [tc_xml(ri, 3 + j, 1, 1, wd[j], RH, 'tval', fmt(v, d)) for j, (v, d) in enumerate(tot['vals'])]
    rows.append('<hp:tr>' + ''.join(cl) + '</hp:tr>'); ri += 1; total_h += RH
    for g in dict.fromkeys(r['grp'] for r in t['rows'][1:]):
        rs = [r for r in t['rows'][1:] if r['grp'] == g]
        hs = [max(RH, line(len(wrap(r['label'], w1 - 400, EM_D)))) for r in rs]
        for j, r in enumerate(rs):
            bf = BF_LAST if (len(rs) == 1 or j == len(rs) - 1) else (BF_FIRST if j == 0 else BF_MID)
            cl = []
            if j == 0: cl.append(tc_xml(ri, 0, len(rs), 1, w0, sum(hs), 'grp', g))
            cl.append(tc_xml(ri, 1, 1, 1, w1, hs[j], with_bf('lab', bf), r['label']))
            cl.append(tc_xml(ri, 2, 1, 1, WN, hs[j], with_bf('n', bf), f"{r['n']:,}"))
            cl += [tc_xml(ri, 3 + q, 1, 1, wd[q], hs[j], with_bf('val', bf), fmt(v, d)) for q, (v, d) in enumerate(r['vals'])]
            rows.append('<hp:tr>' + ''.join(cl) + '</hp:tr>'); ri += 1; total_h += hs[j]
    new = copy.deepcopy(src_tbl)
    for tr in new.findall(HP + 'tr'): new.remove(tr)
    TBL_ID[0] += 1
    new.set('id', str(TBL_ID[0])); new.set('rowCnt', str(ri)); new.set('colCnt', str(3 + ncol))
    new.find(HP + 'sz').set('height', str(total_h))
    for x in etree.fromstring(f'<r xmlns:hp="{NS["hp"]}">' + ''.join(rows) + '</r>'): new.append(x)
    return new, total_h

# ---------------------------------------------------------------- 페이지 조립
def table_paras(t):
    out = [copy.deepcopy(T['pb2'])]
    p = copy.deepcopy(T['tcap']); D.set_text(p, t['title']); out.append(p)
    p = copy.deepcopy(T['unit']); D.set_text(p, H.unit_text(t)); out.append(p)
    p = copy.deepcopy(T['tbl']); run = p.find(HP + 'run')
    new, th = build_table(t); run.replace(run.find(HP + 'tbl'), new)
    seg = p.find(HP + 'linesegarray').find(HP + 'lineseg')
    seg.set('vertsize', str(th)); seg.set('textheight', str(th)); seg.set('baseline', str(int(th * 0.85)))
    out.append(p); return out

def text_paras(proto, lines):
    out = []
    for s in lines:
        p = copy.deepcopy(proto); D.set_text(p, s); out.append(p)
    return out

def box_para(t):
    p = copy.deepcopy(T['box'])
    inner = p.find('.//' + HP + 'tbl').find(HP + 'tr').find('.//' + HP + 'p')
    D.set_text(inner, f"(Base: {t['base']}, " + H.unit_text(t)[1:])
    TBL_ID[0] += 1; p.find('.//' + HP + 'tbl').set('id', str(TBL_ID[0]))
    return p

def full_paras(t):
    out = [copy.deepcopy(T['pb1'])]
    p = copy.deepcopy(T['sub']); D.set_text(p, t['title']); out.append(p)
    p = copy.deepcopy(T['all']); D.sub_head(p, '가. 전체 결과'); out.append(p)
    t['_ov'] = OV(t); t['_gr'] = GR(t)
    out += text_paras(T_TXT[0], t['_ov'])
    out.append(copy.deepcopy(T['blank1']))
    p = copy.deepcopy(T['cap']); D.set_text(p, t['title']); out.append(p)
    out.append(box_para(t)); out.append(copy.deepcopy(T['blank2']))
    p = copy.deepcopy(T['grp']); D.sub_head(p, '나. 특성별 결과'); out.append(p)
    out += text_paras(T['gtxt'], t['_gr'])
    out += [copy.deepcopy(b) for b in G_BLANKS]
    out += table_paras(t)
    return out

new_paras = []
for t in todo:
    new_paras += table_paras(t) if t['type'] == 'scale' else full_paras(t)

D.sub_head(T['all'], '가. 전체 결과'); D.sub_head(T['grp'], '나. 특성별 결과')
if A.fill_sample:
    olds = [P[i] for i in range(i_all + 1, i_cap - 1) if txt(P[i])]
    for p in olds[1:]: D.root.remove(p)
    ov = OV(sample)
    for s_ in reversed(ov[1:]):
        p = copy.deepcopy(T['txt']); D.set_text(p, s_); T['txt'].addnext(p)
    D.set_text(T['txt'], ov[0])
    gl = GR(sample); olds = [P[i] for i in g_txt]
    for p in olds[1:]: D.root.remove(p)
    D.set_text(olds[0], gl[0]) if gl else None
    for s_ in reversed(gl[1:]):
        p = copy.deepcopy(T['gtxt']); D.set_text(p, s_); olds[0].addnext(p)
    inner = T['box'].find('.//' + HP + 'tbl').find(HP + 'tr').find('.//' + HP + 'p')
    D.set_text(inner, f"(Base: {sample['base']}, " + H.unit_text(sample)[1:])
at = D.root.index(T['tbl']) + 1
for j, p in enumerate(new_paras): D.root.insert(at + j, p)
D.save(OUT)

for t in todo:
    print(f"\n## <표{t['num']}> {t['title']} [{t['type']}]")
    for s_ in t.get('_ov', []) + t.get('_gr', []): print('  -', s_)
print(f"\n표 {len(todo)}개 생성 → {OUT}")
