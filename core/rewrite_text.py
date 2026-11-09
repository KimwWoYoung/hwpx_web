"""사용자가 한글에서 이미 손본 보고서(hwpx)의 분석문구만 다시 쓴다. 표·그래프·직접 고친 다른 부분은 그대로 둔다.

  python3 rewrite_text.py --hwpx 수정본.hwpx --excel 통계표.xlsx --sample 2-1 \
      --groups "비자유형,국적,연령대,성별" --part both --keep 4 --out 결과.hwpx

--part   overall(가. 전체 결과) / group(나. 특성별 결과) / both
--keep N 앞에서부터 N개 페이지는 건드리지 않음 (사용자가 직접 다듬은 페이지)
--only   특정 표 번호만 다시 씀 (예: --only 2-3,3-6)

페이지와 표의 연결: '가. 전체 결과'가 나오는 순서 = 엑셀의 표 순서(척도 세부표 제외).
한 페이지의 특성별 문단에 같은 구분(예: '(비자유형)')이 두 번 나오면 하위표 2개를 합친 페이지로 보고 표 2개를 연결한다.
"""
import argparse, copy, os, re, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hwtx_common as H

ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
ap.add_argument('--hwpx', required=True); ap.add_argument('--excel', required=True); ap.add_argument('--out', required=True)
ap.add_argument('--sample', required=True, help='보고서 첫 페이지의 표 번호')
ap.add_argument('--groups', required=True)
ap.add_argument('--part', default='both', choices=['overall', 'group', 'both'])
ap.add_argument('--keep', type=int, default=0)
ap.add_argument('--only', default='')
ap.add_argument('--min-n', type=int, default=10)
ap.add_argument('--abbr', action='append', default=[])
ap.add_argument('--title', action='append', default=[], help='표 제목 바꾸기 "표번호 또는 엑셀제목=새 제목" (같은 번호가 여러 개면 "3-1|Base원문=새 제목")')
ap.add_argument('--base', action='append', default=[], help='Base 문구 풀어쓰기 "원문=쉬운 말"')
ap.add_argument('--texts', help='[웹앱] 표 순번별 AI 문구 JSON {"순번": {"overall": [...], "group": [...]}}')
ap.add_argument('--plan', help='[웹앱] 고치지 않고 페이지↔표 연결과 현재 문구만 JSON으로 저장')
A = ap.parse_args()
for a in A.abbr:
    k, v = a.split('=', 1); H.ABBR[k] = v
for a in A.base:
    k, v = a.split('=', 1); H.BASE_MAP[k.strip()] = v.strip()
for a in A.title:
    k, v = a.split('=', 1); H.TITLE_MAP[k.strip()] = v.strip()
GROUPS = [g.strip() for g in A.groups.split(',') if g.strip()]
ONLY = {x.strip() for x in A.only.split(',') if x.strip()}

tables = H.classify(H.parse_excel(A.excel))
for _i, _t in enumerate(tables): _t['_idx'] = str(_i)
import json as _json
TEXTS = _json.load(open(A.texts, encoding='utf-8')) if A.texts else {}
PLAN = []
i0 = next(i for i, t in enumerate(tables) if t['num'] == A.sample)
jobs = [t for t in tables[i0:] if t['type'] != 'scale']

D = H.Doc(A.hwpx, tempfile.mkdtemp())
P = D.paras(); txt = D.text
heads = [i for i, p in enumerate(P) if txt(p).endswith('전체 결과')]
ji = 0
for n, hi in enumerate(heads):
    gi = next(i for i in range(hi, len(P)) if txt(P[i]).endswith('특성별 결과'))
    ov = []; k = hi + 1
    while k < gi and txt(P[k]): ov.append(P[k]); k += 1
    gp = []; k = gi + 1
    while k < len(P) and txt(P[k]).startswith('('): gp.append(P[k]); k += 1
    first = re.match(r'\(([^)]+)\)', txt(gp[0])).group(1) if gp else None
    cnt = max(1, sum(1 for p in gp if txt(p).startswith(f'({first})'))) if first else 1
    ts = jobs[ji:ji + cnt]; ji += cnt
    PLAN.append(dict(page=n + 1, title=txt(P[hi - 1]), tables=[t['_idx'] for t in ts], nums=[t['num'] for t in ts],
                     merged=len(ts) > 1, overall=[txt(p) for p in ov], group=[txt(p) for p in gp],
                     keep=n < A.keep or bool(ONLY and not any(t['num'] in ONLY for t in ts))))
    if A.plan: continue
    if n < A.keep or (ONLY and not any(t['num'] in ONLY for t in ts)): continue
    merged = len(ts) > 1
    if A.part in ('overall', 'both') and ov:
        lines = []
        for t in ts:
            ai = (TEXTS.get(t['_idx']) or {}).get('overall')
            lines += ai if ai else ([H.overall_tagged(t, t.get('suffix', t['title']))] if merged else H.overall_paras(t, subj=txt(P[hi - 1]) or None))
        while len(ov) < len(lines):
            c = copy.deepcopy(ov[-1]); ov[-1].addnext(c); ov.append(c)
        for p in ov[len(lines):]: D.root.remove(p)
        for p, s in zip(ov, lines): D.set_text(p, s)
    if A.part in ('group', 'both') and gp:
        lines = []
        for t in ts:
            ai = (TEXTS.get(t['_idx']) or {}).get('group')
            lines += ai if ai else H.group_paras(t, GROUPS, A.min_n, tag=t.get('suffix', '') if merged else '')
        while len(gp) < len(lines):
            c = copy.deepcopy(gp[-1]); gp[-1].addnext(c); gp.append(c)
        for p in gp[len(lines):]: D.root.remove(p)
        for p, s in zip(gp, lines): D.set_text(p, s)
    print(f"\n## {' + '.join('<표' + t['num'] + '>' for t in ts)} {txt(P[hi - 1])}")
    for p in (ov if A.part != 'group' else []) + (gp if A.part != 'overall' else []):
        if p.getparent() is not None: print('  -', txt(p))
if A.plan:
    _json.dump(PLAN, open(A.plan, 'w', encoding='utf-8'), ensure_ascii=False); print(f'pages {len(PLAN)}'); sys.exit(0)
if H.BASE_MAP:                                   # 그래프 박스 Base 문구도 풀어 쓴 말로 교체
    nb = 0
    for p in list(D.root.iter(H.HP + 'p')):
        if p.find('.//' + H.HP + 'p') is not None: continue
        t_ = txt(p)
        if t_.startswith('(Base:'):
            for k_, v_ in H.BASE_MAP.items():
                if k_ in t_: D.set_text(p, t_.replace(k_, v_)); nb += 1; break
    print(f'Base 문구 {nb}곳 교체')
D.save(A.out)
print(f"\n페이지 {len(heads)}개 / 표 {ji}개 연결 → {A.out}")
if ji != len(jobs): print(f"주의: 엑셀 표 {len(jobs)}개와 페이지 연결 수 {ji}개가 다릅니다. 범위를 확인하세요.")
