"""결과물 숫자를 원본 엑셀과 칸마다 대조한다 (전달 전 필수). 불일치가 0개일 때만 전달한다.

  python3 verify.py --excel 통계표.xlsx --hwpx 결과.hwpx [--pptx 결과.pptx] [--title …] [--base …]
  (--title/--base 는 build_report.py에 준 것과 똑같이 넘긴다 — 표 제목으로 엑셀 표를 찾기 때문)

검증 기준값은 만들 때와 같은 엑셀식 반올림(xround: 유효숫자 15자리 → 사사오입)으로 만든다.
- 한글: 통계표마다 사례수·모든 값 칸, 행 수·열 수 / 분석문구 속 %·평균 수치가 그 표에 있는 값인지(경고)
- PPT: 차트 캐시 값, 임베디드 엑셀 값, 평균 배지
끝 코드: 불일치 있으면 1, 없으면 0. 대조하지 못한 표·그래프는 이름과 개수를 따로 출력한다.
"""
import argparse, io, os, re, sys, zipfile
from lxml import etree
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hwtx_common as H
from hwtx_common import HP, xfmt

ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
ap.add_argument('--excel', required=True)
ap.add_argument('--hwpx'); ap.add_argument('--pptx')
ap.add_argument('--title', action='append', default=[]); ap.add_argument('--base', action='append', default=[])
ap.add_argument('--abbr', action='append', default=[])
ap.add_argument('--scale-full', action='store_true')
ap.add_argument('--include-scale-detail', action='store_true')
A = ap.parse_args()
for a in A.abbr:
    k, v = a.split('=', 1); H.ABBR[k] = v
for a in A.base:
    k, v = a.split('=', 1); H.BASE_MAP[k.strip()] = v.strip()
for a in A.title:
    k, v = a.split('=', 1); H.TITLE_MAP[k.strip()] = v.strip()

tables = H.classify(H.parse_excel(A.excel), scale_full=A.scale_full)
by_title = {}
for t in tables: by_title.setdefault(t['title'], []).append(t)
used = {}
def find(title):
    title = re.sub(r'^<표\s*[\d\-]+>\s*', '', title.strip())
    lst = by_title.get(title) or by_title.get(title.replace(' ', '')) or []
    if not lst:                                           # 공백 차이만 있는 제목
        key = title.replace(' ', '')
        lst = next((v for k, v in by_title.items() if k.replace(' ', '') == key), [])
    if not lst: return None
    i = used.get(id(lst), 0); used[id(lst)] = i + 1
    return lst[min(i, len(lst) - 1)]

ERR = []; SKIP = []; WARN = []; OK = [0]
def err(where, what): ERR.append(f'{where}: {what}')

# ====================================================================== 한글
def tc_text(tc): return ''.join(tc.xpath('.//hp:t/text()', namespaces=H.NS)).strip()

def check_hwpx(path):
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read('Contents/section0.xml'))
    paras = list(root); txt = lambda p: ''.join(p.xpath('.//hp:t/text()', namespaces=H.NS)).strip()
    buf = []; n_tbl = 0
    for i, p in enumerate(paras):
        tbl = p.find('.//' + HP + 'tbl')
        if tbl is None:
            s = txt(p)
            if s: buf.append(s)
            continue
        cells = {}
        for tc in tbl.iter(HP + 'tc'):
            a = tc.find(HP + 'cellAddr'); sp = tc.find(HP + 'cellSpan')
            cells[(int(a.get('rowAddr')), int(a.get('colAddr')))] = (tc_text(tc), int(sp.get('colSpan')))
        if not any(v[0] == '사례수' for (r, c), v in cells.items() if r == 0):
            continue                                              # 그래프 박스 등 통계표가 아닌 표
        n_tbl += 1
        cap = next((txt(paras[j]) for j in range(i - 1, max(i - 6, -1), -1)
                    if txt(paras[j]) and not txt(paras[j]).startswith('(단위') and paras[j].find('.//' + HP + 'tbl') is None), '')
        t = find(cap)
        where = f'[한글] {cap or f"통계표 {n_tbl}번째"}'
        if t is None:
            SKIP.append(f'{where} (엑셀에서 같은 제목의 표를 찾지 못함)'); buf = []; continue
        hr = min(r for (r, c), v in cells.items() if c == 0 and v[0] == '전체' and r > 0) if any(
            c == 0 and v[0] == '전체' and r > 0 for (r, c), v in cells.items()) else None
        if hr is None:
            SKIP.append(f'{where} (전체 행을 찾지 못함)'); buf = []; continue
        got = []
        for r in sorted({r for (r, c) in cells if r >= hr}):
            row = {c: v for (rr, c), v in cells.items() if rr == r}
            lab = '전체' if (0 in row and row[0][1] == 2) else row.get(1, ('', 1))[0]
            got.append((lab, row.get(2, ('', 1))[0], [row[c][0] for c in sorted(row) if c >= 3]))
        exp = [(r['label'], f"{r['n']:,}", [xfmt(v, d) for v, d in r['vals']]) for r in t['rows']]
        # [웹앱] 열 이름으로 맞추기: 사람이 열 순서를 바꾼 표도 칸마다 대조 (순서가 다르면 경고)
        nk = lambda x: re.sub(r'\s+', '', str(x))
        hcols = sorted({c for (r, c) in cells if c >= 3 and r < hr})
        hdr = []
        for c in hcols:
            rr = max((r for (r, cc) in cells if cc == c and r < hr), default=None)
            hdr.append(cells[(rr, c)][0] if rr is not None else '')
        xl = [nk(H.iname(t, k)) for k in range(len(t['cols']))]
        order = None
        if len(hdr) == len(xl) and [nk(h) for h in hdr] != xl and sorted(nk(h) for h in hdr) == sorted(xl) and len(set(xl)) == len(xl):
            order = [xl.index(nk(h)) for h in hdr]
            WARN.append(f'{where}: 열 순서가 엑셀과 다름 (열 이름으로 맞춰 대조함)')
            exp = [(a, b, [v[k] for k in order]) for a, b, v in exp]
        names = [H.iname(t, k) for k in (order or range(len(t['cols'])))]
        if len(got) != len(exp):
            err(where, f'행 수 다름 (한글 {len(got)}행, 엑셀 {len(exp)}행)')
        for (gl, gn, gv), (el, en, ev) in zip(got, exp):
            if gl != el: err(where, f"행 이름 '{gl}' ≠ 엑셀 '{el}'")
            if gn != en: err(where, f"{el} 사례수 {gn} ≠ 엑셀 {en}")
            if len(gv) != len(ev): err(where, f'{el} 값 칸 수 {len(gv)} ≠ 엑셀 {len(ev)}'); continue
            for k, (a, b) in enumerate(zip(gv, ev)):
                if a != b: err(where, f"{el} × {names[k]}: 한글 {a} ≠ 엑셀 {b}")
                else: OK[0] += 1
        # 분석문구 속 수치 (경고) — 표에 실제 있는 값인지
        allowed = {xfmt(v, d) for r in t['rows'] for v, d in r['vals']} | {xfmt(v, 1) for r in t['rows'] for v, d in r['vals']}
        start = next((k for k, s in enumerate(buf) if s.endswith('전체 결과')), None)
        for s in (buf[start + 1:] if start is not None else []):
            if s.startswith('(Base') or s.startswith('(단위') or '합하면' in s: continue
            for m in re.finditer(r'(\d+\.\d+)(%(?!p)|년(?! 더)|점(?! 더))', s):
                if m.group(1) not in allowed:
                    WARN.append(f"{where}: 문구 수치 {m.group(0)} 가 표에 없음 → \"{s[:60]}…\"")
        buf = []
    return n_tbl

# ====================================================================== PPT
PNS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
       'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
       'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
def check_pptx(path):
    import openpyxl
    z = zipfile.ZipFile(path); names = set(z.namelist())
    def rels(part):
        d, f = os.path.split(part); rp = os.path.join(d, '_rels', f + '.rels')
        if rp not in names: return []
        return [(m.group(1), os.path.normpath(os.path.join(d, m.group(2))))
                for m in re.finditer(r'Type="[^"]*/(\w+)" Target="([^"]+)"', z.read(rp).decode())]
    pres = z.read('ppt/presentation.xml').decode(); prels = dict(re.findall(r'Id="(\w+)"[^>]*Target="([^"]+)"', z.read('ppt/_rels/presentation.xml.rels').decode()))
    order = ['ppt/' + prels[r] for r in re.findall(r'<p:sldId [^>]*r:id="(\w+)"', pres)]
    n = 0
    for sp in order:
        n += 1; rs = dict((typ, tgt) for typ, tgt in rels(sp))
        note = ''
        if 'notesSlide' in rs:
            note = ''.join(etree.fromstring(z.read(rs['notesSlide'])).xpath('.//a:t/text()', namespaces=PNS)).strip()
        m = re.match(r'(<표\s*[\d\-]+>\s*.+?)\s*\(n=', note)
        where = f'[PPT] 슬라이드 {n} {m.group(1) if m else ""}'.strip()
        if 'chart' not in rs:
            SKIP.append(f'{where} (차트 없음 — 이미지 그래프는 대조 불가)'); continue
        t = find(m.group(1)) if m else None
        if t is None:
            SKIP.append(f'{where} (노트에서 표 이름을 찾지 못함)'); continue
        tot = t['rows'][0]
        exp = {}
        for k, c in enumerate(t['cols']):
            if '평균' in c['top']: continue
            v = xfmt(tot['vals'][k][0], 1)
            exp.setdefault(H.lab(H.iname(t, k)), v); exp.setdefault(re.sub(r'^[①-⑩]\s*', '', c['top']), v)
        croot = etree.fromstring(z.read(rs['chart']))
        ser = croot.find('.//c:ser', PNS)
        cats = [x.text for x in ser.find('c:cat', PNS).iter('{%s}v' % PNS['c'])]
        vals = [x.text for x in ser.find('c:val', PNS).iter('{%s}v' % PNS['c'])]
        for lab, v in zip(cats, vals):
            if lab not in exp: err(where, f"그래프 항목 '{lab}'이 엑셀 표에 없음"); continue
            if xfmt(float(v), 1) != exp[lab]: err(where, f"차트 {lab}: {v} ≠ 엑셀 {exp[lab]}")
            else: OK[0] += 1
        emb = next((tgt for typ, tgt in rels(rs['chart']) if typ == 'package'), None)
        if emb:
            ws = openpyxl.load_workbook(io.BytesIO(z.read(emb)), data_only=True).worksheets[0]
            c = 2
            while ws.cell(1, c).value is not None:
                lab, v = str(ws.cell(1, c).value), ws.cell(2, c).value
                if lab in exp and (not isinstance(v, (int, float)) or xfmt(v, 1) != exp[lab]):
                    err(where, f"임베디드 엑셀 {lab}: {v} ≠ 엑셀 {exp[lab]}")
                c += 1
        if t['mean_idx']:
            v, d = tot['vals'][t['mean_idx'][0]]
            sroot = etree.fromstring(z.read(sp))
            badge = [''.join(e.xpath('.//a:t/text()', namespaces=PNS)) for e in sroot.iter('{%s}sp' % PNS['p'])
                     if e.find('.//p:cNvPr', PNS).get('name') == '모서리가 둥근 직사각형 20']
            if badge and xfmt(v, d) not in badge[0]: err(where, f"평균 배지 '{badge[0].strip()}' ≠ 엑셀 {xfmt(v, d)}")
    return n

if A.hwpx: print(f'한글 통계표 {check_hwpx(A.hwpx)}개 확인')
if A.pptx: print(f'PPT 슬라이드 {check_pptx(A.pptx)}장 확인')
print(f'일치 {OK[0]}칸 / 불일치 {len(ERR)}건 / 대조 못 함 {len(SKIP)}개 / 문구 경고 {len(WARN)}건')
for e in ERR: print('  ✗', e)
for s in SKIP: print('  ? 대조 못 함:', s)
for w in WARN: print('  ! 경고:', w)
sys.exit(1 if ERR else 0)
