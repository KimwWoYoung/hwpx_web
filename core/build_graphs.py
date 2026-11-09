"""엑셀 통계표 → PPT 그래프 자동 생성 (그래프 양식 pptx의 슬라이드를 복제해 값만 교체).

  python3 build_graphs.py --pptx 그래프양식.pptx --excel 통계표.xlsx --start 2-1 --out 결과.pptx
양식 슬라이드 번호가 다르면 --bar/--avg/--scale/--donut 으로 지정 (pptx 스킬 thumbnail.py로 확인).
"""
import glob, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hwtx_common as H
# ======================================================================= 그래프 생성
import argparse, copy, os, re, shutil, subprocess, sys, tempfile, zipfile
from lxml import etree

ap = argparse.ArgumentParser(description='엑셀 통계표 → PPT 그래프(양식 슬라이드 복제) 자동 생성')
ap.add_argument('--pptx', required=True, help='그래프 양식 pptx')
ap.add_argument('--excel', required=True)
ap.add_argument('--out')
ap.add_argument('--start', help='시작 표 번호 (예: 2-1)')
ap.add_argument('--end', help='마지막 표 번호')
ap.add_argument('--bar', type=int, default=1, help='일반 막대 양식 슬라이드 번호')
ap.add_argument('--avg', type=int, default=45, help='평균 배지 막대 양식 슬라이드 번호')
ap.add_argument('--scale', type=int, default=19, help='5점 척도(평균+묶음) 양식 슬라이드 번호')
ap.add_argument('--donut', type=int, default=12, help='2개 보기 도넛 양식 슬라이드 번호')
ap.add_argument('--max-items', type=int, default=10, help='보기 수 초과 시 값 큰 순 상위 N개')
ap.add_argument('--include-scale-detail', action='store_true', help='척도 세부표(x-y-z)도 그래프 생성')
ap.add_argument('--abbr', action='append', default=[], help='라벨 축약 "원문=축약"')
A = ap.parse_args()
for a_ in A.abbr:
    k_, v_ = a_.split('=', 1); H.ABBR[k_] = v_
OUT = A.out or '/mnt/user-data/outputs/' + os.path.splitext(os.path.basename(A.pptx))[0] + '_완성.pptx'

NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def q(tag): p, t = tag.split(':'); return '{%s}%s' % (NS[p], t)
DARK = '155750'

# ---------------------------------------------------------------- 표 선택/유형
tables = H.classify(H.parse_excel(A.excel), scale_full=A.include_scale_detail)
i0 = next(i for i, t in enumerate(tables) if t['num'] == A.start) if A.start else 1
i1 = max(i for i, t in enumerate(tables) if t['num'] == A.end) + 1 if A.end else len(tables)
jobs = []
for t in tables[i0:i1]:
    if t['num'] == '0' or t['type'] == 'scale': continue
    tot = t['rows'][0]
    mean = [tot['vals'][k] for k in t['mean_idx']]
    if t['two']:                                    # 단독 5점 척도: ①~⑤ 막대 + 평균 + 양끝 묶음
        idx = [k for k, c in enumerate(t['cols']) if not c['sub'] and '평균' not in c['top']]
        labels = [re.sub(r'^[①-⑩]\s*', '', t['cols'][k]['top']) for k in idx]
        vals = [tot['vals'][k][0] for k in idx]
        subs = [(c['sub'], tot['vals'][k][0]) for k, c in enumerate(t['cols']) if c['sub']]
        kind = 'scale'; extra = dict(mean=mean[0] if mean else None, unit=t['unit'] or '점', left=subs[0], right=subs[-1])
    else:
        labels = [H.lab(H.iname(t, k)) for k in t['item_idx']]
        vals = [tot['vals'][k][0] for k in t['item_idx']]
        extra = {}
        if mean: kind = 'avg'; extra = dict(mean=mean[0], unit=t['unit'] or '점')
        elif len(vals) == 2 and not t['multi']: kind = 'donut'
        else: kind = 'bar'
        if len(vals) > A.max_items:
            top = sorted(range(len(vals)), key=lambda k: -vals[k])[:A.max_items]
            labels = [labels[k] for k in top]; vals = [vals[k] for k in top]
    jobs.append(dict(num=t['num'], title=t['title'], kind=kind, labels=labels, vals=[H.xround(v, 1) for v in vals], n=tot['n'], **extra))

# ---------------------------------------------------------------- 패키지 열기
W = tempfile.mkdtemp()
zipfile.ZipFile(A.pptx).extractall(W)
def rd(p): return open(os.path.join(W, p), encoding='utf-8').read()
def wr(p, s):
    os.makedirs(os.path.dirname(os.path.join(W, p)), exist_ok=True)
    open(os.path.join(W, p), 'w', encoding='utf-8').write(s)
def rels_of(part):
    d, f = os.path.split(part); return os.path.join(d, '_rels', f + '.rels')
def targets(part):
    s = rd(rels_of(part)); base = os.path.dirname(part)
    return [(m.group(1), m.group(2), os.path.normpath(os.path.join(base, m.group(3))))
            for m in re.finditer(r'Id="(rId\d+)" Type="[^"]*/(\w+)" Target="([^"]+)"', s)]
ct = rd('[Content_Types].xml')
CT = {'slide': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
      'notesSlide': 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
      'chart': 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      'themeOverride': 'application/vnd.openxmlformats-officedocument.themeOverride+xml',
      'chartUserShapes': 'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml'}
new_ct = []

def clone_slide(src_no, k):
    """양식 슬라이드 + 차트 + 임베디드 엑셀 + 노트를 통째로 복제 (k = 새 번호)"""
    sp = f'ppt/slides/slide{src_no}.xml'; dp = f'ppt/slides/slide{k}.xml'
    shutil.copy(os.path.join(W, sp), os.path.join(W, dp))
    srels = rd(rels_of(sp))
    for rid, typ, tgt in targets(sp):
        if typ == 'chart':
            nc = f'ppt/charts/chart{k}.xml'
            shutil.copy(os.path.join(W, tgt), os.path.join(W, nc)); new_ct.append(('/' + nc, CT['chart']))
            crels = rd(rels_of(tgt))
            for rid2, typ2, tgt2 in targets(tgt):
                ext = os.path.splitext(tgt2)[1]
                if typ2 == 'package': nn = f'ppt/embeddings/graph_{k}.xlsx'
                elif typ2 == 'themeOverride': nn = f'ppt/theme/themeOverride{k}.xml'; new_ct.append(('/' + nn, CT['themeOverride']))
                elif typ2 == 'chartUserShapes': nn = f'ppt/drawings/drawing{k}.xml'; new_ct.append(('/' + nn, CT['chartUserShapes']))
                else: continue
                shutil.copy(os.path.join(W, tgt2), os.path.join(W, nn))
                crels = re.sub(r'(Id="%s"[^>]*Target=")[^"]+' % rid2, r'\g<1>../' + nn.split('ppt/')[1].replace('charts/', ''), crels)
            wr(rels_of(nc), crels)
            srels = re.sub(r'(Id="%s"[^>]*Target=")[^"]+' % rid, r'\g<1>../charts/chart%d.xml' % k, srels)
        elif typ == 'notesSlide':
            nn = f'ppt/notesSlides/notesSlide{k}.xml'
            shutil.copy(os.path.join(W, tgt), os.path.join(W, nn)); new_ct.append(('/' + nn, CT['notesSlide']))
            nr = re.sub(r'Target="\.\./slides/slide\d+\.xml"', f'Target="../slides/slide{k}.xml"', rd(rels_of(tgt)))
            wr(rels_of(nn), nr)
            srels = re.sub(r'(Id="%s"[^>]*Target=")[^"]+' % rid, r'\g<1>../notesSlides/notesSlide%d.xml' % k, srels)
    wr(rels_of(dp), srels); new_ct.append(('/' + dp, CT['slide']))
    return dp

# ---------------------------------------------------------------- 편집 함수
def parse(p): return etree.parse(os.path.join(W, p))
def save(tree, p): tree.write(os.path.join(W, p), xml_declaration=True, encoding='UTF-8', standalone=True)
def col(i): s = ''; i += 1
def colname(i):
    s = ''; i += 1
    while i: i, r = divmod(i - 1, 26); s = chr(65 + r) + s
    return s

def set_cache(parent, labels_or_vals, is_num):
    cache = parent.find('.//' + q('c:numCache' if is_num else 'c:strCache'))
    for pt in cache.findall(q('c:pt')): cache.remove(pt)
    cache.find(q('c:ptCount')).set('val', str(len(labels_or_vals)))
    for i, v in enumerate(labels_or_vals):
        pt = etree.SubElement(cache, q('c:pt')); pt.set('idx', str(i))
        etree.SubElement(pt, q('c:v')).text = str(v)
    f = parent.find('.//' + q('c:f'))
    row = 2 if is_num else 1
    f.text = f"Sheet1!$B${row}:${colname(len(labels_or_vals))}${row}"

def edit_chart(cp, labels, vals, kind):
    tree = parse(cp); root = tree.getroot()
    ser = root.find('.//' + q('c:ser'))
    if kind == 'donut':
        # idx0 = 두 번째 보기(연한색, 오른쪽), idx1 = 첫 번째 보기(진한색, 왼쪽)
        labels, vals = [labels[1], labels[0]], [vals[1], vals[0]]
    else:
        dark = None
        base_fill = ser.find(q('c:spPr') + '/' + q('a:solidFill') + '/' + q('a:srgbClr'))
        base_col = base_fill.get('val') if base_fill is not None else None
        for d in ser.findall(q('c:dPt')):
            fill = d.find('.//' + q('a:srgbClr'))
            # [웹앱] 강조색: DARK 이거나, 계열 기본색과 다른 색이면 강조 막대로 봄 (양식마다 색이 달라서)
            if fill is not None and dark is None and (fill.get('val') == DARK or fill.get('val') != base_col): dark = copy.deepcopy(d)
            ser.remove(d)
        dl = ser.find(q('c:dLbls'))
        if dl is not None:
            for x in dl.findall(q('c:dLbl')): dl.remove(x)
        mx = max(range(len(vals)), key=lambda i: vals[i])
        if dark is not None:
            dark.find(q('c:idx')).set('val', str(mx))
            anchor = ser.find(q('c:invertIfNegative'))
            anchor.addnext(dark)
    set_cache(ser.find(q('c:cat')), labels, False)
    set_cache(ser.find(q('c:val')), vals, True)
    save(tree, cp)
    # 임베디드 엑셀
    for rid, typ, tgt in targets(cp):
        if typ == 'package':
            import openpyxl
            wb = openpyxl.load_workbook(os.path.join(W, tgt)); ws = wb.worksheets[0]
            for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, 2), min_col=2, max_col=max(ws.max_column, 2)):
                for c in row: c.value = None
            for i, (l, v) in enumerate(zip(labels, vals)):
                ws.cell(1, 2 + i, l); ws.cell(2, 2 + i, v)
            wb.save(os.path.join(W, tgt))
        if typ == 'chartUserShapes' and kind == 'donut':
            set_runs(tgt, [labels[0], None, H.xfmt(vals[0])])

def set_runs(part, texts):
    """파트 안의 a:t를 순서대로 교체 (None이면 유지)"""
    tree = parse(part); ts = tree.getroot().findall('.//' + q('a:t'))
    for t, v in zip(ts, texts):
        if v is not None: t.text = v
    save(tree, part)

def shape(tree, name):
    for el in tree.getroot().iter(q('p:sp'), q('p:graphicFrame')):
        if el.find('.//' + q('p:cNvPr')).get('name') == name: return el

def para_texts(sp, texts):
    """도형의 문단별 첫 run에 텍스트, 나머지 run 비움"""
    ps = [p for p in sp.findall('.//' + q('a:p')) if p.findall(q('a:r'))]
    for p, txt in zip(ps, texts):
        rs = p.findall(q('a:r'))
        rs[0].find(q('a:t')).text = txt
        for r in rs[1:]: r.find(q('a:t')).text = ''

def edit_label_table(tree, labels):
    gf = shape(tree, '표 8')
    tbl = gf.find('.//' + q('a:tbl')); grid = tbl.find(q('a:tblGrid'))
    gcs = grid.findall(q('a:gridCol')); total = sum(int(g.get('w')) for g in gcs)
    proto_gc = copy.deepcopy(gcs[0])
    for g in gcs: grid.remove(g)
    n = len(labels); w = total // n
    for i in range(n):
        g = copy.deepcopy(proto_gc); g.set('w', str(w if i < n - 1 else total - w * (n - 1)))
        cid = g.find('.//{http://schemas.microsoft.com/office/drawing/2014/main}colId')
        if cid is not None: cid.set('val', str(1000000000 + i * 7919))
        grid.append(g)
    tr = tbl.find(q('a:tr')); tcs = tr.findall(q('a:tc')); proto = copy.deepcopy(tcs[0])
    for tc in tcs: tr.remove(tc)
    ext = tr.find(q('a:extLst'))
    for lab in labels:
        tc = copy.deepcopy(proto)
        body = tc.find(q('a:txBody')); ps = body.findall(q('a:p'))
        p0 = ps[0]
        for p in ps[1:]: body.remove(p)
        rs = p0.findall(q('a:r'))
        for r in rs[1:]: p0.remove(r)
        rs[0].find(q('a:t')).text = lab
        if ext is not None: ext.addprevious(tc)
        else: tr.append(tc)
    gf.find('.//' + q('a:ext')) if False else None
    xf = gf.find(q('p:xfrm')).find(q('a:ext')); xf.set('cx', str(total))

def set_notes(k, text):
    p = f'ppt/notesSlides/notesSlide{k}.xml'
    tree = parse(p)
    for sp in tree.getroot().iter(q('p:sp')):
        ph = sp.find('.//' + q('p:ph'))
        if ph is not None and ph.get('type') == 'body':
            txb = sp.find(q('p:txBody'))
            for pp in txb.findall(q('a:p')): txb.remove(pp)
            pp = etree.SubElement(txb, q('a:p')); r = etree.SubElement(pp, q('a:r'))
            etree.SubElement(r, q('a:rPr'), lang='ko-KR', altLang='en-US')
            etree.SubElement(r, q('a:t')).text = text
    save(tree, p)

# ---------------------------------------------------------------- 실행
BASE = {'bar': A.bar, 'avg': A.avg, 'scale': A.scale, 'donut': A.donut}
_have = {int(m) for m in re.findall(r'slide(\d+)\.xml', ' '.join(os.listdir(os.path.join(W, 'ppt/slides'))))}
for _k, _v in list(BASE.items()):          # [웹앱] 양식에 없는 슬라이드 번호면 막대 양식으로 대체
    if _v not in _have:
        BASE[_k] = A.bar; print(f'(양식에 {_k} 슬라이드({_v}번)가 없어 막대 양식으로 만듦)')
new_slides = []
for j, job in enumerate(jobs):
    k = 1001 + j
    if BASE[job['kind']] == A.bar and job['kind'] != 'bar':
        job['kind'] = 'bar'
    sp = clone_slide(BASE[job['kind']], k)
    cp = next(t for _, ty, t in targets(sp) if ty == 'chart')
    edit_chart(cp, job['labels'], job['vals'], job['kind'])
    tree = parse(sp)
    name = f"<표{job['num']}> {job['title']}"
    for el in tree.getroot().iter(q('p:graphicFrame')):
        nv = el.find('.//' + q('p:cNvPr'))
        if nv.get('name', '').startswith('차트'): nv.set('name', name[:200]); nv.set('descr', name)
    if job['kind'] == 'donut':
        para_texts(shape(tree, '사각형: 둥근 모서리 4'), [job['labels'][0], H.xfmt(job['vals'][0])])
    else:
        edit_label_table(tree, job['labels'])
    if job['kind'] in ('avg', 'scale') and job.get('mean'):
        v, dec = job['mean']
        rect = shape(tree, '모서리가 둥근 직사각형 20')
        rs = rect.findall('.//' + q('a:r'))
        rs[0].find(q('a:t')).text = H.xfmt(v, dec) + ' '; rs[-1].find(q('a:t')).text = job['unit']
        for r in rs[1:-1]: r.find(q('a:t')).text = ''
    if job['kind'] == 'scale':
        para_texts(shape(tree, '직사각형 17'), [job['left'][0], H.xfmt(job['left'][1]) + '%'])
        para_texts(shape(tree, '직사각형 22'), [job['right'][0], H.xfmt(job['right'][1]) + '%'])
    save(tree, sp)
    set_notes(k, name + f"  (n={job['n']:,})")
    new_slides.append(k)
    print(f"{name} | {job['kind']} | {len(job['labels'])}개 | " + ', '.join(f'{l[:12]}={v}' for l, v in zip(job['labels'], job['vals'])))

# Content_Types
for pn, ctype in new_ct:
    if f'PartName="{pn}"' not in ct:
        ct = ct.replace('</Types>', f'<Override PartName="{pn}" ContentType="{ctype}"/></Types>')
wr('[Content_Types].xml', ct)
# presentation: 새 슬라이드만 남김
pr = rd('ppt/_rels/presentation.xml.rels'); pres = rd('ppt/presentation.xml')
ids = []
for j, k in enumerate(new_slides):
    rid = f'rIdG{k}'
    pr = pr.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{k}.xml"/></Relationships>')
    ids.append(f'<p:sldId id="{3000 + j}" r:id="{rid}"/>')
pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', '<p:sldIdLst>' + ''.join(ids) + '</p:sldIdLst>', pres, flags=re.S)
wr('ppt/_rels/presentation.xml.rels', pr); wr('ppt/presentation.xml', pres)
# 고아 파트 정리 (pptx 스킬 clean.py)
# pptx 스킬의 clean.py로 안 쓰는 슬라이드·차트 정리 (설치 위치가 환경마다 달라 찾아서 사용)
CLEAN = next(iter(sorted([os.path.join(os.path.dirname(os.path.abspath(__file__)), 'clean.py')] + glob.glob('/mnt/skills/*/pptx/scripts/clean.py') + glob.glob('/root/.claude/skills/**/pptx/scripts/clean.py', recursive=True))), '')
if CLEAN:
    subprocess.run([sys.executable, CLEAN, W], check=False, stdout=subprocess.DEVNULL)
os.makedirs(os.path.dirname(OUT), exist_ok=True)
if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    order = ['[Content_Types].xml']
    for dp, _, fs in os.walk(W):
        for f in fs:
            rel = os.path.relpath(os.path.join(dp, f), W)
            if rel not in order: order.append(rel)
    for rel in order: z.write(os.path.join(W, rel), rel)
print('slides', len(new_slides), '->', OUT)
