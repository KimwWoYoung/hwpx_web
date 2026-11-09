"""[웹앱] 그래프 PPT → 300dpi 그림(기본) 또는 EMF(메타파일) → 한글 보고서의 그래프 박스에 삽입.

  python3 insert_graphs.py --hwpx 보고서.hwpx --pptx 그래프.pptx --out 결과.hwpx [--emf-zip 슬라이드들.zip] [--replace-all]

- 슬라이드와 그래프 박스는 '표 제목'으로 짝을 맞춘다
  · 슬라이드: 노트의 '<표2-1> 제목 (n=…)' (없으면 차트 이름)
  · 한글: 그래프 박스(첫 칸이 '(Base: …' 인 표) 바로 앞 문단(그림 캡션)
- EMF 만들기
  · 기본: LibreOffice로 슬라이드마다 EMF 변환 (PowerPoint와 글꼴·모양이 약간 다를 수 있음)
  · --emf-zip: PowerPoint '다른 이름으로 저장 → 확장 메타파일(*.emf) → 모든 슬라이드'로 만든 파일 묶음.
    파일 이름의 숫자(슬라이드1.emf, Slide1.emf …)를 슬라이드 순서로 본다
- 이 스크립트가 넣은 그림은 다시 실행하면 교체된다. 사람이 직접 붙인 그림이 있는 박스는 건너뛴다 (--replace-all 이면 교체)
"""
import argparse, os, re, shutil, struct, subprocess, sys, tempfile, zipfile
from lxml import etree
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hwtx_common as H
from hwtx_common import HP

HC = '{http://www.hancom.co.kr/hwpml/2011/core}'
MARK = 'hwtx-graph'
PNS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
       'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
       'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}


def norm(s):
    s = re.sub(r'^\s*(<\s*표\s*[\d\-]+\s*>|\[\s*(그림|표)\s*[\d\-]+\s*\]|(그림|표)\s*[\d\-]+\.?)\s*', '', str(s))
    s = re.sub(r'\s*\(n\s*=\s*[\d,]+\).*$', '', s)
    return re.sub(r'\s+', '', s)


def table_num(s):
    m = re.search(r'<\s*표\s*([\d\-]+)\s*>', str(s)); return m.group(1) if m else None


# ---------------------------------------------------------------- PPT 읽기
def read_slides(pptx):
    """[(순서, 슬라이드 파트, 제목문자열)]"""
    z = zipfile.ZipFile(pptx)
    pres = etree.fromstring(z.read('ppt/presentation.xml'))
    rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    rmap = {r.get('Id'): 'ppt/' + r.get('Target').lstrip('/').replace('ppt/', '') for r in rels}
    out = []
    for i, sid in enumerate(pres.find('p:sldIdLst', PNS)):
        part = rmap[sid.get('{%s}id' % PNS['r'])]
        title = ''
        rp = os.path.dirname(part) + '/_rels/' + os.path.basename(part) + '.rels'
        if rp in z.namelist():
            for r in etree.fromstring(z.read(rp)):
                if r.get('Type', '').endswith('/notesSlide'):
                    np_ = os.path.normpath(os.path.join(os.path.dirname(part), r.get('Target'))).replace('\\', '/')
                    title = ''.join(etree.fromstring(z.read(np_)).xpath('.//a:t/text()', namespaces=PNS)).strip()
        if '<표' not in title:
            s = etree.fromstring(z.read(part))
            for nv in s.iter('{%s}cNvPr' % PNS['p']):
                if '<표' in (nv.get('name', '') + nv.get('descr', '')):
                    title = nv.get('descr') or nv.get('name'); break
        out.append((i + 1, part, title))
    return out


def uno_convert(pptx, work, fmt='emf'):
    """도형 범위만 EMF로 (pptx_to_emf.py, LibreOffice UNO). 실패하면 None"""
    out = os.path.join(work, 'u' + fmt); os.makedirs(out, exist_ok=True)
    here = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pptx_to_emf.py')
    for py in dict.fromkeys([sys.executable, '/usr/bin/python3']):
        if not os.path.exists(py): continue
        try:
            r = subprocess.run([py, here, pptx, out, fmt], capture_output=True, text=True, timeout=900)
        except Exception:
            continue
        got = {int(f[:3]): os.path.join(out, f) for f in os.listdir(out) if f.endswith('.' + fmt)}
        if r.returncode == 0 and got:
            return got
    return None


def split_and_convert(pptx, work):
    """슬라이드마다 1장짜리 pptx를 만들어 LibreOffice로 한 번에 EMF 변환 → {순서: emf경로}"""
    z = zipfile.ZipFile(pptx)
    pres = z.read('ppt/presentation.xml').decode('utf-8')
    os.makedirs(work, exist_ok=True)
    ids = re.findall(r'<p:sldId [^>]*/>', re.search(r'<p:sldIdLst>(.*?)</p:sldIdLst>', pres, re.S).group(1))
    files = []
    for i, sid in enumerate(ids, 1):
        one = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', '<p:sldIdLst>' + sid + '</p:sldIdLst>', pres, flags=re.S)
        fp = os.path.join(work, f's{i:03d}.pptx')
        with zipfile.ZipFile(fp, 'w', zipfile.ZIP_DEFLATED) as o:
            for info in z.infolist():
                data = one.encode('utf-8') if info.filename == 'ppt/presentation.xml' else z.read(info.filename)
                o.writestr(zipfile.ZipInfo(info.filename, date_time=info.date_time), data, compress_type=zipfile.ZIP_DEFLATED)
        files.append(fp)
    outdir = os.path.join(work, 'emf'); os.makedirs(outdir, exist_ok=True)
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:
        sys.exit('LibreOffice(soffice)가 없어 EMF로 바꿀 수 없습니다. --emf-zip 으로 PowerPoint에서 만든 EMF를 넣어 주세요.')
    env = dict(os.environ, HOME=work)                    # 동시 실행 시 프로필 충돌 방지
    for k in range(0, len(files), 40):
        subprocess.run([soffice, '--headless', f'-env:UserInstallation=file://{work}/lo', '--convert-to', 'emf',
                        '--outdir', outdir] + files[k:k + 40], check=False, env=env,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
    return {i: os.path.join(outdir, f's{i:03d}.emf') for i in range(1, len(files) + 1)
            if os.path.exists(os.path.join(outdir, f's{i:03d}.emf'))}


def from_zip(zp, work):
    d = os.path.join(work, 'emfzip'); os.makedirs(d, exist_ok=True)
    out = {}
    with zipfile.ZipFile(zp) as z:
        for info in z.infolist():
            name = info.filename
            try: name = name.encode('cp437').decode('cp949')       # 윈도우 압축 한글 파일명
            except Exception: pass
            if not name.lower().endswith('.emf'): continue
            m = re.findall(r'(\d+)', os.path.basename(name))
            if not m: continue
            fp = os.path.join(d, f'{int(m[-1]):03d}.emf'); open(fp, 'wb').write(z.read(info))
            out[int(m[-1])] = fp
    return out


def emf_size(fp):
    """EMF 헤더 rclFrame (0.01mm) → (가로, 세로) mm"""
    b = open(fp, 'rb').read(64)
    l, t, r, btm = struct.unpack('<4i', b[24:40])
    return (r - l) / 100, (btm - t) / 100


# ---------------------------------------------------------------- 한글에 넣기
def graph_boxes(D):
    """[(문단 인덱스, 박스 문단, 그림 칸 문단, 그림 칸 tc, 캡션 제목)]"""
    P = D.paras(); out = []
    for i, p in enumerate(P):
        tbl = p.find('.//' + HP + 'tbl')
        if tbl is None: continue
        trs = tbl.findall(HP + 'tr')
        if len(trs) < 2: continue
        first = ''.join(trs[0].xpath('.//hp:t/text()', namespaces=H.NS)).strip()
        if not first.startswith('(Base'): continue
        tc = trs[1].find(HP + 'tc'); cp = tc.find('.//' + HP + 'p')
        j = i - 1
        while j >= 0 and not D.text(P[j]): j -= 1
        out.append((i, p, cp, tc, tbl, D.text(P[j]) if j >= 0 else ''))
    return out


PIC = ('<hp:pic xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
       'id="{id}" zOrder="{z}" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" '
       'href="" groupLevel="0" instid="{inst}" reverse="0"><hp:offset x="0" y="0"/><hp:orgSz width="{ow}" height="{oh}"/>'
       '<hp:curSz width="{w}" height="{h}"/><hp:flip horizontal="0" vertical="0"/>'
       '<hp:rotationInfo angle="0" centerX="{cx}" centerY="{cy}" rotateimage="1"/><hp:renderingInfo>'
       '<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:scaMatrix e1="{sx}" e2="0" e3="0" e4="0" e5="{sy}" e6="0"/>'
       '<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo><hp:imgRect><hc:pt0 x="0" y="0"/>'
       '<hc:pt1 x="{ow}" y="0"/><hc:pt2 x="{ow}" y="{oh}"/><hc:pt3 x="0" y="{oh}"/></hp:imgRect>'
       '<hp:imgClip left="0" right="{dw}" top="0" bottom="{dh}"/><hp:inMargin left="0" right="0" top="0" bottom="0"/>'
       '<hp:imgDim dimwidth="{dw}" dimheight="{dh}"/><hc:img binaryItemIDRef="{bid}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>'
       '<hp:effects/><hp:sz width="{w}" widthRelTo="ABSOLUTE" height="{h}" heightRelTo="ABSOLUTE" protect="0"/>'
       '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" '
       'horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
       '<hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:shapeComment>{mark}</hp:shapeComment></hp:pic>')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--hwpx', required=True); ap.add_argument('--pptx', required=True); ap.add_argument('--out', required=True)
    ap.add_argument('--emf-zip'); ap.add_argument('--replace-all', action='store_true')
    ap.add_argument('--format', default='png', choices=['emf', 'png'], help='그림 형식: png(300dpi 그림, 기본) / emf(메타파일)')
    A = ap.parse_args()
    work = tempfile.mkdtemp()
    slides = read_slides(A.pptx)
    if A.format == 'png':
        emfs = uno_convert(A.pptx, work, 'png'); how = 'LibreOffice PNG 300dpi'
        if not emfs: sys.exit('PNG 변환 실패')
    elif A.emf_zip:
        emfs = from_zip(A.emf_zip, work); how = 'PowerPoint EMF'
    else:
        emfs = uno_convert(A.pptx, work); how = 'LibreOffice(도형 범위)'
        if not emfs:
            emfs = split_and_convert(A.pptx, os.path.join(work, 'lo')); how = 'LibreOffice(슬라이드 전체, 예비 방식)'
    print(f'EMF 만들기: {how}')
    by_title = {}
    for no, part, title in slides:
        if no in emfs and title: by_title.setdefault(norm(title), []).append((no, title))
    D = H.Doc(A.hwpx, os.path.join(work, 'hwpx'))
    D.extra = []
    hpf_path = os.path.join(D.dir, 'Contents/content.hpf'); hpf = open(hpf_path, encoding='utf-8').read()
    used, done, skipped, missing = set(), [], [], []
    seq = 0
    for i, p, cp, tc, tbl, cap in graph_boxes(D):
        cands = [c for c in by_title.get(norm(cap), []) if c[0] not in used]
        if not cands:
            missing.append(cap); continue
        no, title = cands[0]
        old = cp.findall('.//' + HP + 'pic')
        if old and not A.replace_all and not all(MARK in ''.join(o.xpath('.//hp:shapeComment/text()', namespaces=H.NS)) for o in old):
            skipped.append(cap); continue
        run = cp.find(HP + 'run')
        for o in old: o.getparent().remove(o)
        used.add(no); seq += 1
        # 그림 크기: 칸 폭(여백 제외)에 맞추고 세로는 비율대로
        cw = int(tc.find(HP + 'cellSz').get('width')); mg = tc.find(HP + 'cellMargin')
        inner = cw - int(mg.get('left')) - int(mg.get('right')) - 200
        src_img = emfs[no]; ext = os.path.splitext(src_img)[1].lower().lstrip('.')
        if ext == 'png':
            from PIL import Image
            im = Image.open(src_img)
            if im.mode in ('RGBA', 'LA', 'P'):              # 투명 배경 → 흰 배경
                bg = Image.new('RGB', im.size, 'white'); bg.paste(im.convert('RGBA'), mask=im.convert('RGBA').split()[-1]); bg.save(src_img)
            px, py_ = im.size
            ow, oh = px * 24, py_ * 24                     # 300dpi 기준 원래 크기 (HWPUNIT = 1/7200인치)
            dw, dh = px * 75, py_ * 75                     # 한글이 쓰는 그림 픽셀 단위(96dpi 기준)
        else:
            wmm, hmm = emf_size(src_img)
            ow, oh = int(wmm * 7200 / 25.4), int(hmm * 7200 / 25.4)
            dw, dh = ow, oh
        w = min(inner, ow) if ext == 'emf' and ow <= inner else inner
        h = int(w * oh / ow)
        bid = f'hwtxgraph{seq}'; fn = f'BinData/{bid}.{ext}'
        shutil.copy(src_img, os.path.join(D.dir, fn)); D.extra.append(fn)
        hpf = re.sub(r'<opf:item id="%s"[^>]*/>' % bid, '', hpf)
        hpf = hpf.replace('</opf:manifest>', f'<opf:item id="{bid}" href="{fn}" media-type="{'image/png' if ext == 'png' else 'image/x-emf'}" isEmbeded="1"/></opf:manifest>')
        pic = etree.fromstring(PIC.format(id=1300000000 + seq, z=200 + seq, inst=1400000000 + seq, w=w, h=h, cx=w // 2, cy=h // 2,
                                          ow=ow, oh=oh, dw=dw, dh=dh, sx=f'{w / ow:.6f}', sy=f'{h / oh:.6f}',
                                          bid=bid, mark=f'{MARK} {table_num(title) or ""} slide{no}'))
        if run is None:
            run = etree.SubElement(cp, HP + 'run'); run.set('charPrIDRef', '0'); cp.insert(0, run)
        run.insert(0, pic)
        # 칸·표 높이를 그림에 맞춤
        newh = h + int(mg.get('top')) + int(mg.get('bottom')) + 300
        oldh = int(tc.find(HP + 'cellSz').get('height'))
        tc.find(HP + 'cellSz').set('height', str(newh))
        sz = tbl.find(HP + 'sz'); sz.set('height', str(int(sz.get('height')) - oldh + newh))
        seg = cp.find('.//' + HP + 'lineseg')
        if seg is not None:
            seg.set('vertsize', str(h)); seg.set('textheight', str(h)); seg.set('baseline', str(int(h * 0.85)))
        oseg = p.find(HP + 'linesegarray/' + HP + 'lineseg')
        th = int(sz.get('height'))
        if oseg is not None:
            oseg.set('vertsize', str(th)); oseg.set('textheight', str(th)); oseg.set('baseline', str(int(th * 0.85)))
        done.append((cap, no))
    open(hpf_path, 'w', encoding='utf-8').write(hpf)
    D.save(A.out)
    for cap, no in done: print(f'✓ {cap}  ← 슬라이드 {no}')
    for cap in skipped: print(f'· 건너뜀 (직접 붙인 그림 있음): {cap}')
    for cap in missing: print(f'? 맞는 슬라이드 없음: {cap}')
    left = [f'{no}:{t}' for no, _, t in slides if no not in used]
    if left: print('? 쓰이지 않은 슬라이드: ' + ' / '.join(left))
    print(f'그래프 {len(done)}개 삽입 / 건너뜀 {len(skipped)} / 슬라이드 없음 {len(missing)} → {A.out}')


if __name__ == '__main__':
    main()
