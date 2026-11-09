"""[웹앱] 그래프 PPT의 슬라이드마다 '도형만' EMF로 내보내기 (LibreOffice UNO).

PowerPoint에서 차트를 복사해 '그림(메타파일)'로 붙이는 것과 같은 결과를 목표로 한다.
- 슬라이드 전체가 아니라 슬라이드 위 도형(차트·라벨 표·배지)의 실제 범위만 내보냄
  → 슬라이드 밖으로 넘친 라벨이 잘리지 않고, 빈 여백도 없음
- LibreOffice가 기본으로 켜는 '한글-숫자/영문 사이 자동 띄우기'를 끔 (1년 → '1 년' 방지)

  python3 pptx_to_emf.py 그래프.pptx 출력폴더      → 출력폴더/001.emf, 002.emf …
시스템 파이썬에 uno 모듈(python3-uno)이 있어야 한다.
"""
import os, subprocess, sys, tempfile, time, uuid

import uno
from com.sun.star.beans import PropertyValue


def pv(name, value):
    p = PropertyValue(); p.Name = name; p.Value = value; return p


def connect(profile):
    pipe = 'hwtx' + uuid.uuid4().hex[:8]
    soffice = 'soffice'
    proc = subprocess.Popen([soffice, '--headless', '--invisible', '--nologo', '--norestore', '--nodefault',
                             f'-env:UserInstallation=file://{profile}', f'--accept=pipe,name={pipe};urp;'],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    for _ in range(120):
        try:
            ctx = resolver.resolve(f'uno:pipe,name={pipe};urp;StarOffice.ComponentContext')
            return proc, ctx
        except Exception:
            time.sleep(0.5)
    proc.kill(); raise RuntimeError('LibreOffice에 연결하지 못했습니다')


def no_autospace(text):
    try:
        en = text.createEnumeration()
        while en.hasMoreElements():
            para = en.nextElement()
            try: para.setPropertyValue('ParaIsCharacterDistance', False)
            except Exception: pass
    except Exception:
        pass


def fix_shape(shape):
    st = shape.getShapeType()
    if st == 'com.sun.star.drawing.TableShape':
        model = shape.getPropertyValue('Model')
        for r in range(model.getRows().getCount()):
            for c in range(model.getColumns().getCount()):
                no_autospace(model.getCellByPosition(c, r).getText())
    elif st == 'com.sun.star.drawing.GroupShape':
        for i in range(shape.getCount()): fix_shape(shape.getByIndex(i))
    elif hasattr(shape, 'getText'):
        no_autospace(shape.getText())


def main(src, outdir, fmt='emf', dpi=300):
    os.makedirs(outdir, exist_ok=True)
    profile = tempfile.mkdtemp(prefix='lo_')
    proc, ctx = connect(profile)
    try:
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)), '_blank', 0, (pv('Hidden', True),))
        pages = doc.getDrawPages()
        for i in range(pages.getCount()):
            page = pages.getByIndex(i)
            coll = smgr.createInstanceWithContext('com.sun.star.drawing.ShapeCollection', ctx)
            for k in range(page.getCount()):
                sh = page.getByIndex(k)
                fix_shape(sh)
                coll.add(sh)
            if coll.getCount() == 0: continue
            flt = smgr.createInstanceWithContext('com.sun.star.drawing.GraphicExportFilter', ctx)
            flt.setSourceDocument(coll)
            out = os.path.join(outdir, f'{i + 1:03d}.{fmt}')
            props = [pv('URL', uno.systemPathToFileUrl(os.path.abspath(out))), pv('MediaType', 'image/x-emf' if fmt == 'emf' else 'image/png')]
            if fmt == 'png':                                   # 고해상도 그림 (도형 범위 크기 × dpi)
                bw = max(sh.getPosition().X + sh.getSize().Width for sh in (coll.getByIndex(j) for j in range(coll.getCount()))) - \
                     min(sh.getPosition().X for sh in (coll.getByIndex(j) for j in range(coll.getCount())))
                bh = max(sh.getPosition().Y + sh.getSize().Height for sh in (coll.getByIndex(j) for j in range(coll.getCount()))) - \
                     min(sh.getPosition().Y for sh in (coll.getByIndex(j) for j in range(coll.getCount())))
                fd = uno.Any('[]com.sun.star.beans.PropertyValue', (pv('PixelWidth', int(bw / 2540 * dpi)), pv('PixelHeight', int(bh / 2540 * dpi))))
                props.append(pv('FilterData', fd))
            uno.invoke(flt, 'filter', (tuple(props),))
            print(out)
        doc.close(True)
    finally:
        try: desktop.terminate()
        except Exception: pass
        time.sleep(0.5); proc.kill()


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], *(sys.argv[3:4] or ['emf']))
