"""보고서 검증: ① 숫자 대조(verify.py) ② 규칙 점검(스크립트) ③ 내용 검토(Claude).

자동 생성본과 사람이 쓴 최종본(hwpx) 모두 같은 방식으로 본다.
페이지(‘전체 결과’ 소제목 단위)와 엑셀 표는 '페이지 제목'으로 짝을 맞추고, 못 맞춘 페이지는 따로 알린다.
"""
import json, os, re, subprocess, sys, zipfile
from lxml import etree

from . import tables as TB
from .tables import H

NS = H.NS; HP = H.HP


def _txt(p): return ''.join(p.xpath('.//hp:t/text()', namespaces=NS)).strip()


def norm(s):
    s = re.sub(r'^\s*(<\s*표\s*[\d\-]+\s*>|\[\s*(그림|표)\s*[\d\-]+\s*\]|(그림|표)\s*[\d\-]+\.?)\s*', '', str(s))
    return re.sub(r'\s+', '', s)


# ---------------------------------------------------------------------- 한글에서 페이지 꺼내기
def extract_pages(hwpx):
    root = etree.fromstring(zipfile.ZipFile(hwpx).read('Contents/section0.xml'))
    P = list(root)
    heads = [i for i, p in enumerate(P) if _txt(p).endswith('전체 결과') and p.find('.//' + HP + 'tbl') is None]
    def title_at(hi):
        j = hi - 1
        while j >= 0 and not _txt(P[j]): j -= 1
        return j
    tis = [title_at(h) for h in heads]
    pages = []
    for n, hi in enumerate(heads):
        end = tis[n + 1] if n + 1 < len(heads) else len(P)
        j = tis[n]
        title = _txt(P[j]) if j >= 0 else ''
        pg = dict(page=n + 1, title=title, overall=[], group=[], captions=[], base='', unit='', has_table=False, has_graph=False)
        mode = 'overall'
        for k in range(hi + 1, end):
            p = P[k]; s = _txt(p); tbl = p.find('.//' + HP + 'tbl')
            if tbl is not None:
                first = ''.join(tbl.find(HP + 'tr').xpath('.//hp:t/text()', namespaces=NS)).strip()
                if first.startswith('(Base'):
                    pg['base'] = first; pg['has_graph'] = tbl.find('.//' + HP + 'pic') is not None
                elif '사례수' in s:
                    pg['has_table'] = True
                continue
            if not s: continue
            if s.endswith('특성별 결과'): mode = 'group'; continue
            if s.startswith('(단위'): pg['unit'] = s; continue
            if mode == 'overall':
                if norm(s) == norm(title) or s.startswith('[그림') or s.startswith('<그림'):
                    pg['captions'].append(s); mode = 'after_cap'; continue
                pg['overall'].append(s)
            elif mode == 'group':
                if s.startswith('('):
                    pg['group'].append(s)
                else:
                    pg['captions'].append(s)
            else:
                pg['captions'].append(s)
        pages.append(pg)
    return pages


def match(pages, tables):
    """페이지 제목 ↔ 엑셀 표(척도 세부표 제외). 같은 제목이 여럿이면 순서대로"""
    by = {}
    for t in tables:
        if t['type'] == 'scale': continue
        for key in {norm(t['title']), norm(t['name'])}:
            by.setdefault(key, []).append(t)
    used = set()
    for pg in pages:
        cands = [t for t in by.get(norm(pg['title']), []) if t['idx'] not in used]
        pg['table'] = cands[0]['idx'] if cands else None
        if cands: used.add(cands[0]['idx'])
    return pages


# ---------------------------------------------------------------------- 규칙 점검 (스크립트)
STYLE_BAD = [(re.compile(r'[.。]\s*$'), '문장 끝 마침표 (개조식은 마침표 없음)'),
             (re.compile(r'(습니다|이다|였다|했다|된다|있다|없다|한다)\s*$'), '서술형 종결 (~임/~나타남 형태로)'),
             (re.compile(r'→'), '화살표 사용'),
             (re.compile(r'(매우|압도적으로|월등히)'), '주관적 수식어')]


def local_checks(pg, t, groups, min_n, abbr, all_titles=None):
    out = []
    def add(where, sent, prob, sev='확인', fix=''):
        out.append(dict(위치=where, 문장=sent[:120], 문제=prob, 근거='', 수정제안=fix, 심각도=sev, 출처='규칙'))
    for where, lines in (('가. 전체 결과', pg['overall']), ('나. 특성별 결과', pg['group'])):
        for s in lines:
            for rx, msg in STYLE_BAD:
                if rx.search(s): add(where, s, msg)
    if t is None:
        return out
    for b in TB.check_numbers(pg['overall'] + pg['group'], t, abbr):
        add('문구 수치', b.split(' ← ', 1)[-1], f'표에 없는 수치 {b.split(" ← ")[0]}', '오류')
    small = [r for r in t['rows'][1:] if r['n'] < min_n]
    for s in pg['group']:
        for r in small:
            lab = H.glab(r['label'])
            if re.search(r'(?<![\w가-힣])' + re.escape(lab) + r'(?![\w가-힣])', s):
                add('나. 특성별 결과', s, f"사례수 {r['n']}명인 '{lab}' 언급 (기준 {min_n}명 미만은 제외)", '오류')
    heads = [re.match(r'\(([^),]+)', s).group(1) for s in pg['group'] if re.match(r'\(([^),]+)', s)]
    present = [g for g in groups if any(r['grp'] == g and r['n'] >= min_n for r in t['rows'][1:])]
    miss = [g for g in present if g not in heads]
    if miss and pg['group']:
        add('나. 특성별 결과', '', f"설정한 구분 중 문단이 없는 것: {', '.join(miss)}", '확인')
    other = [h for h in heads if h not in {r['grp'] for r in t['rows']}]
    if other:
        add('나. 특성별 결과', '', f"표에 없는 구분 이름: {', '.join(other)} (다른 조사 문구가 남았을 수 있음)", '오류')
    if pg['base'] and t['base'] and t['base'] not in pg['base'] and t['raw_base'] not in pg['base']:
        add('그래프 Base', pg['base'], f"Base 문구가 표와 다름 (엑셀: {t['base']})", '확인')
    known = {norm(x) for x in (all_titles or [])}
    for c in pg['captions'][:2]:                          # 그림 캡션, 통계표 캡션
        if norm(c) != norm(pg['title']) and norm(c) not in known and len(c) < 60 and not c.startswith('('):
            add('캡션', c, f"캡션이 소제목('{pg['title']}')과 다름", '확인')
    if not pg['has_graph']:
        add('그래프', '', '그래프 박스에 그림이 없음', '확인')
    return out


# ---------------------------------------------------------------------- 숫자 대조 (verify.py)
def run_verify(excel, hwpx=None, pptx=None, title=None, base=None, abbr=None):
    cmd = [sys.executable, os.path.join(TB.CORE, 'verify.py'), '--excel', excel]
    if hwpx: cmd += ['--hwpx', hwpx]
    if pptx: cmd += ['--pptx', pptx]
    for flag, m in (('--title', title), ('--base', base), ('--abbr', abbr)):
        for k, v in (m or {}).items(): cmd += [flag, f'{k}={v}']
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    log = r.stdout + r.stderr
    rows = []
    for line in log.splitlines():
        s = line.strip()
        if s.startswith('✗'): rows.append(dict(심각도='오류', 내용=s[1:].strip()))
        elif s.startswith('!'): rows.append(dict(심각도='경고', 내용=re.sub(r'^!\s*경고:\s*', '', s)))
        elif s.startswith('?'): rows.append(dict(심각도='대조 못 함', 내용=s[1:].strip()))
    summary = next((l for l in log.splitlines() if l.startswith('일치')), '')
    return dict(ok=r.returncode == 0, summary=summary, rows=rows, log=log)


# ---------------------------------------------------------------------- 결과 엑셀
def to_xlsx(path, numeric, pages, issues):
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = '내용 검토'
    head = ['페이지', '제목', '표 번호', '위치', '심각도', '문제', '근거', '문장', '수정 제안', '출처']
    ws.append(head)
    num = {pg['page']: pg for pg in pages}
    for it in issues:
        pg = num.get(it.get('page'), {})
        ws.append([it.get('page'), pg.get('title', ''), it.get('num', ''), it.get('위치', ''), it.get('심각도', ''), it.get('문제', ''),
                   it.get('근거', ''), it.get('문장', ''), it.get('수정제안', ''), it.get('출처', '')])
    ws2 = wb.create_sheet('숫자 대조')
    ws2.append(['요약', numeric.get('summary', '')]); ws2.append([])
    ws2.append(['심각도', '내용'])
    for r in numeric.get('rows', []): ws2.append([r['심각도'], r['내용']])
    for w in (ws, ws2):
        for c in w[1 if w is ws else 3]: c.font = Font(bold=True); c.fill = PatternFill('solid', fgColor='E8E5E4')
        for col, width in zip('ABCDEFGHIJ', [7, 28, 9, 14, 9, 45, 35, 50, 50, 8] if w is ws else [12, 120]):
            w.column_dimensions[col].width = width
        for row in w.iter_rows(min_row=2):
            for c in row: c.alignment = Alignment(wrap_text=True, vertical='top')
    wb.save(path)
