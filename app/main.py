"""보고서 작업 웹앱 (FastAPI).

실행: uvicorn app.main:app --host 0.0.0.0 --port 8000
환경변수
  ANTHROPIC_API_KEY  Claude API 키 (서버에만 둠)
  APP_PASSWORD       팀 공용 접속 비밀번호 (비우면 비밀번호 없음)
  HWTX_MODEL         문구·검토에 쓸 모델 (기본 claude-opus-5-5)
  DATA_DIR           작업 파일 보관 위치 (기본 /tmp/hwtx-sessions, 3일 지나면 삭제)
"""
import json, os, re, shutil, subprocess, sys, time, uuid, zipfile
from typing import Optional

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from pydantic import BaseModel

from . import ai, review as RV, tables as TB

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.environ.get('DATA_DIR', '/tmp/hwtx-sessions')
os.makedirs(DATA, exist_ok=True)
KEEP_DAYS = 3
app = FastAPI(title='보고서 작업 도우미')

FILES = {  # 종류: (저장 이름, 허용 확장자)
    'excel': ('stat.xlsx', ('.xlsx',)),
    'template': ('template.hwpx', ('.hwpx', '.hwtx')),
    'questionnaire': ('questionnaire', ('.pdf', '.docx', '.hwpx', '.txt')),
    'graph_template': ('graph_template.pptx', ('.pptx',)),
    'graphs_edited': ('graphs_edited.pptx', ('.pptx',)),
    'emf_zip': ('emf.zip', ('.zip',)),
    'edited_hwpx': ('edited.hwpx', ('.hwpx', '.hwtx')),
    'review_hwpx': ('review.hwpx', ('.hwpx', '.hwtx')),
}


# ---------------------------------------------------------------------- 공통
def auth(x_app_password: Optional[str] = Header(default=None)):
    pw = os.environ.get('APP_PASSWORD')
    if pw and x_app_password != pw:
        raise HTTPException(401, '비밀번호가 맞지 않습니다')


def cleanup():
    now = time.time()
    for d in os.listdir(DATA):
        p = os.path.join(DATA, d)
        if os.path.isdir(p) and now - os.path.getmtime(p) > KEEP_DAYS * 86400:
            shutil.rmtree(p, ignore_errors=True)


def sdir(sid):
    if not re.fullmatch(r'[0-9a-f]{32}', sid or ''): raise HTTPException(400, '잘못된 작업 번호')
    d = os.path.join(DATA, sid)
    if not os.path.isdir(d): raise HTTPException(404, '작업이 만료되었습니다. 파일을 다시 올려 주세요')
    return d


def state(sid):
    p = os.path.join(sdir(sid), 'state.json')
    return json.load(open(p, encoding='utf-8')) if os.path.exists(p) else {}


def save_state(sid, st):
    json.dump(st, open(os.path.join(sdir(sid), 'state.json'), 'w', encoding='utf-8'), ensure_ascii=False)


def fpath(sid, kind):
    st = state(sid); name = st.get('files', {}).get(kind)
    return os.path.join(sdir(sid), name) if name else None


def need(sid, kind, label):
    p = fpath(sid, kind)
    if not p or not os.path.exists(p): raise HTTPException(400, f'{label} 파일을 먼저 올려 주세요')
    return p


DEFAULT = dict(sample=None, start=None, end=None, groups=[], min_n=10, abbr={}, base={}, title={},
               fill_sample=False, scale_full=False)


def settings(st):
    s = dict(DEFAULT); s.update(st.get('settings', {})); return s


def load_tables(sid):
    st = state(sid); s = settings(st)
    tabs, groups = TB.parse(need(sid, 'excel', '통계표 엑셀'), s['abbr'], s['base'], s['title'], s['scale_full'])
    return tabs, groups, s


def run(cmd, timeout=900):
    r = subprocess.run([sys.executable] + cmd, capture_output=True, text=True, timeout=timeout)
    return r.returncode, (r.stdout + ('\n' + r.stderr if r.stderr.strip() else '')).strip()


def map_args(flag, m):
    out = []
    for k, v in (m or {}).items():
        if k and v is not None and str(v).strip(): out += [flag, f'{k}={v}']
    return out


def out_name(sid, suffix):
    st = state(sid); base = os.path.splitext(st.get('names', {}).get('template') or st.get('names', {}).get('excel') or '보고서')[0]
    return f'{base}_{suffix}'


# ---------------------------------------------------------------------- 화면
@app.get('/', response_class=HTMLResponse)
def index():
    return open(os.path.join(ROOT, 'static', 'simple.html'), encoding='utf-8').read()


@app.get('/pro', response_class=HTMLResponse)
def index_pro():
    return open(os.path.join(ROOT, 'static', 'index.html'), encoding='utf-8').read()


@app.get('/api/config')
def config():
    return dict(ai=ai.enabled(), model=ai.MODEL, password=bool(os.environ.get('APP_PASSWORD')),
                fake=os.environ.get('HWTX_FAKE_AI') == '1')


@app.post('/api/login', dependencies=[Depends(auth)])
def login():
    return dict(ok=True)


# ---------------------------------------------------------------------- 작업 만들기·파일
@app.post('/api/session', dependencies=[Depends(auth)])
def new_session():
    cleanup()
    sid = uuid.uuid4().hex; os.makedirs(os.path.join(DATA, sid))
    save_state(sid, dict(files={}, names={}, settings={}, texts={}))
    return dict(sid=sid)


@app.post('/api/s/{sid}/file/{kind}', dependencies=[Depends(auth)])
async def upload(sid: str, kind: str, file: UploadFile = File(...)):
    if kind not in FILES: raise HTTPException(400, '알 수 없는 파일 종류')
    name, exts = FILES[kind]
    ext = os.path.splitext(file.filename or '')[1].lower()
    if ext not in exts: raise HTTPException(400, f'{", ".join(exts)} 파일만 올릴 수 있습니다')
    if kind == 'questionnaire': name += ext
    if kind == 'template' and ext == '.hwtx': name = 'template.hwpx'
    dest = os.path.join(sdir(sid), name)
    with open(dest, 'wb') as f: shutil.copyfileobj(file.file, f)
    if ext in ('.xlsx', '.hwpx', '.hwtx', '.pptx', '.docx', '.zip') and not zipfile.is_zipfile(dest):
        os.remove(dest); raise HTTPException(400, '파일이 손상되었거나 형식이 다릅니다')
    st = state(sid); st['files'][kind] = name; st['names'][kind] = file.filename; save_state(sid, st)
    out = dict(ok=True, kind=kind, name=file.filename)
    if kind == 'excel':
        out.update(overview(sid))
    return out


def overview(sid):
    tabs, groups, s = load_tables(sid)
    bases = list(dict.fromkeys(t['raw_base'] for t in tabs))
    return dict(tables=[TB.summary(t) for t in tabs], groups=groups, bases=bases, settings=s)


@app.get('/api/s/{sid}', dependencies=[Depends(auth)])
def get_session(sid: str):
    st = state(sid); out = dict(files=st.get('names', {}), settings=settings(st), texts=st.get('texts', {}),
                                outputs=st.get('outputs', {}))
    if fpath(sid, 'excel'): out.update(overview(sid))
    return out


class Settings(BaseModel):
    sample: Optional[str] = None
    start: Optional[str] = None
    end: Optional[str] = None
    groups: list = []
    min_n: int = 10
    abbr: dict = {}
    base: dict = {}
    title: dict = {}
    fill_sample: bool = False
    scale_full: bool = False


@app.post('/api/s/{sid}/settings', dependencies=[Depends(auth)])
def set_settings(sid: str, body: Settings):
    st = state(sid); st['settings'] = body.model_dump(); save_state(sid, st)
    return overview(sid)


# ---------------------------------------------------------------------- Base 풀어쓰기 (AI)
def questionnaire_text(p):
    ext = os.path.splitext(p)[1].lower()
    if ext == '.pdf':
        from pypdf import PdfReader
        return '\n'.join((pg.extract_text() or '') for pg in PdfReader(p).pages)
    if ext == '.docx':
        import docx
        d = docx.Document(p)
        out = [x.text for x in d.paragraphs]
        for t in d.tables:
            for row in t.rows: out.append(' | '.join(c.text for c in row.cells))
        return '\n'.join(out)
    if ext == '.hwpx':
        from lxml import etree
        z = zipfile.ZipFile(p); out = []
        for n in sorted(x for x in z.namelist() if re.match(r'Contents/section\d+\.xml', x)):
            root = etree.fromstring(z.read(n))
            for para in root.iter('{http://www.hancom.co.kr/hwpml/2011/paragraph}p'):
                s = ''.join(para.xpath('./hp:run/hp:t/text()', namespaces=TB.H.NS)).strip()
                if s: out.append(s)
        return '\n'.join(out)
    return open(p, encoding='utf-8', errors='ignore').read()


@app.post('/api/s/{sid}/bases/ai', dependencies=[Depends(auth)])
def bases_ai(sid: str):
    if not ai.enabled(): raise HTTPException(400, 'API 키가 설정되지 않았습니다')
    q = need(sid, 'questionnaire', '설문지')
    tabs, _, _ = load_tables(sid)
    text = questionnaire_text(q)
    if len(text.strip()) < 200: raise HTTPException(400, '설문지에서 글자를 거의 읽지 못했습니다 (스캔 PDF라면 hwpx/docx로 올려 주세요)')
    return ai.explain_bases([t['raw_base'] for t in tabs], text)


# ---------------------------------------------------------------------- 분석문구 (AI)
class TextReq(BaseModel):
    instruction: str = ''
    previous: Optional[dict] = None
    part: str = 'both'
    store: bool = True


@app.post('/api/s/{sid}/text/{idx}', dependencies=[Depends(auth)])
def write_text(sid: str, idx: str, body: TextReq):
    if not ai.enabled(): raise HTTPException(400, 'API 키가 설정되지 않았습니다')
    tabs, _, s = load_tables(sid)
    t = next((x for x in tabs if x['idx'] == idx), None)
    if t is None: raise HTTPException(404, '표를 찾지 못했습니다')
    if not s['groups']: raise HTTPException(400, '특성별 결과에 쓸 구분을 먼저 고르세요')
    try:
        out = ai.write_text(t, s['groups'], s['min_n'], s['abbr'], body.instruction, body.previous, body.part)
    except Exception as e:
        raise HTTPException(502, f'AI 호출 실패: {e}')
    if body.store:
        st = state(sid); st.setdefault('texts', {})[idx] = dict(overall=out['overall'], group=out['group'], notes=out['notes'],
                                                                problems=out['problems'], source='ai'); save_state(sid, st)
    return out


@app.get('/api/s/{sid}/rule/{idx}', dependencies=[Depends(auth)])
def rule_text(sid: str, idx: str):
    tabs, _, s = load_tables(sid)
    t = next((x for x in tabs if x['idx'] == idx), None)
    if t is None: raise HTTPException(404, '표를 찾지 못했습니다')
    f, rule = TB.facts(t, s['groups'], s['min_n'], s['abbr'])
    return dict(rule=rule, facts=f)


class TextEdit(BaseModel):
    overall: list = []
    group: list = []


@app.put('/api/s/{sid}/text/{idx}', dependencies=[Depends(auth)])
def edit_text(sid: str, idx: str, body: TextEdit):
    tabs, _, s = load_tables(sid)
    t = next((x for x in tabs if x['idx'] == idx), None)
    if t is None: raise HTTPException(404, '표를 찾지 못했습니다')
    ov = [x.strip() for x in body.overall if x.strip()]; gr = [x.strip() for x in body.group if x.strip()]
    probs = TB.check_numbers(ov + gr, t, s['abbr'])
    st = state(sid); old = st.setdefault('texts', {}).get(idx, {})
    st['texts'][idx] = dict(overall=ov, group=gr, notes=old.get('notes', []), problems=probs, source='edit'); save_state(sid, st)
    return dict(problems=probs)


@app.delete('/api/s/{sid}/text/{idx}', dependencies=[Depends(auth)])
def drop_text(sid: str, idx: str):
    st = state(sid); st.get('texts', {}).pop(idx, None); save_state(sid, st); return dict(ok=True)


# ---------------------------------------------------------------------- 한글 보고서 만들기
def report_args(sid, s):
    a = ['--excel', need(sid, 'excel', '통계표 엑셀')]
    a += map_args('--abbr', s['abbr']) + map_args('--base', s['base']) + map_args('--title', s['title'])
    return a


def record(sid, key, path, label):
    st = state(sid); st.setdefault('outputs', {})[key] = dict(file=os.path.basename(path), label=label, at=time.strftime('%H:%M'))
    save_state(sid, st)


@app.post('/api/s/{sid}/build', dependencies=[Depends(auth)])
def build(sid: str):
    st = state(sid); s = settings(st)
    if not s['sample']: raise HTTPException(400, '양식의 예시 표 번호를 고르세요')
    if not s['groups']: raise HTTPException(400, '특성별 결과에 쓸 구분을 고르세요')
    d = sdir(sid)
    texts = {k: dict(overall=v['overall'], group=v['group']) for k, v in st.get('texts', {}).items()}
    tp = os.path.join(d, 'texts.json'); json.dump(texts, open(tp, 'w', encoding='utf-8'), ensure_ascii=False)
    out = os.path.join(d, out_name(sid, '완성.hwpx'))
    cmd = [os.path.join(TB.CORE, 'build_report.py'), '--template', need(sid, 'template', '양식 한글')] + report_args(sid, s) + \
          ['--sample', s['sample'], '--groups', ','.join(s['groups']), '--min-n', str(s['min_n']), '--texts', tp, '--out', out]
    if s['start']: cmd += ['--start', s['start']]
    if s['end']: cmd += ['--end', s['end']]
    if s['fill_sample']: cmd += ['--fill-sample']
    if s['scale_full']: cmd += ['--scale-full']
    code, log = run(cmd)
    if code != 0 or not os.path.exists(out): raise HTTPException(500, '보고서 생성 실패\n' + log[-3000:])
    record(sid, 'report', out, '한글 보고서 (그래프 전)')
    ver = RV.run_verify(need(sid, 'excel', ''), hwpx=out, title=s['title'], base=s['base'], abbr=s['abbr'])
    n_ai = sum(1 for v in st.get('texts', {}).values() if v.get('source') in ('ai', 'edit'))
    return dict(file=os.path.basename(out), log=log[-4000:], verify=ver, ai_texts=n_ai)


# ---------------------------------------------------------------------- PPT 그래프
@app.post('/api/s/{sid}/graphs', dependencies=[Depends(auth)])
def graphs(sid: str):
    st = state(sid); s = settings(st); d = sdir(sid)
    out = os.path.join(d, out_name(sid, '그래프.pptx'))
    cmd = [os.path.join(TB.CORE, 'build_graphs.py'), '--pptx', need(sid, 'graph_template', '그래프 양식 PPT'),
           '--excel', need(sid, 'excel', '통계표 엑셀'), '--out', out] + map_args('--abbr', s['abbr'])
    if s['sample']: cmd += ['--start', s['sample']]
    if s['end']: cmd += ['--end', s['end']]
    code, log = run(cmd)
    if code != 0 or not os.path.exists(out): raise HTTPException(500, '그래프 생성 실패\n' + log[-3000:])
    record(sid, 'graphs', out, '그래프 PPT')
    ver = RV.run_verify(need(sid, 'excel', ''), pptx=out, title=s['title'], base=s['base'], abbr=s['abbr'])
    return dict(file=os.path.basename(out), log=log[-4000:], verify=ver)


class InsertReq(BaseModel):
    target: str = 'report'        # report(이 웹에서 만든 보고서) / edited(올린 보고서)
    use_edited_pptx: bool = True  # 손본 PPT를 올렸으면 그것 사용
    use_emf_zip: bool = False     # PowerPoint에서 저장한 EMF 묶음 사용
    replace_all: bool = False
    format: str = 'png'           # png(300dpi 그림, 기본) / emf(메타파일)


@app.post('/api/s/{sid}/insert', dependencies=[Depends(auth)])
def insert(sid: str, body: InsertReq):
    st = state(sid); d = sdir(sid)
    if body.target == 'edited':
        src = need(sid, 'edited_hwpx', '손본 한글 보고서')
    else:
        o = st.get('outputs', {}).get('report')
        if not o: raise HTTPException(400, '먼저 한글 보고서를 만드세요')
        src = os.path.join(d, o['file'])
    pp = fpath(sid, 'graphs_edited') if body.use_edited_pptx else None
    if not pp or not os.path.exists(pp):
        o = st.get('outputs', {}).get('graphs')
        if not o: raise HTTPException(400, '그래프 PPT를 만들거나 손본 PPT를 올려 주세요')
        pp = os.path.join(d, o['file'])
    out = os.path.join(d, out_name(sid, '완성_그래프포함.hwpx'))
    cmd = [os.path.join(TB.CORE, 'insert_graphs.py'), '--hwpx', src, '--pptx', pp, '--out', out]
    if body.format == 'png': cmd += ['--format', 'png']
    elif body.use_emf_zip: cmd += ['--emf-zip', need(sid, 'emf_zip', 'EMF 묶음(zip)')]
    if body.replace_all: cmd += ['--replace-all']
    code, log = run(cmd)
    if code != 0 or not os.path.exists(out): raise HTTPException(500, '그래프 넣기 실패\n' + log[-3000:])
    record(sid, 'final', out, '한글 보고서 (그래프 포함)')
    m = re.search(r'그래프 (\d+)개 삽입 / 건너뜀 (\d+) / 슬라이드 없음 (\d+)', log)
    return dict(file=os.path.basename(out), log=log[-6000:], inserted=int(m.group(1)) if m else 0,
                skipped=int(m.group(2)) if m else 0, missing=int(m.group(3)) if m else 0,
                pptx=os.path.basename(pp), emf='PNG 300dpi' if body.format == 'png' else ('PowerPoint EMF' if body.use_emf_zip else 'LibreOffice EMF'))


# ---------------------------------------------------------------------- 검증
class ReviewReq(BaseModel):
    target: str = 'final'          # final / report / uploaded


@app.post('/api/s/{sid}/review', dependencies=[Depends(auth)])
def review_start(sid: str, body: ReviewReq):
    st = state(sid); d = sdir(sid); s = settings(st)
    if body.target == 'uploaded':
        hw = need(sid, 'review_hwpx', '검증할 한글 보고서')
    else:
        o = st.get('outputs', {}).get(body.target) or st.get('outputs', {}).get('report')
        if not o: raise HTTPException(400, '검증할 보고서가 없습니다')
        hw = os.path.join(d, o['file'])
    tabs, _, _ = load_tables(sid)
    pp = None
    if body.target != 'uploaded':
        for k in ('graphs_edited',):
            if fpath(sid, k) and os.path.exists(fpath(sid, k)): pp = fpath(sid, k)
        if not pp and st.get('outputs', {}).get('graphs'): pp = os.path.join(d, st['outputs']['graphs']['file'])
    numeric = RV.run_verify(need(sid, 'excel', ''), hwpx=hw, pptx=pp, title=s['title'], base=s['base'], abbr=s['abbr'])
    pages = RV.match(RV.extract_pages(hw), tabs)
    titles = [t['title'] for t in tabs]
    issues = []
    for pg in pages:
        t = next((x for x in tabs if x['idx'] == pg['table']), None)
        pg['num'] = t['num'] if t else None
        for it in RV.local_checks(pg, t, s['groups'], s['min_n'], s['abbr'], titles):
            it.update(page=pg['page'], num=pg['num']); issues.append(it)
    rv = dict(target=body.target, file=os.path.basename(hw), numeric=numeric, pages=pages, issues=issues, ai_done=[])
    json.dump(rv, open(os.path.join(d, 'review.json'), 'w', encoding='utf-8'), ensure_ascii=False)
    return dict(rv, unmatched=[pg['title'] for pg in pages if pg['table'] is None])


@app.post('/api/s/{sid}/review/ai/{page}', dependencies=[Depends(auth)])
def review_ai(sid: str, page: int):
    if not ai.enabled(): raise HTTPException(400, 'API 키가 설정되지 않았습니다')
    d = sdir(sid); p = os.path.join(d, 'review.json')
    if not os.path.exists(p): raise HTTPException(400, '먼저 검증을 시작하세요')
    rv = json.load(open(p, encoding='utf-8'))
    pg = next((x for x in rv['pages'] if x['page'] == page), None)
    if pg is None or pg['table'] is None: return dict(issues=[])
    tabs, _, s = load_tables(sid)
    t = next(x for x in tabs if x['idx'] == pg['table'])
    local = [x for x in rv['issues'] if x['page'] == page]
    try:
        out = ai.review_page(pg, t, s['groups'], s['min_n'], s['abbr'], local)
    except Exception as e:
        raise HTTPException(502, f'AI 호출 실패: {e}')
    for it in out['issues']: it.update(page=page, num=pg.get('num'))
    rv = json.load(open(p, encoding='utf-8'))            # 동시에 여러 페이지를 돌리므로 다시 읽어서 합침
    rv['issues'] = [x for x in rv['issues'] if not (x['page'] == page and x.get('출처') == 'AI')] + out['issues']
    rv['ai_done'] = sorted(set(rv.get('ai_done', [])) | {page})
    json.dump(rv, open(p, 'w', encoding='utf-8'), ensure_ascii=False)
    return out


@app.get('/api/s/{sid}/review.xlsx', dependencies=[Depends(auth)])
def review_xlsx(sid: str):
    d = sdir(sid); rv = json.load(open(os.path.join(d, 'review.json'), encoding='utf-8'))
    out = os.path.join(d, out_name(sid, '검증결과.xlsx'))
    order = {'오류': 0, '확인': 1, '제안': 2}
    issues = sorted(rv['issues'], key=lambda x: (x['page'], order.get(x.get('심각도'), 3)))
    RV.to_xlsx(out, rv['numeric'], rv['pages'], issues)
    return FileResponse(out, filename=os.path.basename(out))


# ---------------------------------------------------------------------- 손본 보고서 문구 다시 쓰기
class RewriteReq(BaseModel):
    keep: int = 0
    only: str = ''
    part: str = 'both'


def rewrite_cmd(sid, s, body, extra):
    return [os.path.join(TB.CORE, 'rewrite_text.py'), '--hwpx', need(sid, 'edited_hwpx', '손본 한글 보고서')] + report_args(sid, s) + \
           ['--sample', s['sample'] or '', '--groups', ','.join(s['groups']), '--min-n', str(s['min_n']),
            '--part', body.part, '--keep', str(body.keep)] + (['--only', body.only] if body.only else []) + extra


@app.post('/api/s/{sid}/rewrite/plan', dependencies=[Depends(auth)])
def rewrite_plan(sid: str, body: RewriteReq):
    st = state(sid); s = settings(st); d = sdir(sid)
    if not s['sample'] or not s['groups']: raise HTTPException(400, '예시 표 번호와 특성별 구분을 먼저 정하세요')
    pp = os.path.join(d, 'plan.json')
    code, log = run(rewrite_cmd(sid, s, body, ['--out', os.path.join(d, 'x.hwpx'), '--plan', pp]))
    if code != 0 or not os.path.exists(pp): raise HTTPException(500, '보고서를 읽지 못했습니다\n' + log[-3000:])
    plan = json.load(open(pp, encoding='utf-8'))
    tabs, _, _ = load_tables(sid)
    names = {t['idx']: t['title'] for t in tabs}
    for p in plan: p['table_titles'] = [names.get(i, '') for i in p['tables']]
    return dict(plan=plan)


class RewriteApply(RewriteReq):
    texts: dict = {}               # {표순번: {overall, group}}


@app.post('/api/s/{sid}/rewrite/apply', dependencies=[Depends(auth)])
def rewrite_apply(sid: str, body: RewriteApply):
    st = state(sid); s = settings(st); d = sdir(sid)
    if not body.texts: raise HTTPException(400, '바꿀 문구가 없습니다')
    tabs, _, _ = load_tables(sid)
    nums = sorted({t['num'] for t in tabs if t['idx'] in body.texts})
    tp = os.path.join(d, 'rewrite_texts.json'); json.dump(body.texts, open(tp, 'w', encoding='utf-8'), ensure_ascii=False)
    out = os.path.join(d, out_name(sid, '문구수정.hwpx'))
    b = RewriteReq(keep=body.keep, only=','.join(nums), part=body.part)
    code, log = run(rewrite_cmd(sid, s, b, ['--texts', tp, '--out', out]))
    if code != 0 or not os.path.exists(out): raise HTTPException(500, '문구 바꾸기 실패\n' + log[-3000:])
    record(sid, 'rewritten', out, '문구 고친 보고서')
    return dict(file=os.path.basename(out), log=log[-4000:])


# ---------------------------------------------------------------------- 내려받기
@app.get('/api/s/{sid}/download/{name}', dependencies=[Depends(auth)])
def download(sid: str, name: str):
    d = sdir(sid); p = os.path.join(d, os.path.basename(name))
    if not os.path.exists(p): raise HTTPException(404, '파일이 없습니다')
    return FileResponse(p, filename=os.path.basename(p))


@app.exception_handler(HTTPException)
async def http_err(request, exc):
    return JSONResponse(status_code=exc.status_code, content=dict(error=exc.detail))


# ---------------------------------------------------------------------- 한 번에 만들기 (간단 화면)
import threading
from concurrent.futures import ThreadPoolExecutor
JOBS = {}


def detect_sample(template, tabs):
    """양식 한글에서 '전체 결과' 바로 앞 소제목이 엑셀 표 제목과 같은 첫 페이지 → 예시 표 번호"""
    pages = RV.extract_pages(template)
    for pg in pages:
        for t in tabs:
            if t['type'] != 'scale' and RV.norm(pg['title']) in (RV.norm(t['name']), RV.norm(t['title'])):
                return t['num']
    return None


@app.post('/api/s/{sid}/prepare', dependencies=[Depends(auth)])
def prepare(sid: str):
    """세 파일이 올라오면: 예시 표 자동 찾기 + 특성별 구분 기본값"""
    tabs, groups, s = load_tables(sid)
    sample = s['sample'] or detect_sample(need(sid, 'template', '양식 한글'), tabs)
    if not s['groups']:
        s['groups'] = [g for g, subs in groups.items() if sum(1 for x in subs if x['n'] >= s['min_n']) >= 2]
    s['sample'] = sample
    st = state(sid); st['settings'] = s; save_state(sid, st)
    return dict(sample=sample, sample_title=next((t['title'] for t in tabs if t['num'] == sample), ''),
                groups=groups, chosen=s['groups'], tables=len(tabs), min_n=s['min_n'])


class RunReq(BaseModel):
    groups: list = []
    sample: Optional[str] = None
    ai: bool = True


def _run(sid, body):
    J = JOBS[sid]
    def step(name, done=0, total=0): J.update(step=name, done=done, total=total)
    try:
        st = state(sid); s = settings(st)
        if body.groups: s['groups'] = body.groups
        if body.sample: s['sample'] = body.sample
        s['fill_sample'] = True                      # 양식 예시 페이지에 남은 다른 조사 문구도 새로 씀
        st['settings'] = s; save_state(sid, st)
        tabs, _, s = load_tables(sid)
        if not s['sample']: raise RuntimeError('양식 한글에서 예시 표를 찾지 못했습니다. 양식의 소제목이 엑셀 표 제목과 같은지 확인해 주세요')
        # 1) 분석문구 (AI)
        T = [t for t in tabs if t['type'] != 'scale']
        iS = next(i for i, t in enumerate(T) if t['num'] == s['sample'])
        todo = T[iS:]
        texts = {}
        if body.ai and ai.enabled():
            step('분석문구 작성 (AI)', 0, len(todo)); cnt = [0]; lock = threading.Lock()
            def one(t):
                try:
                    out = ai.write_text(t, s['groups'], s['min_n'], s['abbr'])
                    with lock: texts[t['idx']] = dict(overall=out['overall'], group=out['group'], notes=out['notes'], problems=out['problems'], source='ai')
                except Exception as e:
                    with lock: J.setdefault('warn', []).append(f"<표{t['num']}> AI 실패 → 규칙 문구 사용 ({str(e)[:80]})")
                with lock: cnt[0] += 1; J['done'] = cnt[0]
            with ThreadPoolExecutor(6) as ex: list(ex.map(one, todo))
            st = state(sid); st['texts'] = texts; save_state(sid, st)
        # 2) 한글 보고서
        step('한글 보고서 만들기'); r = build(sid); J['verify_report'] = r['verify']
        # 3) 그래프 PPT
        has_ppt = bool(fpath(sid, 'graph_template'))
        if has_ppt:
            step('그래프 PPT 만들기'); graphs(sid)
            step('그래프를 한글에 넣기'); r = insert(sid, InsertReq(format='png', use_edited_pptx=False)); J['graphs'] = dict(inserted=r['inserted'], missing=r['missing'])
        # 4) 검증
        step('검증 (숫자 대조·규칙 점검)')
        rv = review_start(sid, ReviewReq(target='final' if has_ppt else 'report'))
        J['review'] = dict(summary=rv['numeric']['summary'], ok=rv['numeric']['ok'],
                           errors=sum(1 for x in rv['issues'] if x['심각도'] == '오류') + sum(1 for x in rv['numeric']['rows'] if x['심각도'] == '오류'),
                           checks=sum(1 for x in rv['issues'] if x['심각도'] == '확인'), unmatched=rv['unmatched'])
        J['notes'] = [f"<표{next(t['num'] for t in tabs if t['idx'] == k)}> " + n for k, v in texts.items() for n in v.get('notes', [])][:40]
        J['problems'] = [f"<표{next(t['num'] for t in tabs if t['idx'] == k)}> " + p for k, v in texts.items() for p in v.get('problems', [])]
        st = state(sid); J['outputs'] = st.get('outputs', {}); J['ai_count'] = len(texts)
        step('완료'); J['status'] = 'done'
    except HTTPException as e:
        J.update(status='error', error=str(e.detail))
    except Exception as e:
        J.update(status='error', error=str(e))


@app.post('/api/s/{sid}/run', dependencies=[Depends(auth)])
def run_all(sid: str, body: RunReq):
    if JOBS.get(sid, {}).get('status') == 'running': raise HTTPException(409, '이미 만들고 있습니다')
    JOBS[sid] = dict(status='running', step='시작', done=0, total=0, started=time.time())
    threading.Thread(target=_run, args=(sid, body), daemon=True).start()
    return JOBS[sid]


@app.get('/api/s/{sid}/run', dependencies=[Depends(auth)])
def run_status(sid: str):
    J = JOBS.get(sid)
    if not J: raise HTTPException(404, '진행 중인 작업이 없습니다')
    return dict(J, elapsed=int(time.time() - J['started']))
