"""hwtx-report 공용 모듈: 엑셀 통계표 파싱, 표 유형 판별, 분석문구 생성, 한글(hwpx) 편집 도우미.

build_report.py / build_graphs.py / rewrite_text.py 가 모두 이 파일을 가져다 쓴다.
"""
import copy, os, re, zipfile
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
from lxml import etree

NS = {'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph'}
HP = '{%s}' % NS['hp']

# ====================================================================== 엑셀식 반올림
def xround(v, d=1):
    """엑셀 화면 값과 같은 반올림: 유효숫자 15자리로 정리한 뒤 사사오입 (28.75 → 28.8).
    파이썬 round()·f'{v:.1f}'는 28.749999…로 보고 28.7로 내리므로 엑셀 값에 쓰지 않는다."""
    if not isinstance(v, (int, float)): return v
    return float(Decimal(format(v, '.15g')).quantize(Decimal(1).scaleb(-d), ROUND_HALF_UP))

def xfmt(v, d=1):
    return f'{xround(v, d):.{d}f}' if isinstance(v, (int, float)) else str(v)

def fmt_dec(number_format, default=1):
    """셀 서식의 소수 자릿수: '.' 뒤의 0 개수만 센다 ('#,##0.0_ ' → 1, '0.00_);(0.00)' → 2, '#,##0' → 0)"""
    f = str(number_format or '')
    if f in ('', 'General'): return default
    f = f.split(';')[0]
    if '.' not in f: return 0 if re.search(r'[0#]', f) else default
    return len(re.match(r'0*', f.split('.', 1)[1]).group(0))

# ====================================================================== 글자/줄바꿈
def cw(ch, em):
    o = ord(ch)
    if 0xAC00 <= o <= 0xD7A3 or 0x3130 <= o <= 0x318F or 0x4E00 <= o <= 0x9FFF or ch in '‘’“”·×①②③④⑤⑥⑦⑧⑨▩～':
        return em
    if ch == ' ': return em * 0.32
    if ch.isdigit(): return em * 0.55
    if ch.isupper(): return em * 0.68
    if ch.isalpha(): return em * 0.52
    return em * 0.38

def wrap(text, width, em):
    """줄 시작 위치(textpos) 목록 — 공백 우선, 불가하면 글자 단위"""
    starts = [0]; cur = 0.0; last_space = None; i = 0; line_start = 0
    while i < len(text):
        w = cw(text[i], em)
        if cur + w > width and i > line_start:
            brk = last_space + 1 if (last_space is not None and last_space > line_start) else i
            starts.append(brk); line_start = brk; last_space = None
            cur = sum(cw(c, em) for c in text[brk:i + 1]); i += 1
            continue
        if text[i] == ' ': last_space = i
        cur += w; i += 1
    return starts

# ====================================================================== 조사/라벨
def has_batchim(word):
    s = (word or '').strip().rstrip('’\'"” .')
    for ch in reversed(s):
        o = ord(ch)
        if 0xAC00 <= o <= 0xD7A3: return (o - 0xAC00) % 28 != 0
        if ch.isdigit(): return ch in '013678'
        if ch.isalpha(): return ch.lower() in 'lmnr'
    return False

def josa(word, a, b):
    """a: 받침 있을 때(이/은/을/과), b: 없을 때(가/는/를/와). 괄호로 끝나면 괄호 앞 글자 기준"""
    core = re.sub(r'\s*\([^)]*\)\s*$', '', (word or '').strip()) or word
    return a if has_batchim(core) else b

ABBR = {}            # --abbr 로 채움: {원문 일부: 축약}
TITLE_MAP = {}       # --title 로 채움: {엑셀 표 제목(또는 번호): 사용자가 정한 제목}  예) '3-1=최종학위 취득국가'
BASE_MAP = {}        # --base 로 채움: {엑셀 BASE 원문: 설문지 보고 풀어 쓴 말}  예) 'SQ(문5) ④~⑦ 응답자(D-10)' → '구직비자(D-10) 응답자'

def lab(s):
    """문장 안에서 쓰는 보기 이름 (표 원문은 그대로 둠)"""
    s = str(s).strip()
    for k, v in ABBR.items(): s = s.replace(k, v)
    if ' / ' in s: s = s.split(' / ')[0]                               # 'A / B' 이중 표기 → 앞쪽
    s = re.sub(r'\s*\(예[:,)\s][^)]*\)', '', s)                        # (예: …) 예시 괄호
    s = re.sub(r'\s*\([^()]*\s등\)', '', s)                            # (서빙, 주방보조 등) 예시 나열
    s = re.sub(r"\s*\((?=[^)]*[A-Za-z])[A-Za-z'’ \-]+\)", '', s)       # (Bachelor's Degree), (Hi-Korea)
    s = s.replace('‘', '“').replace('’', '”').rstrip('.').strip()
    return s

def q(s): return f'‘{lab(s)}’'
def J(s, a, b): return josa(lab(s), a, b)
def pct(v): return f'{xfmt(v, 1)}%'

def clean_label(s):
    s = str(s).replace('\n', ' ').strip()
    s = re.sub(r'\s{2,}', ' ', s)
    if s.count('`') == 1: s = s.replace('`', "'")
    else:
        k = [0]
        def rep(m):
            k[0] += 1; return '‘' if k[0] % 2 else '’'
        s = re.sub('`', rep, s)
    if '평균' in s:
        m = re.search(r'평균\s*-?\s*(\S+)', s.strip('<> '))
        unit = m.group(1).strip('<> ') if m else ''
        return f'평균({unit})' if unit else '평균'
    return s

def clean_base(s):
    s = re.sub(r'^(BASE|Base)\s*[:：]?\s*', '', str(s)).strip()
    return s

# ====================================================================== 엑셀 파싱
def parse_excel(path, sheet=None):
    """회사 표준 통계표: '<표x-y> 제목' → BASE/헤더 행 → (2단 헤더) → ▩전체▩ → 배너 행"""
    wb = openpyxl.load_workbook(path, data_only=True)
    def score(w):   # '<표' 제목 바로 아래에 BASE 행이 있는 시트 (목차 시트 제외)
        return sum(1 for r in range(1, w.max_row) if isinstance(w.cell(r, 1).value, str) and w.cell(r, 1).value.startswith('<표')
                   and str(w.cell(r + 1, 1).value or '').upper().startswith('BASE'))
    ws = wb[sheet] if sheet else max(wb.worksheets, key=score)
    maxc = ws.max_column
    def cell(r, c): return ws.cell(r, c).value
    tables = []; r = 1
    while r <= ws.max_row:
        v = cell(r, 1)
        if not (isinstance(v, str) and v.startswith('<표')):
            r += 1; continue
        hr = r + 1
        sub = hr + 1 if (cell(hr + 1, 1) is None and any(cell(hr + 1, c) is not None for c in range(4, maxc + 1))) else None
        cols = []; carry = None
        for c in range(4, maxc + 1):
            t = cell(hr, c); s = cell(sub, c) if sub else None
            if t is not None: carry = t
            if t is None and s is None: continue
            top = t if t is not None else carry
            if str(top).strip() in ('계', '합계'): continue
            cols.append(dict(c=c, top=clean_label(top), sub=clean_label(s) if s is not None else None))
        # 모든 보기가 같은 'XXX-' 접두어를 가지면 제거 (예: '한국 내 대학 (On-shore)-학사')
        tops = [c['top'] for c in cols if '평균' not in c['top']]
        if len(tops) > 1 and all('-' in x for x in tops):
            pre = os.path.commonprefix(tops)
            if '-' in pre:
                pre = pre[:pre.rfind('-') + 1]
                for c in cols:
                    if c['top'].startswith(pre): c['top'] = c['top'][len(pre):].strip()
        rows = []; grp = None; rr = (sub or hr) + 1
        while rr <= ws.max_row and any(cell(rr, c) is not None for c in range(1, 4)):
            a, b, n = cell(rr, 1), cell(rr, 2), cell(rr, 3)
            if a: grp = '전체' if '▩' in str(a) else str(a).strip()
            try: nn = int(str(n).strip('() ').replace(',', ''))
            except Exception: nn = 0
            vals = []
            for col in cols:
                x = cell(rr, col['c']); f = ws.cell(rr, col['c']).number_format
                dec = fmt_dec(f)
                vals.append((xround(x, dec) if isinstance(x, (int, float)) else 0, dec))   # 읽는 즉시 엑셀식 반올림
            if nn > 0:
                rows.append(dict(grp=grp, label='전체' if grp == '전체' and not b else str(b).strip(), n=nn, vals=vals))
            rr += 1
        m = re.match(r'<표\s*([\d\-]+)>', v)
        tables.append(dict(num=m.group(1) if m else v, name=v.split('>', 1)[1].strip(),
                           base=clean_base(cell(hr, 1)), cols=cols, two=sub is not None, rows=rows))
        r = rr
    return [t for t in tables if t['rows']]

def classify(tables, scale_full=False):
    """표 유형(normal/summary/scale), 제목, 보기 인덱스, 중복응답/평균 여부를 붙인다"""
    nums = [t['num'] for t in tables]
    for i, t in enumerate(tables):
        raw_base = t['base']
        if '요약표' in t['name']:
            t['type'] = 'summary'
            t['title'] = re.sub(r'\s*평균\s*요약표', ' 종합', t['name']).strip()
            t['pos'] = None
            for t2 in tables[i + 1:]:
                if t2['two']:
                    t['pos'] = [c['sub'] for c in t2['cols'] if c['sub']][-1]; break
        elif t['two'] and t['num'].count('-') >= 2 and not scale_full:
            t['type'] = 'scale'; t['title'] = re.sub(r'_(.+)$', r'(\1)', t['name'])
        else:
            t['type'] = 'normal'; t['title'] = re.sub(r'_(.+)$', r'(\1)', t['name'])
        t['title'] = t['title'].replace(' × ', ' 및 ').replace('×', ' 및 ')
        if nums.count(t['num']) > 1 and t['base'] not in ('전체', ''):      # 같은 번호 여러 개 → Base로 구분
            sfx = re.sub(r'[①-⑳]\s*', '', t['base'])
            sfx = re.sub(r"\s*\((?=[^)]*[A-Za-z])[A-Za-z'’ \-]+\)", '', sfx).strip()
            t['title'] += f'({sfx})'
            t['suffix'] = sfx
        t['base'] = BASE_MAP.get(raw_base, raw_base)
        for key in (f"{t['num']}|{raw_base}", t['title'], t['num']):     # 사용자가 정한 제목 우선
            if key in TITLE_MAP: t['title'] = TITLE_MAP[key]; break
        tot = t['rows'][0]
        t['mean_idx'] = [k for k, c in enumerate(t['cols']) if '평균' in c['top']]
        t['item_idx'] = [k for k, c in enumerate(t['cols']) if c['sub']] if t['two'] else \
                        [k for k in range(len(t['cols'])) if k not in t['mean_idx']]
        t['multi'] = t['type'] == 'normal' and not t['two'] and sum(tot['vals'][k][0] for k in t['item_idx']) > 101
        t['unit'] = ''
        if t['mean_idx']:
            m = re.search(r'\(([^)]+)\)', t['cols'][t['mean_idx'][0]]['top'])
            u = m.group(1) if m else ''
            t['unit'] = '점' if '점' in u else ('년' if '년' in u else u)
    return tables

def iname(t, k):
    c = t['cols'][k]; return c['sub'] or c['top']

def unit_text(t):
    u = '(단위: %'
    if t['mean_idx']: u += f", {t['unit'] or '점'}"
    if t['multi']: u += ', 중복응답'
    if t['type'] == 'summary' and t.get('pos'): u += f", {q(t['pos'])} 비율"
    return u + ')'

# ====================================================================== 분석문구: 가. 전체 결과
_CNT = {}
def _turn(key, n):
    _CNT[key] = _CNT.get(key, -1) + 1
    return _CNT[key] % n

def mean_str(t, r):
    v, d = r['vals'][t['mean_idx'][0]]
    return f"{xfmt(v, d)}{t['unit']}"

def overall_paras(t, subj=None):
    """전체 결과 문단 목록. 같은 문형이 연달아 나오지 않게 표마다 돌려 씀"""
    subj = subj or re.sub(r'\s*종합$', '', t['title'])
    tot = t['rows'][0]
    it = [(iname(t, k), tot['vals'][k][0]) for k in t['item_idx']]
    s = sorted(it, key=lambda x: -x[1]); a, rest = s[0], s[1:]
    eul = josa(subj, '을', '를'); eun = josa(subj, '은', '는')
    mean = mean_str(t, tot) if t['mean_idx'] else None
    if t['two']:                                               # 단독 척도표
        lo, mid, hi = it[0], it[len(it) // 2], it[-1]
        out = [f"{subj}{eul} 보면, {q(hi[0])}{J(hi[0],'이라는','라는')} 응답이 {pct(hi[1])}로 {q(lo[0])}({pct(lo[1])}){J(lo[0],'을','를')} 크게 웃돎"
               if hi[1] > lo[1] else
               f"{subj}{eul} 보면, {q(lo[0])}{J(lo[0],'이라는','라는')} 응답이 {pct(lo[1])}로 {q(hi[0])}({pct(hi[1])}){J(hi[0],'보다','보다')} 높게 나타남",
               f"{q(mid[0])}{J(mid[0],'은','는')} {pct(mid[1])}" + (f"이며, 평균은 {mean}임" if mean else '임')]
        return out
    if t['type'] == 'summary':
        P = q(t.get('pos') or '긍정'); vs = [v for _, v in s]
        forms = [
            [f"{subj}{eul} 항목별 {P} 비율로 보면, {q(a[0])}{J(a[0],'이','가')} {pct(a[1])}로 가장 높고, " + ', '.join(f"{q(x)}({pct(v)})" for x, v in rest[:2]) + " 순임"],
            [f"{subj}에서 {P}{J(t.get('pos') or '긍정','이라는','라는')} 응답은 {q(a[0])}에서 {pct(a[1])}로 가장 많았음",
             f"반면 {q(s[-1][0])}{J(s[-1][0],'은','는')} {pct(s[-1][1])}로 상대적으로 낮게 나타남"],
            [f"{subj}{eul} 살펴보면, 모든 항목에서 {P} 비율이 {xfmt(min(vs))}~{xfmt(max(vs))}% 수준으로 나타났으며, 이 중 {q(a[0])}{J(a[0],'이','가')} {pct(a[1])}로 가장 높음"],
        ]
        return forms[_turn('summary', len(forms))]
    if len(it) == 2:
        b = s[1]
        forms = [
            [f"{subj}{eul} 보면, {q(a[0])}{J(a[0],'이','가')} {pct(a[1])}로 {q(b[0])}({pct(b[1])})보다 {xfmt(a[1]-b[1])}%p 높게 나타남"],
            [f"{subj}{eul} 살펴보면, 응답자 10명 중 약 {xfmt(a[1]/10, 0)}명({pct(a[1])})이 {q(a[0])}{J(a[0],'이라고','라고')} 응답함"],
        ]
        return forms[_turn('two', len(forms))]
    mul = '(중복응답)' if t['multi'] else ''
    two = ', '.join(f"{q(x)}({pct(v)})" for x, v in rest[:2])
    last2 = rest[min(1, len(rest) - 1)][0]
    if a[1] >= 50 and not t['multi']:
        out = [f"{subj}{eul} 보면, {q(a[0])}{J(a[0],'이','가')} {pct(a[1])}로 절반 이상을 차지함", f"이어서 {two} 순으로 나타남"]
    elif len(rest) >= 2 and a[1] - rest[0][1] < 3:
        b = rest[0]
        out = [f"{subj}{eul} 보면{mul}, {q(a[0])}({pct(a[1])}){J(a[0],'과','와')} {q(b[0])}({pct(b[1])}){J(b[0],'이','가')} 비슷한 수준으로 높게 나타났으며, "
               f"그다음은 {q(rest[1][0])}({pct(rest[1][1])}) 순임"]
    else:
        forms = [
            [f"{subj}{eul} 보면{mul}, {q(a[0])}{J(a[0],'이','가')} {pct(a[1])}로 가장 높으며, {two} 순임"],
            [f"{subj}{mul}{eun if not mul else '는'} {q(a[0])}{J(a[0],'이','가')} {pct(a[1])}로 가장 많았고, {two}{J(last2,'이','가')} 뒤를 이음"],
            [f"{subj}{eul} 살펴보면{mul}, {q(a[0])} 응답이 {pct(a[1])}로 가장 높게 나타남", f"다음으로 {two} 순임"],
            [f"{subj}의 경우{mul}, {q(a[0])}{J(a[0],'이','가')} {pct(a[1])}로 가장 큰 비중을 차지하였으며, {two} 순으로 나타남"],
        ]
        out = forms[_turn('normal', len(forms))]
    ys = [v for x, v in it if lab(x).startswith('예')]; ns = [v for x, v in it if lab(x).startswith('아니')]
    if len(ys) > 1 and ns:
        out = out + [f"‘예’ 응답을 모두 합하면 {xfmt(sum(ys))}%, ‘아니오’ 응답은 {xfmt(sum(ns))}%임"]
    if mean:
        out = out + [f"평균은 {mean}으로 나타남"]
    return out

# ====================================================================== 분석문구: 나. 특성별 결과
def _tops(t, r):
    mx = max(r['vals'][k][0] for k in t['item_idx'])
    return tuple(k for k in t['item_idx'] if abs(r['vals'][k][0] - mx) < 1e-9)

def glab(s):
    """문장 속 집단 이름: 'D-10-1 (일반구직)' → 'D-10-1(일반구직)'"""
    return re.sub(r'(\w) \(', r'\1(', str(s))

def _join_labels(rs):
    return '·'.join(glab(r['label']) for r in rs)

def _vals(rs, k):
    vs = [pct(r['vals'][k][0]) for r in rs]
    return vs[0] if len(vs) == 1 else ', '.join(vs)

def _item_phrase(t, keys, rs, summ):
    """‘X’(v1, v2)  /  동률: ‘X’·‘Y’(각 v)"""
    if len(keys) == 1:
        return f"{q(iname(t, keys[0]))}({_vals(rs, keys[0])})"
    return '·'.join(q(iname(t, k)) for k in keys) + f"(각 {pct(rs[0]['vals'][keys[0]][0])})"

def _tail(t, last_key, summ, verb='가장 높게 나타남'):
    return f"{J(iname(t, last_key),'이','가')} {verb}"

def overall_tagged(t, tag):
    """한 페이지에 하위표 여러 개를 합친 경우의 짧은 전체 결과: (한국 내 대학) ‘학사’가 43.7%로 가장 높으며, …"""
    tot = t['rows'][0]
    s = sorted([(iname(t, k), tot['vals'][k][0]) for k in t['item_idx']], key=lambda x: -x[1])
    a, rest = s[0], s[1:]
    two = ', '.join(f"{q(x)}({pct(v)})" for x, v in rest[:2])
    if a[1] >= 70: return f"({tag}) {q(a[0])}{J(a[0],'이','가')} {pct(a[1])}로 대부분을 차지하며, {two} 순임"
    if rest and a[1] - rest[0][1] < 3:
        return f"({tag}) {q(a[0])}({pct(a[1])}){J(a[0],'과','와')} {q(rest[0][0])}({pct(rest[0][1])}){J(rest[0][0],'이','가')} 비슷한 수준으로 높게 나타남"
    return f"({tag}) {q(a[0])}{J(a[0],'이','가')} {pct(a[1])}로 가장 높으며, {two} 순임"

def group_paras(t, groups, min_n=10, tag=''):
    """구분(배너)마다 한 문단.
    - 사례수 min_n 미만 하위집단은 제외 (기본 10명)
    - 평균 열이 있는 표는 평균 기준으로 서술
    - 1위가 모두 같으면 2위·집단 차이 중심, 다르면 같은 1위끼리 묶어 나열
    - 화살표·괄호 남발 없이 짧게"""
    summ = t['type'] == 'summary'
    out = []
    for g in groups:
        rs = [r for r in t['rows'][1:] if r['grp'] == g and r['n'] >= min_n]
        if not rs: continue
        head = f"({g}{', ' + tag if tag else ''}) "
        # ---------- 평균 기준
        if t['mean_idx'] and not t['two']:
            mk = t['mean_idx'][0]; u = t['unit']; d = rs[0]['vals'][mk][1]
            hi, lo, more = ('길고', '짧게', '길게') if u == '년' else ('높고', '낮게', '높게')
            rv = lambda r: xround(r['vals'][mk][0], d)
            mv = lambda r: f"{xfmt(rv(r), d)}{u}"
            ro = lambda x: josa(x, '으로', '로')
            if len(rs) == 1:
                out.append(f"{head}평균은 {glab(rs[0]['label'])} {mv(rs[0])}임"); continue
            if len(rs) == 2:
                a, b = rs; diff = xround(rv(a) - rv(b), d); big = a if diff >= 0 else b; L = glab(big['label'])
                out.append(f"{head}평균은 {glab(a['label'])} {mv(a)}, {glab(b['label'])} {mv(b)}{ro(mv(b))} {L}{josa(L,'이','가')} "
                           f"{xfmt(abs(diff), d)}{u} 더 {more} 나타남"); continue
            srt = sorted(rs, key=lambda r: -rv(r))
            tops = [r for r in srt if rv(r) == rv(srt[0])]; bots = [r for r in srt if rv(r) == rv(srt[-1])]
            tl = _join_labels(tops); bl = _join_labels(bots)
            out.append(f"{head}평균은 {tl}{josa(tl,'이','가')} {mv(tops[0])}{ro(mv(tops[0]))} 가장 {hi}, "
                       f"{bl}{josa(bl,'이','가')} {mv(bots[0])}{ro(mv(bots[0]))} 가장 {lo} 나타남")
            continue
        tops = [_tops(t, r) for r in rs]
        # ---------- 집단 1개
        if len(rs) == 1:
            k = tops[0][-1]
            out.append(f"{head}{glab(rs[0]['label'])}{josa(rs[0]['label'],'은','는')} {_item_phrase(t, tops[0], rs, summ)}{_tail(t, k, summ)}"); continue
        # ---------- 1위가 모두 같음 → 2위·차이 중심
        if len(set(tops)) == 1 and len(tops[0]) == 1:
            k = tops[0][0]; X = q(iname(t, k))
            seconds = []
            for r in rs:
                rest = [kk for kk in t['item_idx'] if kk != k]
                seconds.append(max(rest, key=lambda kk: r['vals'][kk][0]) if rest else None)
            gap_big = max(r['vals'][k][0] for r in rs) - min(r['vals'][k][0] for r in rs)
            if len(t['item_idx']) > 2 and len(set(seconds)) > 1 and len(rs) <= 6:
                buckets = {}
                for r, s2 in zip(rs, seconds): buckets.setdefault(s2, []).append(r)
                parts = [f"{_join_labels(b)}{josa(_join_labels(b),'은','는')} {q(iname(t, s2))}({_vals(b, s2)})" for s2, b in buckets.items()]
                last = list(buckets)[-1]
                lead = f"{glab(rs[0]['label'])}{josa(rs[0]['label'],'과','와')} {glab(rs[1]['label'])} 모두" if len(rs) == 2 else f"모든 {g}에서"
                out.append(f"{head}{lead} {X}{J(iname(t,k),'이','가')} 가장 높으나, " + ', '.join(parts) + f"{J(iname(t, last),'이','가')} 다음으로 높게 나타남")
            elif len(rs) == 2:
                a, b = rs
                out.append(f"{head}{glab(a['label'])}({pct(a['vals'][k][0])}){josa(a['label'],'과','와')} {glab(b['label'])}({pct(b['vals'][k][0])}) 모두 "
                           f"{X}{J(iname(t,k),'이','가')} 가장 높게 나타남")
            else:
                mx = max(rs, key=lambda r: r['vals'][k][0])
                if gap_big >= 10:
                    out.append(f"{head}모든 {g}에서 {X}{J(iname(t,k),'이','가')} 가장 높으며, 특히 {glab(mx['label'])}({pct(mx['vals'][k][0])})에서 높게 나타남")
                else:
                    out.append(f"{head}모든 {g}에서 {X}{J(iname(t,k),'이','가')} 가장 높게 나타남")
            continue
        # ---------- 1위가 다름 → 같은 1위끼리 묶어서
        buckets = {}
        for r, tk in zip(rs, tops): buckets.setdefault(tk, []).append(r)
        major = max(buckets, key=lambda k_: len(buckets[k_]))
        if len(rs) >= 5 and len(buckets[major]) >= len(rs) - 2 and len(major) == 1:   # 한두 집단만 다름 → '대부분'
            others = [(tk, b) for tk, b in buckets.items() if tk != major]
            parts = [f"{_join_labels(b)}{josa(b[-1]['label'],'은','는')} {_item_phrase(t, tk, b, summ)}" for tk, b in others]
            last = others[-1][0][-1]
            out.append(f"{head}대부분 {g}에서 {q(iname(t, major[0]))}{J(iname(t, major[0]),'이','가')} 가장 높으나, "
                       + ', '.join(parts) + _tail(t, last, summ)); continue
        parts = []
        for tk, b in buckets.items():
            L = _join_labels(b)
            parts.append(f"{L}{josa(b[-1]['label'],'은','는')} {_item_phrase(t, tk, b, summ)}")
        last = list(buckets)[-1][-1]
        out.append(head + ', '.join(parts) + _tail(t, last, summ))
    return out

# ====================================================================== hwpx 도우미
def load_charpr(header_txt):
    """charPr id → 실제 글자 폭(em) 추정"""
    em = {}
    for m in re.finditer(r'<hh:charPr id="(\d+)" height="(\d+)".*?<hh:ratio hangul="(\d+)"', header_txt, re.S):
        em[m.group(1)] = int(m.group(2)) * int(m.group(3)) / 100 * 0.95
    return em

class Doc:
    """hwpx 압축 해제 → section0.xml / header.xml 편집 → 다시 묶기"""
    def __init__(self, path, workdir):
        self.src = path; self.dir = workdir
        zipfile.ZipFile(path).extractall(workdir)
        self.sec = os.path.join(workdir, 'Contents/section0.xml')
        self.tree = etree.parse(self.sec); self.root = self.tree.getroot()
        self.hdr_path = os.path.join(workdir, 'Contents/header.xml')
        self.hdr = open(self.hdr_path, encoding='utf-8').read()
        self.em = load_charpr(self.hdr)
        self._nonum = {}

    def paras(self): return list(self.root)

    @staticmethod
    def text(p): return ''.join(p.xpath('.//hp:t/text()', namespaces=NS)).strip()

    @staticmethod
    def has_tbl(p): return p.find('.//' + HP + 'tbl') is not None

    def set_text(self, p, text):
        """문단 첫 run 텍스트 교체 + 줄 배치(lineseg) 재계산. 폭·글자크기는 원래 문단에서 읽음"""
        run = p.findall(HP + 'run')[0]
        for el in run.findall(HP + 't'): run.remove(el)
        for r2 in p.findall(HP + 'run')[1:]:
            for el in r2.findall(HP + 't'): el.text = ''
        etree.SubElement(run, HP + 't').text = text
        lsa = p.find(HP + 'linesegarray'); segs = lsa.findall(HP + 'lineseg')
        proto = copy.deepcopy(segs[0])
        f1 = proto.get('flags'); f2 = segs[1].get('flags') if len(segs) > 1 else '1441792'
        for s in segs: lsa.remove(s)
        width = int(proto.get('horzsize')) - 300
        em = self.em.get(run.get('charPrIDRef'), 950)
        v0 = int(proto.get('vertpos')); step = int(proto.get('vertsize')) + int(proto.get('spacing'))
        for j, st in enumerate(wrap(text, width, em)):
            s = copy.deepcopy(proto); s.set('textpos', str(st)); s.set('vertpos', str(v0 + j * step))
            s.set('flags', f1 if j == 0 else f2); lsa.append(s)

    def nonum_pp(self, pp_id):
        """자동번호가 붙은 문단모양을 복제해 번호 없는 문단모양을 만든다 (가./나.를 글자로 쓰기 위해)"""
        if pp_id in self._nonum: return self._nonum[pp_id]
        m = re.search(r'<hh:paraPr id="%s".*?</hh:paraPr>' % pp_id, self.hdr, re.S)
        if not m or 'type="NONE"' in m.group(0)[:600]:
            self._nonum[pp_id] = pp_id; return pp_id
        new = max(int(x) for x in re.findall(r'<hh:paraPr id="(\d+)"', self.hdr)) + 1
        pp = m.group(0).replace(f'id="{pp_id}"', f'id="{new}"', 1)
        pp = re.sub(r'<hh:heading [^>]*/>', '<hh:heading type="NONE" idRef="0" level="0"/>', pp, 1)
        self.hdr = self.hdr.replace('</hh:paraProperties>', pp + '</hh:paraProperties>')
        self.hdr = re.sub(r'(<hh:paraProperties itemCnt=")(\d+)"', lambda mm: f'{mm.group(1)}{int(mm.group(2)) + 1}"', self.hdr)
        self._nonum[pp_id] = str(new); return str(new)

    def sub_head(self, p, text):
        p.set('paraPrIDRef', self.nonum_pp(p.get('paraPrIDRef'))); self.set_text(p, text)

    def save(self, out):
        open(self.hdr_path, 'w', encoding='utf-8').write(self.hdr)
        self.tree.write(self.sec, xml_declaration=True, encoding='UTF-8', standalone=True)
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        src = zipfile.ZipFile(self.src)
        with zipfile.ZipFile(out, 'w') as z:                     # mimetype 맨 앞·무압축이어야 한글에서 열림
            for info in src.infolist():
                data = open(os.path.join(self.dir, info.filename), 'rb').read()
                z.writestr(zipfile.ZipInfo(info.filename, date_time=info.date_time), data,
                           compress_type=zipfile.ZIP_STORED if info.filename == 'mimetype' else zipfile.ZIP_DEFLATED)
            for name in getattr(self, 'extra', []):              # [웹앱] 새로 넣은 파일 (그래프 그림 등)
                if name in src.namelist(): continue
                z.write(os.path.join(self.dir, name), name, compress_type=zipfile.ZIP_DEFLATED)
