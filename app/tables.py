"""엑셀 통계표 파싱(기존 hwtx_common 그대로)과, AI에 넘길 '사실 묶음'·허용 수치 목록 만들기.

숫자는 모두 hwtx_common의 xround/xfmt(엑셀 화면값과 같은 반올림)를 거친 문자열로만 다룬다.
AI는 여기서 만든 문자열을 그대로 옮겨 쓰기만 하고, 계산은 하지 않는다.
"""
import itertools, os, re, sys, threading

CORE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'core')
sys.path.insert(0, CORE)
import hwtx_common as H  # noqa: E402
from hwtx_common import xfmt, xround  # noqa: E402

LOCK = threading.Lock()   # hwtx_common이 ABBR/BASE_MAP/TITLE_MAP 전역을 쓰므로 요청마다 잠그고 설정


def set_maps(abbr=None, base=None, title=None):
    H.ABBR.clear(); H.ABBR.update(abbr or {})
    H.BASE_MAP.clear(); H.BASE_MAP.update(base or {})
    H.TITLE_MAP.clear(); H.TITLE_MAP.update(title or {})
    H._CNT.clear()


def parse(path, abbr=None, base=None, title=None, scale_full=False):
    with LOCK:
        set_maps(abbr, base, title)
        raw = H.parse_excel(path)
        raw_bases = [t['base'] for t in raw]
        tables = H.classify(raw, scale_full=scale_full)
    for i, t in enumerate(tables):
        t['idx'] = str(i)
        t['raw_base'] = raw_bases[i]
        for r in t['rows']:
            r['vals'] = [list(v) for v in r['vals']]
    groups = {}
    for t in tables:
        for r in t['rows'][1:]:
            groups.setdefault(r['grp'], {}).setdefault(r['label'], r['n'])
    return tables, {g: [{'label': k, 'n': v} for k, v in d.items()] for g, d in groups.items()}


def summary(t):
    tot = t['rows'][0]
    return dict(idx=t['idx'], num=t['num'], name=t['name'], title=t['title'], base=t['base'], raw_base=t['raw_base'],
                type=t['type'], multi=t['multi'], mean=bool(t['mean_idx']), unit=t['unit'], n=tot['n'],
                items=[dict(name=H.iname(t, k), total=xfmt(*tot['vals'][k])) for k in t['item_idx']])


# ---------------------------------------------------------------------- AI 입력 사실 묶음
def facts(t, groups, min_n=10, abbr=None):
    """AI에게 줄 표 정보. 수치는 모두 표시용 문자열."""
    with LOCK:
        set_maps(abbr)
        tot = t['rows'][0]
        subj = re.sub(r'\s*종합$', '', t['title'])
        item_name = {k: H.lab(H.iname(t, k)) for k in t['item_idx']}
        total = [dict(보기=item_name[k], 값=xfmt(*tot['vals'][k])) for k in t['item_idx']]
        ranked = sorted(total, key=lambda x: -float(x['값']))
        f = dict(주어=subj, 표번호=t['num'], Base=t['base'], 전체사례수=tot['n'],
                 표유형={'normal': '일반', 'summary': '요약표(항목별 긍정 비율)', 'scale': '척도 세부표'}[t['type']]
                 + (' / 단독 척도표(2단 헤더)' if t['two'] and t['type'] != 'summary' else ''),
                 중복응답=t['multi'], 단위=H.unit_text(t), 전체결과_보기순=total, 전체결과_높은순=ranked)
        if t['type'] == 'summary' and t.get('pos'):
            f['긍정보기'] = H.lab(t['pos'])
        if t['mean_idx']:
            f['평균'] = H.mean_str(t, tot)
        if len(ranked) >= 2:
            a, b = float(ranked[0]['값']), float(ranked[1]['값'])
            f['1위와2위차이'] = xfmt(xround(a - b, 1)) + '%p'
            if a - b < 3:
                f['주의'] = '1·2위 차이가 3%p 미만 → "비슷한 수준"으로 서술'
        gs = []
        for g in groups:
            rs = [r for r in t['rows'][1:] if r['grp'] == g]
            if not rs:
                continue
            use = [r for r in rs if r['n'] >= min_n]
            drop = [r for r in rs if r['n'] < min_n]
            sub = []
            for r in use:
                vals = {item_name[k]: xfmt(*r['vals'][k]) for k in t['item_idx']}
                mx = max(r['vals'][k][0] for k in t['item_idx'])
                d = dict(집단=H.glab(r['label']), 사례수=r['n'], 값=vals,
                         가장높은보기=[item_name[k] for k in t['item_idx'] if abs(r['vals'][k][0] - mx) < 1e-9])
                srt = sorted(t['item_idx'], key=lambda k: -r['vals'][k][0])
                if len(srt) > 1:
                    d['두번째보기'] = item_name[srt[1]]
                    gap = xround(r['vals'][srt[0]][0] - r['vals'][srt[1]][0], 1)
                    if gap <= 1.0:
                        d['주의'] = f'1·2위 차이 {xfmt(gap)}%p — 단정하지 말 것'
                if t['mean_idx']:
                    d['평균'] = H.mean_str(t, r)
                if r['n'] < 30:
                    d['사례수주의'] = '30명 미만 — 해석 주의'
                sub.append(d)
            if not use:
                continue
            gs.append(dict(구분=g, 하위집단=sub, 제외된집단=[f"{H.glab(r['label'])}({r['n']}명)" for r in drop]))
        f['특성별'] = gs
        f['특성별_쓰는순서'] = [x['구분'] for x in gs]
        rule_ov = H.overall_paras(t)
        rule_gr = H.group_paras(t, groups, min_n)
    return f, dict(overall=rule_ov, group=rule_gr)


# ---------------------------------------------------------------------- 문구 속 수치 대조
NUM = re.compile(r'\d[\d,]*(?:\.\d+)?')


def allowed_numbers(t, abbr=None):
    """문구에 써도 되는 수치 문자열: 표시값, 사례수, 같은 행·열 안의 차이(%p), 예/아니오 합, 라벨 속 숫자, 1~10"""
    ok = {str(i) for i in range(0, 11)}
    rows = t['rows']; ncol = len(t['cols'])
    for r in rows:
        ok |= {str(r['n']), f"{r['n']:,}"}
        disp = [xfmt(*r['vals'][k]) for k in range(ncol)]
        ok |= set(disp)
        for a, b in itertools.combinations(range(ncol), 2):
            d = max(r['vals'][a][1], r['vals'][b][1])
            ok.add(xfmt(abs(xround(r['vals'][a][0] - r['vals'][b][0], d)), d))
    for k in range(ncol):
        col = [r['vals'][k] for r in rows]
        for (va, da), (vb, db) in itertools.combinations(col, 2):
            d = max(da, db); ok.add(xfmt(abs(xround(va - vb, d)), d))
    # 예/아니오 합계, 상위 n개 누적
    tot = rows[0]
    items = [(H.lab(H.iname(t, k)), tot['vals'][k][0]) for k in t['item_idx']]
    ys = [v for x, v in items if x.startswith('예')]; ns = [v for x, v in items if x.startswith('아니')]
    if ys: ok.add(xfmt(sum(ys)))
    if ns: ok.add(xfmt(sum(ns)))
    srt = sorted((v for _, v in items), reverse=True)
    for n in range(2, len(srt) + 1): ok.add(xfmt(sum(srt[:n])))
    text = ' '.join([t['title'], t['name'], t['base'], t['raw_base']] + [r['label'] + ' ' + r['grp'] for r in rows]
                    + [c['top'] + ' ' + (c['sub'] or '') for c in t['cols']])
    ok |= set(NUM.findall(text))
    with LOCK:
        set_maps(abbr)
        ok |= set(NUM.findall(' '.join(H.lab(H.iname(t, k)) for k in range(ncol))))
    return ok


def check_numbers(lines, t, abbr=None):
    """허용 목록에 없는 수치 → 문제 목록"""
    ok = allowed_numbers(t, abbr)
    bad = []
    for s in lines:
        s2 = re.sub(r'‘[^’]*’', ' ', s)          # 보기 이름 안의 숫자는 제외
        for tok in NUM.findall(s2):
            tok = tok.rstrip(',')
            if tok not in ok and tok.replace(',', '') not in ok:
                bad.append(f'"{tok}" ← {s[:60]}')
    return bad
