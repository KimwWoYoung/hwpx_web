"""Claude API 호출: 분석문구 작성, Base 문구 풀어쓰기.

- API 키는 서버 환경변수 ANTHROPIC_API_KEY 에만 둔다 (브라우저로 보내지 않음)
- 시스템 프롬프트(케이스탯 문구 규칙)는 prompt caching 으로 재사용
- 작성된 문구 속 수치는 tables.check_numbers 로 대조하고, 틀리면 한 번 고쳐 쓰게 한다
"""
import json, os, re

from . import tables as TB

MODEL = os.environ.get('HWTX_MODEL', 'claude-opus-5-5')
STYLE = open(os.path.join(TB.CORE, 'style.md'), encoding='utf-8').read()
EXTRA = os.environ.get('HWTX_EXTRA_RULES', '')          # 회사·조사별 추가 규칙 (선택)

SYSTEM = f"""당신은 케이스탯리서치 조사 보고서의 분석문구를 쓰는 연구원입니다.
엑셀 통계표 한 개의 정리된 수치를 받아, 한글 보고서의 '가. 전체 결과'와 '나. 특성별 결과' 문단을 씁니다.

# 반드시 지킬 것
1. 수치는 입력 JSON에 있는 문자열을 글자 그대로 옮긴다. 반올림·재계산·추정 금지.
   %p 차이는 입력의 '1위와2위차이'나, 같은 집단 안 두 값의 차이를 쓸 때만 쓰고 확신이 없으면 쓰지 않는다.
2. '제외된집단'(사례수 10명 미만)은 문장에 넣지 않는다. '사례수주의' 집단은 넣되 단정적 비교는 피한다.
3. '주의'가 붙은 곳(1·2위 차이가 작음)은 "비슷한 수준"으로 쓰거나 단정하지 않는다.
4. 특성별 문단은 '특성별_쓰는순서'의 구분마다 정확히 한 문단, 그 순서대로, "(구분명) "으로 시작한다.
5. 보기 이름은 입력에 있는 그대로 ‘ ’(작은따옴표)로 감싼다. 집단 이름은 입력 그대로.
6. 같은 문형만 반복하지 말고, 모든 집단의 1위가 같으면 2위나 집단 간 차이를 짚는다.
7. '참고 초안'은 규칙으로 만든 문장이라 수치는 정확하지만 어색할 수 있다. 더 자연스럽고 정확한 보고서 문장으로 다듬는다.

# 문구 규칙 (케이스탯 보고서 기준)
{STYLE}
{('# 추가 규칙' + chr(10) + EXTRA) if EXTRA else ''}

# 출력
설명 없이 JSON 하나만 출력한다.
{{"overall": ["가. 전체 결과 문단", ...], "group": ["(구분) 문단", ...], "notes": ["연구원이 확인할 점(없으면 빈 배열)"]}}
"""

_client = None


def client():
    global _client
    if _client is None:
        import anthropic
        _client = anthropic.Anthropic()          # ANTHROPIC_API_KEY 환경변수 사용
    return _client


def enabled():
    return bool(os.environ.get('ANTHROPIC_API_KEY')) or os.environ.get('HWTX_FAKE_AI') == '1'


def _ask(system, messages, max_tokens=2000):
    if os.environ.get('HWTX_FAKE_AI') == '1':             # 키 없이 화면 테스트용
        return _fake(messages), {}
    r = client().messages.create(
        model=MODEL, max_tokens=max_tokens,
        system=[{'type': 'text', 'text': system, 'cache_control': {'type': 'ephemeral'}}],
        messages=messages)
    text = ''.join(b.text for b in r.content if getattr(b, 'type', '') == 'text')
    u = r.usage
    return text, dict(input=u.input_tokens, output=u.output_tokens,
                      cache_read=getattr(u, 'cache_read_input_tokens', 0) or 0)


def _json(text):
    m = re.search(r'\{.*\}', text, re.S)
    if not m:
        raise ValueError('AI 응답에서 JSON을 찾지 못함: ' + text[:200])
    return json.loads(m.group(0))


def _fake(messages):
    """HWTX_FAKE_AI=1: 참고 초안을 그대로 돌려줌 (API 없이 흐름 점검)"""
    body = messages[0]['content']
    draft = json.loads(body.split('## 참고 초안\n', 1)[1].split('\n## ', 1)[0])
    return json.dumps(dict(overall=draft['overall'], group=draft['group'], notes=['(테스트 모드: 규칙 초안 그대로)']), ensure_ascii=False)


def write_text(t, groups, min_n=10, abbr=None, instruction='', previous=None, part='both'):
    f, rule = TB.facts(t, groups, min_n, abbr)
    body = ('## 표 정보\n' + json.dumps(f, ensure_ascii=False, indent=1)
            + '\n## 참고 초안\n' + json.dumps(rule, ensure_ascii=False) + '\n')
    if previous:
        body += '## 현재 문구 (이것을 고친다)\n' + json.dumps(previous, ensure_ascii=False) + '\n'
    if instruction:
        body += f'## 연구원 요청\n{instruction}\n'
    if part != 'both' and previous:
        keep = 'group' if part == 'overall' else 'overall'
        body += f'## 범위\n{"가. 전체 결과" if part == "overall" else "나. 특성별 결과"}만 고치고 {keep} 는 현재 문구 그대로 둔다.\n'
    msgs = [{'role': 'user', 'content': body}]
    usage = []
    for attempt in range(2):
        text, u = _ask(SYSTEM, msgs); usage.append(u)
        out = _json(text)
        out.setdefault('notes', [])
        problems = TB.check_numbers(out.get('overall', []) + out.get('group', []), t, abbr)
        want = f['특성별_쓰는순서']
        heads = [re.match(r'\(([^),]+)', s).group(1) if re.match(r'\(([^),]+)', s) else '' for s in out.get('group', [])]
        if heads != want:
            problems.append(f'특성별 문단 구분/순서가 요청({", ".join(want)})과 다름: {", ".join(heads)}')
        if not problems or attempt == 1:
            break
        msgs += [{'role': 'assistant', 'content': text},
                 {'role': 'user', 'content': '다음 문제를 고쳐 JSON만 다시 출력하세요. 수치는 표 정보의 문자열만 쓰세요.\n- ' + '\n- '.join(problems)}]
    return dict(overall=out.get('overall', []), group=out.get('group', []), notes=out.get('notes', []),
                problems=problems, rule=rule, usage=usage, model=MODEL)


# ---------------------------------------------------------------------- Base 풀어쓰기
BASE_SYSTEM = """설문조사 통계표의 BASE(응답자 조건) 표기를 보고서용 쉬운 말로 바꿉니다.
설문지에서 해당 문항 번호와 보기를 찾아, 그 응답자가 누구인지 짧은 명사구로 씁니다 (예: '~ 응답자').
설문지에서 근거를 찾지 못하면 원문을 그대로 두고 note에 이유를 씁니다.

예시 (외국인 구직자 조사)
- SQ(문5) ④~⑦ 응답자(D-10) → 구직비자(D-10) 응답자
- 문7·문8 동일 연도 응답자 제외 → 현 비자 이전 다른 비자 체류자
- 문15 ⑤(아직 졸업 안함) 제외 응답자 → 대학 졸업 응답자
- 문25①(경험 있다) 응답자 → 한국 내 취업경험 응답자

출력: JSON만. {"map": {"원문": "풀어 쓴 말", ...}, "notes": {"원문": "근거 또는 못 찾은 이유", ...}}"""


def explain_bases(bases, questionnaire_text):
    bases = [b for b in dict.fromkeys(bases) if b and b != '전체']
    if not bases:
        return dict(map={}, notes={})
    if os.environ.get('HWTX_FAKE_AI') == '1':
        return dict(map={b: b for b in bases}, notes={b: '테스트 모드' for b in bases})
    q = questionnaire_text[:120000]
    msgs = [{'role': 'user', 'content': f'## 설문지\n{q}\n\n## 바꿀 BASE 목록\n' + '\n'.join(f'- {b}' for b in bases)}]
    text, u = _ask(BASE_SYSTEM, msgs, max_tokens=3000)
    out = _json(text)
    return dict(map=out.get('map', {}), notes=out.get('notes', {}), usage=u)


# ---------------------------------------------------------------------- 보고서 내용 검토
REVIEW_SYSTEM = f"""당신은 케이스탯리서치 조사 보고서를 납품 전에 검토하는 선임 연구원입니다.
한 페이지의 분석문구(가. 전체 결과 / 나. 특성별 결과)를 원자료 표 정보와 대조해 문제를 찾습니다.

# 볼 것 (중요한 순서)
1. 해석 오류: '가장 높음/낮음', 순위, 'A가 B보다 높음' 같은 비교가 표 정보와 맞지 않는 곳
2. 수치 오류: 문장 속 수치가 해당 보기·집단의 값과 다른 곳 (값이 다른 보기·집단 것과 바뀐 경우 포함)
3. 표본 문제: '제외된집단'(사례수 기준 미만)을 언급한 곳, 30명 미만 집단을 단정적으로 비교한 곳
4. 단정 표현: 1·2위 차이가 1%p 안팎인데 차이가 있는 것처럼 쓴 곳
5. 내용 불일치: 다른 조사·다른 표의 문구가 남아 있는 곳, 표에 없는 구분·보기를 말하는 곳
6. 문체: 아래 문구 규칙 위반 (서술형 종결, 마침표, 조사 오류, 화살표·주관적 수식어, 보기 따옴표)

# 규칙
- 근거에는 표 정보의 수치 문자열을 그대로 인용한다. 표 정보로 확인할 수 없는 것은 지적하지 않는다.
- '규칙 점검 결과'에 이미 있는 문제는 반복하지 않는다.
- 문제가 없으면 빈 배열.
- 수정제안은 문구 규칙에 맞는 완성 문장으로 쓴다.

# 문구 규칙
{STYLE}

# 출력 (JSON만)
{{"issues": [{{"위치": "가. 전체 결과|나. 특성별 결과", "문장": "문제 문장 일부", "문제": "…", "근거": "…", "수정제안": "…", "심각도": "오류|확인|제안"}}]}}
"""


def review_page(pg, t, groups, min_n=10, abbr=None, local=None):
    if os.environ.get('HWTX_FAKE_AI') == '1':
        return dict(issues=[], usage={})
    all_groups = list(dict.fromkeys(r['grp'] for r in t['rows'][1:]))
    f, _ = TB.facts(t, all_groups, min_n, abbr)
    body = ('## 표 정보\n' + json.dumps(f, ensure_ascii=False, indent=1)
            + f"\n## 특성별 결과로 쓰기로 한 구분 (순서)\n{', '.join(groups) or '(지정 없음)'}\n"
            + '## 페이지 문구\n' + json.dumps(dict(제목=pg['title'], 전체결과=pg['overall'], 특성별결과=pg['group'],
                                                 Base=pg['base']), ensure_ascii=False, indent=1)
            + '\n## 규칙 점검 결과 (이미 찾은 것)\n' + json.dumps([dict(문제=x['문제'], 문장=x['문장']) for x in (local or [])], ensure_ascii=False))
    text, u = _ask(REVIEW_SYSTEM, [{'role': 'user', 'content': body}], max_tokens=3000)
    out = _json(text)
    issues = out.get('issues', [])
    for it in issues: it['출처'] = 'AI'
    return dict(issues=issues, usage=u)
