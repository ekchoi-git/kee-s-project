"""회사 GPT 키: 키 종류 자동 구분, 인증 토큰·요청 모양, 응답 해석, 실패 안내를 가짜 회사 서버로 확인(사내망 불필요).
사용: python tests/test_company_provider.py"""
import base64, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import httpx2
from app import ai

URL = "http://gateway.invalid/api"  # 가짜 주소
KEY = "9a1b2c3d-4e5f-6789-abcd-ef0123456789"  # 가짜 키(36자리, 하이픈 빼면 32자리)

# 키 종류 자동 구분: 앞글자가 아니라 모양으로
assert ai.provider(KEY) == "회사 GPT" and ai.provider(KEY.replace("-", "")) == "회사 GPT"
assert ai.provider("sk-ant-api03-x") == "Claude" and ai.provider("sk-proj-abc") == "OpenAI"
assert not ai.is_company(KEY[:-1]) and not ai.is_company(KEY.replace("a", "한", 1))

sent = []
replies = []


def server(req):
    sent.append(req)
    return replies.pop(0)


def client():
    c = ai.connect(KEY, emp="123456", url=URL)
    assert isinstance(c, ai.CompanyClient)
    c.http = httpx2.Client(transport=httpx2.MockTransport(server), headers=c.http.headers)
    return c


def origin(content, tokens=42):  # need_origin=True일 때 온다고 가정한 OpenAI 원본 형식
    return {"choices": [{"message": {"role": "assistant", "content": content}}], "usage": {"total_tokens": tokens}}


plan = {"action": "ask", "summary": "s", "question": "q"}

# 1) 형식 지정이 통하는 경우: 토큰·요청 모양·답 해석
replies.append(httpx2.Response(200, json=origin(json.dumps(plan))))
got, tokens = ai.ask(client(), "규칙", [{"role": "user", "content": "요청"}], {"type": "object"})
assert got == plan and tokens == 42, (got, tokens)
token = sent[-1].headers["Authorization"].removeprefix("Bearer ")
assert json.loads(base64.b64decode(token)) == {"apiKey": KEY, "empNo": "123456", "compNo": "30"}
body = json.loads(sent[-1].content)
assert sent[-1].url == URL and body["model"] == "gpt-5.6-terra" and body["need_origin"] is True
assert body["response_format"]["type"] == "json_schema" and body["messages"][0] == {"role": "system", "content": "규칙"}

# 2) 형식 지정을 안 받아 주면(400) 빼고 다시 보내고, 설명이 섞인 답에서 JSON만 읽음. 다른 값에 싸인 응답도 찾음
replies += [httpx2.Response(400, json={"error": "bad"}),
            httpx2.Response(200, json={"result": "ok", "data": origin("계획입니다:\n```json\n" + json.dumps(plan) + "\n```")})]
got, _ = ai.ask(client(), "규칙", [{"role": "user", "content": "요청"}], {"type": "object"})
assert got == plan and "response_format" not in json.loads(sent[-1].content), got

# 3) 실패 안내: 인증, 서버 오류, 사내망 밖
for status, want in ((401, "직번이 올바르지"), (500, "서버 오류(500)")):
    replies[:] = [httpx2.Response(status, json={})]
    try:
        client().chat([{"role": "user", "content": "x"}])
        raise AssertionError("실패해야 함")
    except Exception as e:
        assert want in ai.explain(e), ai.explain(e)
try:
    ai.check_key(KEY, "")
    raise AssertionError("직번 없이 통과하면 안 됨")
except RuntimeError as e:
    assert "직번" in str(e), e
try:
    ai.check_key(KEY, "123456", "ftp://x")
    raise AssertionError("주소 형식이 틀리면 통과하면 안 됨")
except RuntimeError as e:
    assert "주소" in str(e), e
try:
    ai.CompanyClient(KEY, "123456", None)
    raise AssertionError("주소 없이 만들면 안 됨")
except ai.CompanyError as e:
    assert "주소" in ai.explain(e), e
try:  # 이 PC는 사내망 밖: 실제 주소로 접속하면 연결 실패 안내
    ai.check_key(KEY, "123456", URL)
    print("(사내망 안: 실제 회사 서버 응답을 받음)")
except RuntimeError as e:
    assert "사내망" in str(e) or "올바르지" in str(e), e
    print("사내망 밖 안내:", e)
assert not hasattr(ai, "company_log")  # 채팅 창에 진단 기록을 보이지 않는다
print("ok")
