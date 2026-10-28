"""Claude 키로 바꿔도 같은 계획 흐름이 도는지(가짜 Claude 서버 응답으로 확인), 키 종류 구분, 실제 서버의 잘못된 키 안내.
사용: python tests/test_claude_provider.py"""
import json, os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import anthropic
import httpx2
from app import ai

assert ai.provider("sk-ant-api03-x") == "Claude" and ai.provider("sk-proj-x") == "OpenAI"
assert isinstance(ai.connect("sk-ant-api03-x"), anthropic.Anthropic)

sent = []
def reply(request):
    body = json.loads(request.content)
    sent.append((request.headers.get("anthropic-beta"), body))
    plan = {"action": "ask", "question": "어느 열인가요?"}
    return httpx2.Response(200, json={
        "id": "msg_1", "type": "message", "role": "assistant", "model": body["model"],
        "content": [{"type": "text", "text": json.dumps(plan, ensure_ascii=False)}],
        "stop_reason": "end_turn", "stop_sequence": None,
        "usage": {"input_tokens": 100, "output_tokens": 20},
    })

p = ai.Planner.__new__(ai.Planner)
p.reset()
p.client = anthropic.Anthropic(api_key="sk-ant-api03-x", http_client=httpx2.Client(transport=httpx2.MockTransport(reply)))
plan, tokens = p.plan("info", "첫 요청")
plan, tokens = p.plan("info", "답")
beta, body = sent[-1]
assert plan["question"] == "어느 열인가요?" and tokens == 120
assert body["system"] == ai.SYSTEM and [m["role"] for m in body["messages"]] == ["user", "assistant", "user"]
assert body["output_config"]["format"]["schema"] == ai.PLAN_SCHEMA and body["fallbacks"] == "default"
assert "server-side-fallback-2026-07-01" in beta

try:  # 실제 Claude 서버: 잘못된 키는 인증 실패 안내로
    ai.check_key("sk-ant-api03-invalid")
    raise AssertionError("잘못된 키가 통과함")
except RuntimeError as e:
    assert "인증 실패" in str(e), e
print("ok")
