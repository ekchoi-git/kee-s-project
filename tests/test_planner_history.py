"""AI가 여러 번 되물어도 처음 요청이 AI에 보내는 기록에서 잘려 나가지 않는지(가짜 AI로 확인). 사용: python tests/test_planner_history.py"""
import json, os, sys, types
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app import ai

sent = []
def fake(actions):
    def create(messages, **_):
        sent.append(messages)
        msg = types.SimpleNamespace(content=json.dumps({"action": actions.pop(0), "question": "?"}))
        return types.SimpleNamespace(choices=[types.SimpleNamespace(message=msg)], usage=types.SimpleNamespace(total_tokens=1))
    return create

p = ai.Planner.__new__(ai.Planner)
p.reset()
p.client = types.SimpleNamespace(chat=types.SimpleNamespace(completions=types.SimpleNamespace(create=fake(["ask"] * 6 + ["change"] * 11))))
for req in ["시트를 새 시트에 통합해줘", "답1", "답2", "답3", "답4", "답5", "답6"]:
    p.plan("info", req)
assert any(m["content"] == "시트를 새 시트에 통합해줘" for m in sent[-1]), "6번 되물은 뒤 처음 요청이 빠짐"
for i in range(10):  # 되묻기가 끝나면 다시 최근 8개만
    p.plan("info", f"다른 요청{i}")
assert len(sent[-1]) == 1 + 8 + 1 and p.asking is None
print("ok")
