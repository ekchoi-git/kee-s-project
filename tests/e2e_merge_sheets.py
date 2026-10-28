"""실제 Excel·실제 AI로 확인: 여러 시트(2개, 6개) 합치기가 새 시트에 모든 행을 쌓는지, 되묻기가 이어져도 처음 요청을 기억하는지.
철강산업 파일의 임시 복사본을 열어 쓰고 닫는다(원본은 건드리지 않음). 사용: python tests/e2e_merge_sheets.py"""
import os, shutil, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import win32com.client as w
from app import ai
from app.excel_ops import Target, validate, execute

src = os.path.join(os.path.dirname(__file__), "..", "산업통상부_철강산업 동향_20241231.xlsx")
path = shutil.copy(src, os.path.join(tempfile.mkdtemp(), "merge_copy.xlsx"))
xl = w.GetActiveObject("Excel.Application")
wb = xl.Workbooks.Open(path)


def rows(ws):
    return Target(ws).region.Rows.Count - 1


def run(req, answers=()):
    ws = wb.Worksheets(1)
    ws.Activate()
    t = Target(ws)
    p = ai.Planner()
    plan, _ = p.plan(t.info(), req)
    for a in answers:  # 되물으면 사용자 답을 차례로 준다
        if plan["action"] != "ask":
            break
        print("  AI 질문:", plan["question"], "→", a)
        plan, _ = p.plan(t.info(), a)
    assert plan["action"] == "change", plan
    assert not validate(plan, t.headers), validate(plan, t.headers)
    print("  수식:", plan["writes"][0]["content"][:200])
    execute(t, plan)
    assert plan["new_sheet"], plan
    return wb.Worksheets(2)  # 새 시트는 작업 시트 바로 뒤에 생김(이름이 겹치면 (2)가 붙음)


try:
    names = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
    print("시트:", names)
    want = rows(wb.Worksheets(1)) + rows(wb.Worksheets(2))
    out = run("현재 시트와 2번째 시트를 새 시트에 통합해줘")
    got = out.UsedRange.Rows.Count - 1
    print(f"■ 2개: 새 시트 '{out.Name}' 데이터 {got}행 / 기대 {want}행 → {got == want}")

    for _ in range(4):  # 같은 구조 시트를 늘려 6개로
        wb.Worksheets(names[1]).Copy(After=wb.Worksheets(wb.Worksheets.Count))
    # 원본 2개 + 복사본 4개. 이미 합친 '통합'과 구조가 다른 '연도별 명목소비'는 빠져야 함
    want = rows(wb.Worksheets(names[0])) + 5 * rows(wb.Worksheets(names[1]))
    out = run("모든 시트를 새 시트 하나에 합쳐줘", answers=["같은 구조 시트 전부"])
    got = out.UsedRange.Rows.Count - 1
    print(f"■ 6개: 새 시트 '{out.Name}' 데이터 {got}행 / 기대 {want}행 → {got == want}")
finally:
    wb.Close(SaveChanges=False)
