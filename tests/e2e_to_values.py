"""실제 Excel로 확인: 수식을 값으로 바꾸기(to_values)가 펼쳐진 수식·일반 수식 모두 보이는 값·형식을 그대로 두고 수식만 없애는지, 되돌리면 수식이 돌아오는지.
별도 Excel 인스턴스에서 새 통합 문서를 만들어 쓰고 닫는다(AI 호출 없음). 사용: python tests/e2e_to_values.py"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import win32com.client as w
from app.excel_ops import Target, Undo, execute, source_changes, validate

xl = w.DispatchEx("Excel.Application")
xl.DisplayAlerts = False
wb = xl.Workbooks.Add()
src, ws = wb.Worksheets(1), wb.Worksheets.Add(After=wb.Worksheets(1))
src.Name = "원본"
src.Range("A1:C1").Value = ("일자", "품목", "가격")
for i in range(2, 8):
    src.Cells(i, 1).Formula = f"=DATE(2025,9,{i})"
    src.Range(src.Cells(i, 2), src.Cells(i, 3)).Value = (f"품목{i}", 1000 * i)
src.Range("A2:A7").NumberFormat = "yyyy-mm-dd"
ws.Range("A1").Formula2 = "=VSTACK('원본'!A1:C1,FILTER('원본'!A2:C7,'원본'!C2:C7>3000))"
ws.Range("A2:A10").NumberFormat = "yyyy-mm-dd"
ws.Range("E1").Value = "두배"
ws.Range("E2:E5").Formula = "=C2*2"
ws.Activate()

try:
    shown = ws.Range("A1:E5").Value
    texts = [ws.Range(a).Text for a in ("A2", "C3", "E2")]
    t = Target(ws)
    plan = {"action": "change", "summary": "", "question": None, "overwrite": False, "edit_source": True, "to_number": [],
            "to_values": ["A1", "E2:E5"], "filters": [], "sorts": [], "new_sheet": "", "formats": [], "highlights": [],
            "pivot": None, "classify": None, "writes": []}
    assert not validate(plan, t.headers), validate(plan, t.headers)
    print("확인 창:", source_changes(t, plan))
    u = Undo(t)
    u.before()
    print(execute(t, plan))
    no_formula = not any(str(c.Formula).startswith("=") for c in ws.Range("A1:E5").Cells)
    ok = [("수식 없음", no_formula),
          ("값 같음", ws.Range("A1:E5").Value == shown),
          ("보이는 글자 같음(날짜 형식 유지)", [ws.Range(a).Text for a in ("A2", "C3", "E2")] == texts)]
    u.after()
    u.undo()
    ok.append(("되돌리면 수식 복원", str(ws.Range("A1").Formula2).startswith("=VSTACK") and str(ws.Range("E2").Formula) == "=C2*2"))
    for k, v in ok:
        print(f"{'✔' if v else '✘'} {k}")
    print(f"{sum(v for _, v in ok)}/{len(ok)} 통과")
finally:
    wb.Close(SaveChanges=False)
    xl.Quit()
