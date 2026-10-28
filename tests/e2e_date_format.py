"""실제 Excel로 확인: 원본 행을 뽑아 온 결과(FILTER·VSTACK)가 날짜를 45901 같은 숫자가 아니라 날짜로 보이는지.
별도 Excel 인스턴스에서 날짜 열이 있는 표를 만들어 쓰고 닫는다. 마지막 한 건은 실제 AI 요청. 사용: python tests/e2e_date_format.py"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import win32com.client as w
from app import ai
from app.excel_ops import Target, execute, validate

xl = w.DispatchEx("Excel.Application")
xl.DisplayAlerts = False
wb = xl.Workbooks.Add()
ws = wb.Worksheets(1)
ws.Name = "가격"
ws.Range("A1:F1").Value = ("일자", "시장명", "품목", "등급", "평균가격", "전일평균가격")
for i in range(2, 12):
    ws.Cells(i, 1).Formula = f"=DATE(2025,9,{i - 1})"
    ws.Range(ws.Cells(i, 2), ws.Cells(i, 6)).Value = ("가락", f"품목{i}", "상", 0 if i % 3 == 0 else 1000 * i, 900 * i)
ws.Range("A2:A11").Value = ws.Range("A2:A11").Value  # 수식 말고 날짜 값으로
ws.Range("A2:A11").NumberFormat = "yyyy-mm-dd"
ws.Range("E2:F11").NumberFormat = "#,##0"
SRC = "'가격'!$A$2:$F$11"


def plan(name, writes):
    return {"action": "change", "summary": "", "question": None, "overwrite": False, "edit_source": False, "to_number": [],
            "filters": [], "sorts": [], "new_sheet": name, "formats": [], "highlights": [], "pivot": None, "classify": None,
            "writes": [{"address": a, "content": c} for a, c in writes]}


def check(label, out, data_row):
    a, e = out.Cells(data_row, 1), out.Cells(data_row, 5)
    ok = a.Text.startswith("2025-09") and "," in e.Text
    print(f"{'✔' if ok else '✘'} {label}: A{data_row}={a.Text!r} ({a.NumberFormat}), E{data_row}={e.Text!r}")
    return ok


try:
    t = Target(ws)
    results = []
    execute(t, plan("머리글따로", [("A1", "일자\t시장명\t품목\t등급\t평균가격\t전일평균가격"),
                                   ("A2", f"=FILTER({SRC},'가격'!$E$2:$E$11<>0)")]))
    results.append(check("FILTER(머리글 윗줄)", wb.Worksheets(2), 2))
    t = Target(ws)
    execute(t, plan("머리글포함", [("A1", "=VSTACK('가격'!$A$1:$F$1,FILTER(" + SRC + ",'가격'!$E$2:$E$11<>0))")]))
    results.append(check("VSTACK(결과 첫 줄이 머리글)", wb.Worksheets(2), 2))

    t = Target(ws)
    p, _ = ai.Planner().plan(t.info(), "E열에서 값이 '0'인 것은 삭제처리해. 새시트에서 작업해")
    assert p["action"] == "change" and not validate(p, t.headers), p
    execute(t, p)
    out = wb.Worksheets(2)
    data_row = 2 if out.Cells(1, 1).Text == "일자" else 3
    results.append(check(f"실제 AI ({p['writes'][-1]['content'][:60]}…)", out, data_row))
    print(f"{sum(results)}/{len(results)} 통과")
finally:
    wb.Close(SaveChanges=False)
    xl.Quit()
