"""1단계 점검: 실행 중인 Excel에 연결해 값/수식 쓰기와 피벗 생성이 되는지 확인."""
import win32com.client as w

try:
    xl = w.GetActiveObject("Excel.Application"); xl = w.gencache.EnsureDispatch(xl)
except Exception:
    xl = w.gencache.EnsureDispatch("Excel.Application")  # shortcut: Excel이 안 떠 있으면 새로 띄움
    xl.Visible = True
    xl.Workbooks.Add()

wb = xl.ActiveWorkbook
ws = wb.Worksheets.Add()
ws.Range("A1:B1").Value = ("회사", "금액")
rows = [("가", 10), ("나", 20), ("가", 30)]
for i, r in enumerate(rows, start=2):
    ws.Range(f"A{i}:B{i}").Value = r

ws.Range("D1").Formula2 = '=SUMIFS(B2:B4,A2:A4,"가")'
assert ws.Range("D1").Value == 40, ws.Range("D1").Value
print("수식 OK:", ws.Range("D1").Value)

pc = wb.PivotCaches().Create(1, ws.Range("A1:B4"))
pt = pc.CreatePivotTable(ws.Range("F1"), "PT1")
pt.PivotFields("회사").Orientation = 1
pt.AddDataField(pt.PivotFields("금액"), "합계", -4157)
print("피벗 OK:", ws.PivotTables().Count, "개")
print("Excel 버전:", xl.Version, "/ 시트:", ws.Name)

ws.Range("H1").Formula2 = "=UNIQUE(A2:A4)"
assert [c.Value for c in ws.Range("H1:H2")] == ["가", "나"], ws.Range("H1:H3").Value
print("최신 함수(UNIQUE) OK, Build:", xl.Build)
