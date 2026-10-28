"""실제 Excel·실제 AI로 확인: Excel에서 다른 파일 창으로 옮기면 셀리 대상도 따라가는지, 없는 열을 말하면 대상 시트를 알려 되묻는지.
별도 Excel 인스턴스에 500행 파일의 임시 복사본 두 개를 열어 쓰고 닫는다(사용자 Excel·원본은 건드리지 않음). 사용: python tests/e2e_follow_active.py"""
import os, shutil, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import win32com.client as w
from PySide6.QtWidgets import QApplication
from app import ai, ui
from app.excel_ops import Target

src = os.path.join(os.path.dirname(__file__), "fixtures", "거래내역_검증용_500행.xlsx")
tmp = tempfile.mkdtemp()
xl = w.DispatchEx("Excel.Application")
xl.DisplayAlerts = False
xl.Visible = True  # 숨긴 Excel에서는 Activate()로 활성 파일이 바뀌지 않음
a_wb = xl.Workbooks.Open(shutil.copy(src, os.path.join(tmp, "follow_a.xlsx")))
b_wb = xl.Workbooks.Open(shutil.copy(src, os.path.join(tmp, "follow_b.xlsx")))
# 셀리가 이 시험용 Excel의 파일만 보게 함(사용자가 열어 둔 Excel 파일은 목록에서 뺌)
ui.list_files = lambda: sorted(({"wb": wb, "name": wb.Name} for wb in xl.Workbooks),
                               key=lambda f: f["name"] != xl.ActiveWorkbook.Name)
try:
    app = QApplication([])
    win = ui.Window()
    win.follow_active()
    print("처음:", win.file_combo.currentText(), "→", win.file_combo.currentText() == "follow_b.xlsx")
    win.file_combo.setCurrentText("follow_a.xlsx")  # 셀리에서 다른 파일을 직접 고름
    win.follow_active()
    print("직접 고른 파일 유지:", win.file_combo.currentText() == "follow_a.xlsx")
    win.file_combo.setCurrentText("follow_b.xlsx")
    a_wb.Activate()  # Excel에서 다른 파일 창으로 옮김
    win.follow_active()
    print("Excel 파일 변경 따라감:", win.file_combo.currentText() == "follow_a.xlsx", "/", win.target.text())

    t = Target(a_wb.Worksheets(1))
    plan, _ = ai.Planner().plan(t.info(), "연도별로 b~k열까지 평균을 새 시트에 구해줘")
    print("없는 열 요청:", plan["action"], "/", plan["question"])
finally:
    a_wb.Close(SaveChanges=False)
    b_wb.Close(SaveChanges=False)
    xl.Quit()
