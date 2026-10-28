"""실제 Excel·실제 AI로 확인: 요청 → 적용 확인 → 되돌리기 후 원래와 같은지.
500행 파일의 임시 복사본을 열어 쓰고 닫는다(원본은 건드리지 않음). 사용: python tests/e2e_app.py"""
import os, shutil, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import win32com.client as w
from PySide6.QtWidgets import QApplication, QLabel
from PySide6.QtCore import QTimer
from app.ui import Window, STYLE
from app.excel_ops import fingerprint

tmp = tempfile.mkdtemp()
path = shutil.copy(os.path.join(os.path.dirname(__file__), "fixtures", "거래내역_검증용_500행.xlsx"), os.path.join(tmp, "test_copy.xlsx"))
xl = w.GetActiveObject("Excel.Application")
wb = xl.Workbooks.Open(path)
ws = wb.Worksheets("더미 데이터"); ws.Activate()
orig = fingerprint(ws)

a = QApplication([]); a.setStyleSheet(STYLE)
win = Window(); win.show()
win.file_combo.setCurrentText(wb.Name)
print("대상:", win.target.text())

def check_formula():
    fs = [c.Address for c in ws.Range("H1:Z30").Cells if str(c.Formula2).startswith("=")]
    return f"수식 칸 {len(fs)}개 ({', '.join(fs[:4])})"
def check_sort_filter():
    vis = [r.Row for r in ws.UsedRange.Columns(4).SpecialCells(12).Cells][1:]
    vals = [ws.Cells(r, 4).Value for r in vis]
    nums = [v for v in vals if isinstance(v, float)]
    return (f"보이는 행 {len(vals)}개(글자로 저장된 금액 {len(vals) - len(nums)}개), 숫자는 모두 1억 초과={all(v > 1e8 for v in nums)}, "
            f"숫자 내림차순={nums == sorted(nums, reverse=True)}, 글자 금액: {[v for v in vals if not isinstance(v, float)][:3]}")
def check_top_five():
    rows = [r.Row for r in ws.UsedRange.Columns(4).SpecialCells(12).Cells][1:]
    vals = [ws.Cells(r, 4).Value for r in rows]
    return f"보이는 행 5개={len(vals) == 5}, 내림차순={vals == sorted(vals, reverse=True)}, 값={vals}"
def check_pivot():
    return "시트 목록: " + ", ".join(s.Name for s in wb.Worksheets)

cases = [("회사별 연도별 금액 합계를 수식으로 만들어줘", check_formula),
         ("금액 1억 넘는 것만 보이게 하고 큰 순서로 정렬해줘", check_sort_filter),
         ("금액이 가장 높은 5건만 원본 표에서 필터로 보여줘", check_top_five),
         ("계정명별 금액 합계를 피벗으로 만들어줘", check_pivot)]
results = []
def last():
    return win.scroll.widget().findChildren(QLabel)[-1].text()
def run():
    if not cases:
        wb.Close(SaveChanges=False)
        print("\n".join(results)); a.quit(); return
    req, chk = cases.pop(0)
    win.entry.setText(req); win.send()
    def wait():
        if not win.send_btn.isEnabled():
            QTimer.singleShot(300, wait); return
        msg = last()
        results.append(f"■ {req}\n  답: {msg.splitlines()[0]}")
        if msg.startswith("✔"):
            results.append("  적용 확인: " + chk())
            win.undo()
            results.append("  되돌리기: " + last() + f" / 원래와 같음={fingerprint(ws) == orig} / " + check_pivot())
        else:
            results.append("  전체 답: " + msg)
        run()
    QTimer.singleShot(300, wait)
QTimer.singleShot(800, run)
a.exec()
