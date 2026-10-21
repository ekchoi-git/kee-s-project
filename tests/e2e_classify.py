"""실제 Excel·실제 AI로 의미 분류 확인: 정답 비교 → 되돌리기 후 원래와 같은지.
500행 파일의 임시 복사본을 열어 쓰고 닫는다(원본은 건드리지 않음). 사용: python tests/e2e_classify.py"""
import os, shutil, sys, tempfile, time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import win32com.client as w
from PySide6.QtWidgets import QApplication, QLabel
from PySide6.QtCore import QTimer
from app.ui import Window, STYLE
from app.excel_ops import fingerprint

tmp = tempfile.mkdtemp()
path = shutil.copy(os.path.join(os.path.dirname(__file__), "..", "거래내역_검증용_500행.xlsx"), os.path.join(tmp, "test_copy.xlsx"))
from app.excel_conn import list_files
# 사용자가 편집 중인 Excel은 피하고 응답 가능한 Excel에서 연다
xl = next(f["wb"].Application for f in list_files() if f["wb"].Application.Ready)
wb = xl.Workbooks.Open(path)
ws = wb.Worksheets("더미 데이터"); ws.Activate()
orig = fingerprint(ws)
truth_rows = wb.Worksheets("행별 정답").UsedRange.Value[1:]
truth = {int(r[0]): (r[3], r[2]) for r in truth_rows}  # 원본행 → (예상 거래유형, 생성 사례)

a = QApplication([]); a.setStyleSheet(STYLE)
win = Window(); win.show()
win.file_combo.setCurrentText(wb.Name)

CATS = "공사 용역, 물류비, 원료 구매, 자재 구매, 인프라 사용료, 임대료, 인력 용역, 교육비"


def score():
    heads = [c.Value for c in ws.UsedRange.Rows(1).Cells]
    col = next(i for i, h in enumerate(heads, 1) if h and str(h).endswith(" 분류"))
    got = {r: ws.Cells(r, col).Value for r in truth}
    ok = sum(got[r] == truth[r][0] for r in truth)
    amb = [r for r in truth if truth[r][0] == "확인 필요"]
    flagged = sum(got[r] == "확인 필요" for r in amb)
    infra_vs_people = sum(1 for r in truth if {truth[r][0], got[r]} == {"인프라 사용료", "인력 용역"})
    extra_check = sum(1 for r in truth if got[r] == "확인 필요" and truth[r][0] != "확인 필요")
    wrong = [(r, truth[r][0], got[r]) for r in truth if got[r] != truth[r][0]][:8]
    return (f"정답 일치 {ok}/{len(truth)} ({ok / len(truth):.1%}), 모호 {len(amb)}건 중 확인 필요 표시 {flagged}건, "
            f"명확한데 확인 필요로 남긴 행 {extra_check}건, 인프라↔인력 혼동 {infra_vs_people}건\n  틀린 예: {wrong}")


def cats_used():
    heads = [c.Value for c in ws.UsedRange.Rows(1).Cells]
    col = next(i for i, h in enumerate(heads, 1) if h and str(h).endswith(" 분류"))
    vals = [ws.Cells(r, col).Value for r in truth]
    return "사용된 범주: " + ", ".join(sorted({str(v) for v in vals}))


cases = [(f"DESCRIPTION을 보고 거래유형을 {CATS} 중 하나로 분류해줘", score),
         ("거래 설명 보고 유형 나눠줘", cats_used)]
results = []


def last():
    return win.scroll.widget().findChildren(QLabel)[-1].text()


def run():
    if not cases:
        wb.Close(SaveChanges=False)
        print("\n".join(results)); a.quit(); return
    req, chk = cases.pop(0)
    win.planner = None  # 사례마다 대화 맥락 없이 시작
    t0 = time.time()
    win.entry.setText(req); win.send()

    def wait():
        if not win.send_btn.isEnabled():
            QTimer.singleShot(500, wait); return
        msg = last()
        results.append(f"■ {req}  ({time.time() - t0:.0f}초)\n{msg}")
        if msg.startswith(("✔", "⚠")):
            try:
                results.append("  정답 확인: " + chk())
            except Exception as e:
                results.append(f"  정답 확인 실패: {e}")
            win.undo()
            results.append("  되돌리기: " + last() + f" / 원래와 같음={fingerprint(ws) == orig}")
        run()
    QTimer.singleShot(500, wait)


QTimer.singleShot(800, run)
a.exec()
