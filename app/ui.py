"""대화창. 상단: Excel 연결·작업 대상, 가운데: 대화, 하단: 입력."""
import threading

from PySide6.QtCore import QObject, Qt, QTimer, Signal
from PySide6.QtWidgets import (QApplication, QComboBox, QFrame, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QScrollArea, QVBoxLayout, QWidget)

from . import ai
from .excel_conn import ExcelError, active_sheet, list_files, read_summary, sheet_info


class Bridge(QObject):
    """작업 스레드의 결과를 화면 스레드로 넘긴다."""
    done = Signal(object, object)  # (결과, 오류 문구)

STYLE = """
QWidget { font-family: 'Malgun Gothic'; font-size: 13px; background: white; }
#status { color: #6b7280; }
#target { background: #f3f4f6; border-radius: 8px; padding: 8px 12px; }
#target[bad="true"] { background: #fcebeb; color: #791f1f; }
#me { background: #e6f1fb; color: #0c447c; border-radius: 12px; padding: 8px 12px; }
#bot { background: #f3f4f6; border-radius: 12px; padding: 10px 12px; }
#err { background: #fcebeb; color: #791f1f; border-radius: 12px; padding: 10px 12px; }
QLineEdit, QComboBox { border: 1px solid #d1d5db; border-radius: 8px; padding: 7px 10px; }
QPushButton { border: 1px solid #d1d5db; border-radius: 8px; padding: 7px 14px; background: white; }
QPushButton:hover { background: #f3f4f6; }
#send { background: #111827; color: white; border: none; }
"""


class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel AI 도우미")
        self.resize(640, 720)
        self.files = []
        self.sheet = None       # 지금 작업 대상 시트 개체
        self.sheet_key = None   # (파일, 시트 이름): 바뀌었을 때만 다시 읽기 위함

        self.status = QLabel("Excel 연결 확인 중…")
        self.status.setObjectName("status")
        refresh = QPushButton("새로고침")
        refresh.clicked.connect(self.refresh)
        top = QHBoxLayout()
        top.addWidget(self.status)
        top.addStretch(1)
        top.addWidget(refresh)

        self.file_combo = QComboBox()
        self.file_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.file_combo.currentIndexChanged.connect(lambda: self.update_sheet(force=True))
        pick = QHBoxLayout()
        pick.addWidget(QLabel("파일"))
        pick.addWidget(self.file_combo, 1)

        # 작업 대상 안내는 대화에 쌓지 않고 이 한 칸을 계속 갱신한다
        self.target = QLabel("")
        self.target.setObjectName("target")
        self.target.setWordWrap(True)

        self.chat = QVBoxLayout()
        self.chat.setAlignment(Qt.AlignTop)
        body = QWidget()
        body.setLayout(self.chat)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setWidget(body)

        self.entry = QLineEdit()
        self.entry.setPlaceholderText("하고 싶은 일을 말로 적어주세요")
        self.entry.returnPressed.connect(self.send)
        self.send_btn = QPushButton("보내기")
        self.send_btn.setObjectName("send")
        self.send_btn.clicked.connect(self.send)
        self.busy = QLabel("")
        self.busy.setObjectName("status")
        bottom = QHBoxLayout()
        bottom.addWidget(self.entry, 1)
        bottom.addWidget(self.send_btn)

        self.planner = None
        self.bridge = Bridge()
        self.bridge.done.connect(self.on_plan)

        root = QVBoxLayout(self)
        root.addLayout(top)
        root.addLayout(pick)
        root.addWidget(self.target)
        root.addWidget(self.scroll, 1)
        root.addWidget(self.busy)
        root.addLayout(bottom)

        # Excel에서 시트를 바꾸면 따라간다 (1초마다 확인, 바뀐 경우에만 다시 읽음)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_sheet)
        self.timer.start(1000)
        self.refresh()

    def say(self, text, kind="bot"):
        label = QLabel(text)
        label.setObjectName(kind)
        label.setWordWrap(True)
        label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        label.setMaximumWidth(500)
        row = QHBoxLayout()  # 줄바꿈 높이가 제대로 계산되도록 정렬은 여백(stretch)으로 처리
        if kind == "me":
            row.addStretch(1)
        row.addWidget(label)
        if kind != "me":
            row.addStretch(1)
        self.chat.addLayout(row)
        QTimer.singleShot(0, lambda: self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum()))

    def show_target(self, text, bad=False):
        self.target.setText(text)
        self.target.setProperty("bad", bad)
        self.target.style().unpolish(self.target)
        self.target.style().polish(self.target)

    def refresh(self):
        """열린 파일을 다시 찾는다. 이전에 고른 파일이 아직 열려 있으면 그대로 유지한다."""
        prev = self.file_combo.currentText()
        self.file_combo.blockSignals(True)
        self.file_combo.clear()
        try:
            self.files = list_files()
        except ExcelError as e:
            self.files, self.sheet, self.sheet_key = [], None, None
            self.status.setText("연결 안 됨")
            self.show_target(str(e), bad=True)
        else:
            self.file_combo.addItems([f["name"] for f in self.files])
            if prev in [f["name"] for f in self.files]:
                self.file_combo.setCurrentText(prev)
            self.status.setText("● Excel 연결됨")
        self.file_combo.blockSignals(False)
        self.update_sheet(force=True)

    def update_sheet(self, force=False):
        """선택한 파일에서 지금 열려 있는 시트를 작업 대상으로 잡는다."""
        i = self.file_combo.currentIndex()
        if i < 0:
            return
        wb = self.files[i]["wb"]
        try:
            ws = active_sheet(wb)
            key = (wb.Name, ws.Name)
            if key == self.sheet_key and not force:
                return
            s = read_summary(ws)
        except ExcelError as e:
            self.sheet, self.sheet_key = None, None
            self.show_target(str(e), bad=True)
            return
        except Exception:
            # Excel이 셀 편집 중이거나 파일이 닫힌 경우: 다음 확인 때 다시 시도
            if force:
                self.show_target("Excel에서 시트를 읽지 못했습니다. 셀 편집 중이면 끝낸 뒤 새로고침을 눌러주세요.", bad=True)
            return
        if self.planner and key != self.sheet_key:
            self.planner.reset()  # 대상이 바뀌면 이전 대화 맥락을 그대로 쓰지 않음 (PRD FR-02)
        self.sheet, self.sheet_key = ws, key
        warn = ""
        if s["protected"]:
            warn += "\n⚠ 보호된 시트라 수정할 수 없습니다."
        if s["readonly"]:
            warn += "\n⚠ 읽기 전용 파일이라 수정할 수 없습니다."
        self.show_target(f"작업 시트: {ws.Name}{warn}")

    def send(self):
        text = self.entry.text().strip()
        if not text or not self.send_btn.isEnabled():
            return
        if self.sheet is None:
            self.say("작업할 시트가 없습니다. Excel에서 파일을 열고 새로고침을 눌러주세요.", "err")
            return
        # 요청 시점의 대상 정보를 여기서 읽어 고정한다 (Excel 개체는 화면 스레드에서만 다룸)
        try:
            info = sheet_info(self.sheet)
        except Exception as e:
            self.say(f"시트를 읽지 못했습니다. 셀 편집 중이면 끝낸 뒤 다시 보내주세요. ({e})", "err")
            return
        if self.planner is None:
            try:
                self.planner = ai.Planner()
            except Exception as e:
                self.say(str(e), "err")
                return
        self.entry.clear()
        self.say(text, "me")
        self.send_btn.setEnabled(False)
        self.busy.setText("AI가 요청을 해석하는 중…")
        target = f"{self.sheet_key[0]} › {self.sheet_key[1]}"

        def work():
            try:
                plan, tokens = self.planner.plan(info, text)
                self.bridge.done.emit((plan, tokens, target), None)
            except Exception as e:
                self.bridge.done.emit(None, ai.explain(e) or f"예상하지 못한 오류: {e}")

        threading.Thread(target=work, daemon=True).start()

    def on_plan(self, res, err):
        self.send_btn.setEnabled(True)
        self.busy.setText("")
        if err:
            self.say(err, "err")
            return
        plan, tokens, target = res
        if plan["action"] == "ask":
            self.say(plan["question"] or plan["summary"])
        elif plan["action"] == "unsupported":
            self.say("지원하지 않는 요청입니다. " + plan["summary"], "err")
        else:
            steps = "\n".join("• " + s for s in ai.describe(plan))
            self.say(f"{plan['summary']}\n\n대상: {target}\n{steps}\n\n(아직 Excel에는 적용하지 않았습니다 · 토큰 {tokens:,})")


def main():
    app = QApplication([])
    app.setStyleSheet(STYLE)
    win = Window()
    win.show()
    app.exec()
