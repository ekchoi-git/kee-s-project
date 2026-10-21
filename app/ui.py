"""대화창. 상단: Excel 연결·작업 대상, 가운데: 대화, 하단: 입력."""
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (QApplication, QComboBox, QFrame, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QScrollArea, QVBoxLayout, QWidget)

from .excel_conn import ExcelError, list_targets, read_summary

STYLE = """
QWidget { font-family: 'Malgun Gothic'; font-size: 13px; background: white; }
#status { color: #6b7280; }
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
        self.targets = []

        top = QHBoxLayout()
        self.status = QLabel("Excel 연결 확인 중…")
        self.status.setObjectName("status")
        self.combo = QComboBox()
        self.combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.combo.currentIndexChanged.connect(self.on_target)
        refresh = QPushButton("새로고침")
        refresh.clicked.connect(self.refresh)
        top.addWidget(self.status)
        top.addWidget(self.combo, 1)
        top.addWidget(refresh)

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
        send = QPushButton("보내기")
        send.setObjectName("send")
        bottom = QHBoxLayout()
        bottom.addWidget(self.entry, 1)
        bottom.addWidget(send)

        root = QVBoxLayout(self)
        root.addLayout(top)
        root.addWidget(self.scroll, 1)
        root.addLayout(bottom)
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

    def refresh(self):
        self.combo.blockSignals(True)
        self.combo.clear()
        try:
            self.targets = list_targets()
        except ExcelError as e:
            self.targets = []
            self.status.setText("연결 안 됨")
            self.say(str(e), "err")
        else:
            self.combo.addItems([t["label"] for t in self.targets])
            files = len({t["wb"].Name for t in self.targets})
            self.status.setText("● Excel 연결됨")
            self.say(f"Excel에서 파일 {files}개, 시트 {len(self.targets)}개를 찾았습니다. 위에서 작업 대상을 고르세요.")
        self.combo.blockSignals(False)
        if self.targets:
            self.on_target()

    def on_target(self):
        i = self.combo.currentIndex()
        if i < 0:
            return
        try:
            s = read_summary(self.targets[i]["ws"])
        except Exception as e:
            self.say(f"이 시트를 읽지 못했습니다: {e}", "err")
            return
        warn = ""
        if s["protected"]:
            warn += "\n⚠ 보호된 시트라 수정할 수 없습니다."
        if s["readonly"]:
            warn += "\n⚠ 읽기 전용 파일이라 수정할 수 없습니다."
        self.say(f"작업 대상: {self.combo.currentText()}\n범위 {s['address']} · 데이터 {s['rows']:,}행\n"
                 f"열 제목: {', '.join(h or '(빈칸)' for h in s['headers'])}{warn}")


def main():
    app = QApplication([])
    app.setStyleSheet(STYLE)
    win = Window()
    win.show()
    app.exec()
