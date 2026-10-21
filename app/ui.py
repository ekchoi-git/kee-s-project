"""대화창. 상단: Excel 연결·작업 대상, 가운데: 대화, 하단: 입력."""
import threading

from PySide6.QtCore import QObject, Qt, QTimer, Signal
from PySide6.QtWidgets import (QApplication, QComboBox, QDialog, QFrame, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QScrollArea, QVBoxLayout, QWidget)

from . import ai
from .excel_conn import ExcelError, active_sheet, list_files, read_summary
from .excel_ops import PlanError, Target, Undo, execute, fingerprint, validate


class KeyDialog(QDialog):
    """AI API 키 입력. 화면에서는 가리고, 실제 접속으로 확인한 뒤 Windows 자격 증명 관리자에 저장 (PRD FR-08)."""

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("AI API 키")
        self.setMinimumWidth(420)
        now = ai.current_key()
        info = QLabel(f"현재 키: {ai.masked(now) if now else '없음'}\n"
                      "입력한 키는 이 PC의 Windows 자격 증명 관리자에 암호화되어 저장됩니다.")
        info.setWordWrap(True)
        self.edit = QLineEdit()
        self.edit.setEchoMode(QLineEdit.Password)
        self.edit.setPlaceholderText("sk-로 시작하는 OpenAI API 키")
        self.msg = QLabel("")
        self.msg.setWordWrap(True)
        save, remove, cancel = QPushButton("확인 후 저장"), QPushButton("저장된 키 삭제"), QPushButton("닫기")
        save.setObjectName("send")
        save.clicked.connect(self.save)
        remove.clicked.connect(self.remove)
        cancel.clicked.connect(self.reject)
        btns = QHBoxLayout()
        btns.addWidget(remove)
        btns.addStretch(1)
        btns.addWidget(cancel)
        btns.addWidget(save)
        lay = QVBoxLayout(self)
        for wdg in (info, self.edit, self.msg):
            lay.addWidget(wdg)
        lay.addLayout(btns)

    def save(self):
        k = self.edit.text().strip()
        if not k:
            self.msg.setText("키를 입력해주세요.")
            return
        self.msg.setText("OpenAI에 접속해 키를 확인하는 중…")
        QApplication.processEvents()
        try:
            ai.check_key(k)
        except Exception as e:
            self.msg.setText(f"저장하지 않았습니다. {e}")
            return
        ai.save_key(k)
        self.accept()

    def remove(self):
        ai.delete_key()
        self.msg.setText("저장된 키를 삭제했습니다.")


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
#apikey { background: #2f3ce8; color: white; border: none; border-radius: 16px; padding: 7px 26px;
          font-weight: bold; letter-spacing: 1px; }
#apikey:hover { background: #2531c4; }
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
        self.key_btn = QPushButton("API KEY")
        self.key_btn.setObjectName("apikey")
        self.key_btn.setCursor(Qt.PointingHandCursor)
        self.key_btn.clicked.connect(self.edit_key)
        top = QHBoxLayout()
        top.addWidget(self.status)
        top.addStretch(1)
        top.addWidget(self.key_btn)
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
        self.undo_btn = QPushButton("직전 AI 작업 되돌리기")
        self.undo_btn.clicked.connect(self.undo)
        self.undo_btn.setEnabled(False)
        bottom = QHBoxLayout()
        bottom.addWidget(self.entry, 1)
        bottom.addWidget(self.send_btn)
        bottom.addWidget(self.undo_btn)

        self.planner = None
        self.last = None  # 직전 AI 작업의 Undo (1단계만, PRD FR-07)
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
        self.show_key()

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
        # 요청 시점의 대상을 여기서 고정한다 (Excel 개체는 화면 스레드에서만 다룸)
        try:
            t = Target(self.sheet)
            t.check()
            info, before = t.info(), fingerprint(t.ws)
        except PlanError as e:
            self.say(str(e), "err")
            return
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
        self.set_busy("AI가 요청을 해석하는 중…")

        def work():
            try:
                plan, tokens = self.planner.plan(info, text)
                self.bridge.done.emit((plan, tokens, t, before), None)
            except Exception as e:
                msg = ai.explain(e) or f"예상하지 못한 오류: {e}"
                self.bridge.done.emit(None, msg + " (Excel은 바꾸지 않았습니다)")

        threading.Thread(target=work, daemon=True).start()

    def show_key(self):
        k = ai.current_key()
        self.key_btn.setToolTip(f"현재 키: {ai.masked(k)}" if k else "API 키를 입력해주세요")

    def edit_key(self):
        if KeyDialog(self).exec():
            self.planner = None  # 새 키로 다시 연결 (대화 맥락도 새로 시작)
            self.say("AI 키를 확인하고 저장했습니다.")
        self.show_key()

    def set_busy(self, msg):
        self.busy.setText(msg)
        self.send_btn.setEnabled(not msg)
        self.undo_btn.setEnabled(not msg and self.last is not None)

    def on_plan(self, res, err):
        self.set_busy("")
        if err:
            self.say(err, "err")
            return
        plan, tokens, t, before = res
        tail = f"\n(토큰 {tokens:,})"
        if plan["action"] == "ask":
            self.say((plan["question"] or plan["summary"]) + tail)
            return
        if plan["action"] == "unsupported":
            self.say("지원하지 않는 요청입니다. " + plan["summary"] + tail, "err")
            return
        try:
            self.say(self.apply(plan, t, before) + tail)
        except PlanError as e:
            self.say(str(e), "err")
        except Exception as e:  # 예상 못 한 Excel 오류에도 창이 '적용 중'에 멈추지 않게
            self.say(f"Excel 작업 중 오류가 났습니다: {e}", "err")
        finally:
            self.set_busy("")

    def apply(self, plan, t, before):
        """검사 → 되돌리기 사본 → 즉시 실행. 실패하면 원래대로 되돌리고 PlanError."""
        bad = validate(plan, t.headers)
        if bad:
            raise PlanError("AI 계획에 문제가 있어 실행하지 않았습니다:\n- " + "\n- ".join(bad))
        if fingerprint(t.ws) != before:
            raise PlanError("AI가 해석하는 동안 시트 내용이 바뀌어 실행하지 않았습니다. 다시 요청해주세요.")
        self.set_busy("Excel에 적용하는 중…")
        QApplication.processEvents()
        u = Undo(t)
        self.last = None  # 새 사본을 뜨면 이전 사본은 지워지므로 이전 작업은 더 이상 되돌릴 수 없음
        u.before()
        # shortcut: Excel 적용은 화면 스레드에서 실행해 그동안 창이 잠깐 멈춤, 1만 행 측정 후 필요하면 분리
        t.xl.ScreenUpdating = False
        try:
            done = execute(t, plan)
        except Exception as e:
            try:
                u.undo(force=True)
                state = "원래 상태로 되돌렸습니다."
            except Exception as e2:
                state = f"자동 복구도 실패했습니다({e2}). 파일을 저장하지 말고 닫은 뒤 다시 여는 것을 권합니다."
            reason = str(e) if isinstance(e, PlanError) else f"Excel 실행 오류: {e}"
            raise PlanError(f"{reason}\n→ {state}")
        finally:
            t.xl.ScreenUpdating = True
        u.after()
        self.last = u
        steps = "\n".join("• " + d for d in done)
        return f"✔ 적용 완료\n{plan['summary']}\n\n대상: {t.wb.Name} › {t.ws.Name}\n{steps}"

    def undo(self):
        if not self.last:
            return
        try:
            self.last.undo()
        except PlanError as e:
            self.say(str(e), "err")
            return
        except Exception as e:
            self.say(f"되돌리기 중 Excel 오류: {e}", "err")
            return
        self.last = None
        if self.planner:
            self.planner.undone()
        self.set_busy("")
        self.say("↩ 직전 AI 작업을 되돌렸습니다.")


def main():
    app = QApplication([])
    app.setStyleSheet(STYLE)
    win = Window()
    win.show()
    app.exec()
