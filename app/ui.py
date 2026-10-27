"""대화창. 상단: Excel 연결·작업 대상, 가운데: 대화, 하단: 입력."""
import glob
import html
import os
import math
import threading
import time

from PySide6.QtCore import QByteArray, QObject, QPointF, QSize, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QFont, QFontDatabase, QIcon, QLinearGradient, QPainter, QPen, QPixmap, QTextCursor
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (QApplication, QComboBox, QDialog, QFrame, QHBoxLayout, QLabel, QLineEdit,
                               QPlainTextEdit, QPushButton, QScrollArea, QStyle, QStyledItemDelegate,
                               QVBoxLayout, QWidget)

from . import ai, classify
from .excel_conn import ExcelError, active_sheet, list_files, read_summary
from .excel_ops import (FormulaError, PlanError, Target, Undo, column_values, excel_error, execute, fingerprint,
                        source_changes, suggest_to_number, validate, write_classification)


ARROW_UP = '<path d="M12 5l0 14"/><path d="M18 11l-6 -6"/><path d="M6 11l6 -6"/>'
KEY = ('<path d="M16.555 3.843l3.602 3.602a2.877 2.877 0 0 1 0 4.069l-2.643 2.643a2.877 2.877 0 0 1 -4.069 0l-.301 -.301l-6.558 6.558a2 2 0 0 1 -1.239 .578l-.175 .008h-1.172a1 1 0 0 1 -.993 -.883l-.007 -.117v-1.172a2 2 0 0 1 .467 -1.284l.119 -.13l.414 -.414h2v-2h2v-2l2.144 -2.144l-.301 -.301a2.877 2.877 0 0 1 0 -4.069l2.643 -2.643a2.877 2.877 0 0 1 4.069 0z"/>'
       '<path d="M15 9h.01"/>')
SHEET = ('<path d="M14 3v4a1 1 0 0 0 1 1h4"/><path d="M17 21h-10a2 2 0 0 1 -2 -2v-14a2 2 0 0 1 2 -2h7l5 5v11a2 2 0 0 1 -2 2z"/>'
         '<path d="M8 11h8v7h-8z"/><path d="M8 15h8"/><path d="M11 11v7"/>')
CHECK = '<path d="M5 12l5 5l10 -10"/>'
ARROW_BACK_UP = '<path d="M9 14l-4 -4l4 -4"/><path d="M5 10h11a4 4 0 1 1 0 8h-1"/>'


def line_icon(paths, color, disabled_color, size=20):
    """선 아이콘을 원하는 색으로 그린다. 글꼴 문자(↑, ↶)보다 선이 고르고 또렷하다."""
    icon = QIcon()
    for mode, c in ((QIcon.Normal, color), (QIcon.Disabled, disabled_color)):
        svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="{c}" '
               f'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">{paths}</svg>')
        pm = QPixmap(size * 2, size * 2)  # 고해상도 화면에서도 또렷하게 2배로 그림
        pm.fill(Qt.transparent)
        painter = QPainter(pm)
        QSvgRenderer(QByteArray(svg.encode())).render(painter)
        painter.end()
        pm.setDevicePixelRatio(2)
        icon.addPixmap(pm, mode)
    return icon


class FileDelegate(QStyledItemDelegate):
    """파일 목록 한 줄: 엑셀 아이콘 + 파일명, 고른 파일에는 체크, 마우스 올린 줄은 둥근 회색 배경."""

    def __init__(self, combo):
        super().__init__(combo)
        self.combo = combo
        self.sheet = line_icon(SHEET, "#3b9a3b", "#3b9a3b", 18).pixmap(18, 18)
        self.check = line_icon(CHECK, "#b4532a", "#b4532a", 16).pixmap(16, 16)

    def sizeHint(self, opt, idx):
        s = super().sizeHint(opt, idx)
        return QSize(s.width() + 64, 38)

    def paint(self, p, opt, idx):
        p.save()
        p.setRenderHint(QPainter.Antialiasing)
        r = opt.rect.adjusted(2, 2, -2, -2)
        if opt.state & (QStyle.State_MouseOver | QStyle.State_Selected):
            p.setPen(Qt.NoPen)
            p.setBrush(QColor("#f1efe8"))
            p.drawRoundedRect(r, 10, 10)
        p.drawPixmap(r.left() + 10, r.center().y() - 8, self.sheet)
        if idx.row() == self.combo.currentIndex():
            p.drawPixmap(r.right() - 26, r.center().y() - 7, self.check)
        text_rect = r.adjusted(38, 0, -36, 0)
        p.setPen(QColor("#1f1f1d"))
        f = QFont(opt.font)
        f.setWeight(QFont.Medium)  # 콤보 칸은 굵게, 목록은 보통 굵기
        p.setFont(f)
        p.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft,
                   opt.fontMetrics.elidedText(idx.data(), Qt.ElideMiddle, text_rect.width()))
        p.restore()


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
    progress = Signal(int, int)    # 분류 진행 (처리, 전체)
    classified = Signal(object, object)

STYLE = """
QWidget { font-family: 'Pretendard', 'Malgun Gothic'; font-size: 15px; background: white; }
#title { font-size: 19px; font-weight: bold; color: #1f1f1d; }
#head { border-bottom: 1px solid #e8e6df; }
#hello { font-size: 17px; color: #1f1f1d; }
#hint { color: #888780; font-size: 14px; }
#pill { border: 1px solid #dcdad3; border-radius: 17px; padding: 7px 16px; min-height: 20px; background: white; color: #1f1f1d; }
#pill:hover { background: #f5f4ef; }
#busy { color: #b4532a; padding: 0 4px; }
#dot { color: #3b9a3b; font-size: 13px; background: transparent; }
#dot[bad="true"] { color: #c0392b; }
#target { color: #5f5e5a; font-size: 14px; background: transparent; }
#target[bad="true"] { color: #a32d2d; }
#me { background: #f1efe8; color: #1f1f1d; border-radius: 14px; padding: 8px 12px; }
#bot { background: white; color: #1f1f1d; border: 1px solid #e8e6df; border-radius: 14px; padding: 10px 12px; }
#err { background: #fcebeb; color: #791f1f; border-radius: 14px; padding: 10px 12px; }
QLineEdit, QComboBox { border: 1px solid #d1d5db; border-radius: 8px; padding: 7px 10px; }
QPushButton { border: 1px solid #d1d5db; border-radius: 8px; padding: 7px 14px; background: white; }
QPushButton:hover { background: #f3f4f6; }
#send { background: #111827; color: white; border: none; }
#composer { border: 1px solid #dcdad3; border-radius: 16px; background: #f5f4ef; }
#ctx { background: transparent; }
#ctx QComboBox { border: none; background: transparent; padding: 2px 4px; font-weight: bold; color: #1f1f1d; }
#filelist { background: white; border: 1px solid #e8e6df; border-radius: 14px; padding: 4px; outline: 0; }
#textbtn { border: none; background: transparent; color: #5f5e5a; padding: 2px 6px; font-size: 14px; }
#textbtn:hover { color: #1f1f1d; background: transparent; }
#inputbox { border: 1px solid #dcdad3; border-radius: 16px; background: white; }
#entry { border: none; padding: 6px 8px; background: transparent; color: #1f1f1d; font-size: 16px; }
#sendicon { background: #1f1f1d; border: none; border-radius: 16px; padding: 0; }
#sendicon:hover { background: #3a3a37; }
#sendicon:disabled { background: #b4b2a9; }
#undoicon { border: none; border-radius: 16px; padding: 0; background: transparent; }
#undoicon:hover { background: #f1efe8; }
#apikey { border: none; background: transparent; color: #5f5e5a; padding: 4px 6px; }
#apikey:hover { color: #1f1f1d; background: transparent; }
#botbox { background: white; border: 1px solid #e8e6df; border-radius: 14px; }
#inner { background: transparent; color: #1f1f1d; border: none; padding: 0; }
#suggest { background: #1f1f1d; color: white; border: none; border-radius: 8px; padding: 7px 14px; font-size: 14px; }
#suggest:hover { background: #3a3a37; }
#suggest:disabled { background: #e8e6df; color: #888780; }
#choice { background: white; color: #1f1f1d; border: 1px solid #dcdad3; border-radius: 8px; padding: 7px 14px; font-size: 14px; }
#choice:hover { background: #f5f4ef; }
#choice:disabled { color: #b4b2a9; border-color: #e8e6df; }
"""


EXAMPLES = ["회사별 연도별 금액 합계를 수식으로 만들어줘", "계정명별 금액 합계를 피벗으로 만들어줘",
            "금액 1억 넘는 것만 보이게 하고 큰 순서로 정렬해줘", "거래 설명을 보고 유형을 분류해줘"]

# 작업 중 멘트: 단계마다 몇 초씩 바꿔 보여준다 (Claude Code의 작업 중 표시처럼)
PHRASES = {
    "AI가 요청을 해석하는 중…": ["요청을 읽는 중…", "표 구조를 살펴보는 중…", "작업 방법을 궁리하는 중…", "수식을 다듬는 중…", "계획을 정리하는 중…"],
    "AI가 수식 오류를 고치는 중…": ["오류 원인을 살펴보는 중…", "수식을 다시 짜는 중…", "검산할 방법을 찾는 중…"],
    "Excel에 적용하는 중…": ["Excel에 적는 중…"],
}

class ChatInput(QPlainTextEdit):
    """입력창: 창이 좁아도 글이 다음 줄로 넘어가고 높이가 최대 6줄까지 늘어난다. Enter=보내기, Shift+Enter=줄바꿈."""
    returnPressed = Signal()

    def __init__(self):
        super().__init__()
        self.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.document().documentLayout().documentSizeChanged.connect(self.fit)
        self.fit()

    def text(self):
        return self.toPlainText()

    def setText(self, s):
        self.setPlainText(s)
        self.moveCursor(QTextCursor.End)

    def fit(self, *_):
        total = int(self.document().size().height())  # 줄바꿈된 줄 수(문서 높이 단위가 줄)
        lines = max(1, min(6, total))
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded if total > 6 else Qt.ScrollBarAlwaysOff)
        h = self.fontMetrics().lineSpacing() * lines + 2 * self.document().documentMargin() + 18
        self.setFixedHeight(int(h))

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Return, Qt.Key_Enter) and not e.modifiers() & Qt.ShiftModifier:
            self.returnPressed.emit()
            return
        super().keyPressEvent(e)


class ShimmerLabel(QLabel):
    """작업 중 문구: 밝은 띠가 글자 위를 왼쪽→오른쪽으로 지나가며 반짝인다(Claude Code 방식). 글이 비면 멈춘다."""
    PERIOD = 1.8  # 띠가 한 번 지나가는 초
    BASE, SHINE = QColor("#b4532a"), QColor("#f0b48f")

    def __init__(self):
        super().__init__("")
        self.t0 = time.monotonic()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update)

    def setText(self, text):
        super().setText(text)
        if text and not self.timer.isActive():
            self.t0 = time.monotonic()
            self.timer.start(30)
        elif not text:
            self.timer.stop()

    def paintEvent(self, e):
        if not self.text():
            return
        p = QPainter(self)
        p.setFont(self.font())
        box = self.rect().adjusted(4, 0, -4, 0)  # 스타일의 좌우 여백 4px
        w = p.fontMetrics().horizontalAdvance(self.text())
        band = max(40.0, w * 0.25)
        phase = ((time.monotonic() - self.t0) / self.PERIOD) % 1
        x = box.left() - band + phase * (w + 2 * band)  # 글자 왼쪽 밖에서 오른쪽 밖까지
        g = QLinearGradient(x - band / 2, 0, x + band / 2, 0)
        g.setColorAt(0, self.BASE)
        g.setColorAt(0.5, self.SHINE)
        g.setColorAt(1, self.BASE)
        p.setPen(QPen(g, 1))
        p.drawText(box, int(Qt.AlignLeft | Qt.AlignVCenter), self.text())


class Spinner(QWidget):
    """작업 중 아이콘: 60도씩 엇갈린 기울어진 궤도 3개를 점 3개가 공전한다.
    앞으로 오면 커지고 진해지고, 뒤로 가면 작아지고 흐려져 깊이감이 난다."""
    PERIOD, TILT, DOT = 2.5, 0.3, 0.11  # 한 바퀴 초, 궤도 기울기(세로/가로), 점 반지름(크기 대비)
    COLOR = (180, 83, 42)  # 작업 중 문구(#b4532a)와 같은 색

    def __init__(self, size=20):
        super().__init__()
        self.setFixedSize(size, size)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.t0 = time.monotonic()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update)

    def start(self):
        if not self.timer.isActive():
            self.t0 = time.monotonic()
            self.timer.start(30)  # 약 33프레임/초: 부드럽게 돌 만큼만
        self.show()

    def stop(self):
        self.timer.stop()
        self.hide()

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(Qt.NoPen)
        s = self.width()
        dot = s * self.DOT
        c, rad = s / 2, s / 2 - dot - 0.5
        spin = (time.monotonic() - self.t0) / self.PERIOD * 2 * math.pi
        pts = []
        for k in range(3):
            rot, a = k * math.pi / 3, spin + k * 2 * math.pi / 3
            ex, ey = math.cos(a) * rad, math.sin(a) * rad * self.TILT
            pts.append((math.sin(a), c + ex * math.cos(rot) - ey * math.sin(rot), c + ex * math.sin(rot) + ey * math.cos(rot)))
        for z, x, y in sorted(pts):  # 뒤에 있는 점부터 그려 앞의 점이 가리게
            f = (z + 1) / 2
            p.setBrush(QColor(*self.COLOR, int(255 * (0.3 + 0.7 * f))))
            r = dot * (0.55 + 0.45 * f)
            p.drawEllipse(QPointF(x, y), r, r)


class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("셀리")
        self.resize(420, 760)
        self.files = []
        self.bubbles = []       # (말풍선, 종류, 가장 긴 줄 폭): 창 크기에 맞춰 폭을 다시 잡는다
        self.suggestion = None  # (시트, [(열, 칸 위치, 칸 수)]): 직전 결과에 붙일 원본 숫자 변환 제안
        self.suggest_btn = None
        self.sheet = None       # 지금 작업 대상 시트 개체
        self.sheet_key = None   # (파일, 시트 이름): 바뀌었을 때만 다시 읽기 위함

        # 머리: 제목 + API KEY
        title = QLabel("셀리")
        title.setObjectName("title")
        self.key_btn = QPushButton("API 키")
        self.key_btn.setObjectName("apikey")
        self.key_btn.setIcon(line_icon(KEY, "#5f5e5a", "#d3d1c7", 16))
        self.key_btn.setCursor(Qt.PointingHandCursor)
        self.key_btn.clicked.connect(self.edit_key)
        head = QFrame()
        head.setObjectName("head")
        top = QHBoxLayout(head)
        top.setContentsMargins(16, 10, 12, 10)
        top.addWidget(title)
        top.addStretch(1)
        top.addWidget(self.key_btn)

        # 대화 영역 (처음에는 예시 요청 버튼)
        self.chat = QVBoxLayout()
        self.chat.setAlignment(Qt.AlignTop)
        self.chat.setContentsMargins(14, 14, 14, 14)
        self.chat.setSpacing(10)
        self.empty = QWidget()
        ev = QVBoxLayout(self.empty)
        ev.setContentsMargins(0, 90, 0, 0)
        ev.setSpacing(10)
        hello = QLabel("이런 요청부터 해보세요:")
        hello.setObjectName("hello")
        ev.addWidget(hello, alignment=Qt.AlignHCenter)
        for ex in EXAMPLES:
            b = QPushButton(ex)
            b.setObjectName("pill")
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda _=False, ex=ex: (self.entry.setText(ex), self.entry.setFocus()))
            ev.addWidget(b, alignment=Qt.AlignHCenter)
        hint = QLabel("Excel에서 파일을 열고 아래에서 파일을 고른 뒤\n하고 싶은 일을 말로 적어주세요")
        hint.setObjectName("hint")
        hint.setAlignment(Qt.AlignHCenter)
        ev.addSpacing(8)
        ev.addWidget(hint, alignment=Qt.AlignHCenter)
        self.chat.addWidget(self.empty)
        body = QWidget()
        body.setLayout(self.chat)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setWidget(body)

        # 작업 중 멘트
        self.busy = ShimmerLabel()
        self.busy.setObjectName("busy")
        self.dots = Spinner()
        self.dots.hide()
        self.busy_msg, self.busy_tick = "", 0
        self.spinner = QTimer(self)
        self.spinner.timeout.connect(self.spin)

        # 입력 묶음: 위는 작업 대상 칩(파일·시트), 아래는 입력창과 버튼
        self.status = QLabel("●")
        self.status.setObjectName("dot")
        self.file_combo = QComboBox()
        self.file_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.file_combo.currentIndexChanged.connect(lambda: self.update_sheet(force=True))
        self.file_combo.setItemDelegate(FileDelegate(self.file_combo))
        self.file_combo.view().setObjectName("filelist")
        popup = self.file_combo.view().window()  # 둥근 모서리 바깥이 네모로 비치지 않게 창 자체를 투명하게
        popup.setWindowFlags(popup.windowFlags() | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
        popup.setAttribute(Qt.WA_TranslucentBackground)
        refresh = QPushButton("새로고침")
        refresh.setObjectName("textbtn")
        refresh.setCursor(Qt.PointingHandCursor)
        refresh.clicked.connect(self.refresh)
        # 작업 대상 안내는 대화에 쌓지 않고 이 한 칸을 계속 갱신한다
        self.target = QLabel("")
        self.target.setObjectName("target")
        self.target.setWordWrap(True)
        ctx = QFrame()
        ctx.setObjectName("ctx")
        cv = QVBoxLayout(ctx)
        cv.setContentsMargins(12, 8, 10, 8)
        cv.setSpacing(2)
        crow = QHBoxLayout()
        crow.addWidget(self.status)
        crow.addWidget(self.file_combo, 1)
        crow.addWidget(refresh)
        cv.addLayout(crow)
        cv.addWidget(self.target)

        self.entry = ChatInput()
        self.entry.setObjectName("entry")
        self.entry.setPlaceholderText("셀리에게 하고 싶은 일을 말로 적어주세요")
        self.entry.returnPressed.connect(self.send)
        self.send_btn = QPushButton()
        self.send_btn.setIcon(line_icon(ARROW_UP, "#ffffff", "#ffffff", 18))
        self.send_btn.setIconSize(QSize(18, 18))
        self.send_btn.setObjectName("sendicon")
        self.send_btn.setFixedSize(32, 32)
        self.send_btn.setToolTip("보내기 (Enter)")
        self.send_btn.setAccessibleName("보내기")
        self.send_btn.setCursor(Qt.PointingHandCursor)
        self.send_btn.clicked.connect(self.send)
        self.undo_btn = QPushButton()
        self.undo_btn.setIcon(line_icon(ARROW_BACK_UP, "#5f5e5a", "#d3d1c7"))
        self.undo_btn.setIconSize(QSize(20, 20))
        self.undo_btn.setObjectName("undoicon")
        self.undo_btn.setFixedSize(32, 32)
        self.undo_btn.setToolTip("직전 AI 작업 되돌리기")
        self.undo_btn.setAccessibleName("직전 AI 작업 되돌리기")
        self.undo_btn.setCursor(Qt.PointingHandCursor)
        self.undo_btn.clicked.connect(self.undo)
        self.undo_btn.setEnabled(False)
        box = QFrame()
        box.setObjectName("inputbox")
        iv = QVBoxLayout(box)
        iv.setContentsMargins(6, 8, 8, 8)
        iv.setSpacing(4)
        iv.addWidget(self.entry)
        irow = QHBoxLayout()
        irow.addWidget(self.undo_btn)
        irow.addStretch(1)
        irow.addWidget(self.send_btn)
        iv.addLayout(irow)
        self.bottombar = QFrame()
        self.bottombar.setObjectName("composer")
        bottom = QVBoxLayout(self.bottombar)
        bottom.setContentsMargins(0, 0, 0, 0)
        bottom.setSpacing(0)
        bottom.addWidget(ctx)
        bottom.addWidget(box)

        self.planner = None
        self.last = None  # 직전 AI 작업의 Undo (1단계만, PRD FR-07)
        self.bridge = Bridge()
        self.bridge.done.connect(self.on_plan)
        self.bridge.progress.connect(self.on_progress)
        self.bridge.classified.connect(self.on_classified)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 12)
        root.setSpacing(0)
        root.addWidget(head)
        root.addWidget(self.scroll, 1)
        pad = QVBoxLayout()
        pad.setContentsMargins(12, 0, 12, 0)
        pad.setSpacing(6)
        busy_row = QHBoxLayout()
        busy_row.setSpacing(6)
        busy_row.addWidget(self.dots)
        busy_row.addWidget(self.busy, 1)
        pad.addLayout(busy_row)
        pad.addWidget(self.bottombar)
        root.addLayout(pad)

        # Excel에서 시트를 바꾸면 따라간다 (1초마다 확인, 바뀐 경우에만 다시 읽음)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_sheet)
        self.timer.start(1000)
        self.refresh()
        self.show_key()

    @staticmethod
    def _rich(text):
        """줄이 길어져 바뀌어도 글머리(•, -) 아래 줄이 글자에 맞춰 들여쓰이도록 HTML로 만든다."""
        out = []
        for line in text.split("\n"):
            body = html.escape(line.strip())
            if "임의 판단" in body or body.startswith("⚠"):  # 셀리가 원본과 다르게 다룬 내용은 놓치지 않게 굵게
                body = f"<b>{body}</b>"
            if line.lstrip().startswith(("•", "-", "→")) and body[:1] in "•-→":
                out.append(f'<p style="margin:0 0 3px 14px; text-indent:-14px">{body}</p>')
            else:
                out.append(f'<p style="margin:0 0 3px 0">{body or "&nbsp;"}</p>')
        return "".join(out)

    def say(self, text, kind="bot", action=None, choices=None):
        """action=(버튼 글, 누르면 할 일): 말풍선 안 맨 아래 [제안] 버튼. 누르는 것 자체가 사용자 동의.
        choices=[(버튼 글, 누르면 할 일), …]: 나란히 놓인 선택 버튼(첫 번째가 진한 버튼). 하나를 누르면 모두 꺼진다."""
        self.empty.hide()  # 대화가 시작되면 예시 버튼은 숨김
        label = QLabel(self._rich(text))
        label.setTextFormat(Qt.RichText)
        label.setObjectName(kind)
        label.setWordWrap(True)
        label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        bubble = label
        if action or choices:  # 말풍선 테두리를 바깥 틀로 옮기고 글과 버튼을 그 안에 함께 넣는다
            bubble = QFrame()
            bubble.setObjectName("botbox")
            label.setObjectName("inner")
            box = QVBoxLayout(bubble)
            box.setContentsMargins(12, 10, 12, 10)
            box.setSpacing(8)
            box.addWidget(label)
        if choices:
            line = QHBoxLayout()
            line.setSpacing(8)
            btns = []
            for i, (name, fn) in enumerate(choices):
                b = QPushButton(name)
                b.setObjectName("suggest" if i == 0 else "choice")
                b.setProperty("full", name)
                b.setCursor(Qt.PointingHandCursor)
                b.clicked.connect(lambda _=False, fn=fn: ([x.setEnabled(False) for x in btns], fn()))
                line.addWidget(b)
                btns.append(b)
            line.addStretch(1)
            box.addLayout(line)
            self.suggest_btn = btns
        if action:
            btn = QPushButton("[제안] " + action[0])
            btn.setObjectName("suggest")
            btn.setCursor(Qt.PointingHandCursor)
            btn.setProperty("full", btn.text())
            btn.setMinimumWidth(80)  # 글이 길어도 말풍선보다 넓어지지 않게(넘치면 fit_bubbles가 … 로 줄임)
            if len(action) > 2:
                btn.setToolTip(action[2])
            btn.clicked.connect(lambda: (btn.setEnabled(False), action[1]()))
            box.addWidget(btn, 0, Qt.AlignLeft)
            self.suggest_btn = [btn]
        fm = label.fontMetrics()
        self.bubbles.append((bubble, kind, max(fm.horizontalAdvance(l.strip()) for l in text.splitlines() or [""])))
        self.fit_bubbles()
        row = QHBoxLayout()  # 줄바꿈 높이가 제대로 계산되도록 정렬은 여백(stretch)으로 처리
        if kind == "me":
            row.addStretch(1)
        row.addWidget(bubble, 0 if kind == "me" else 1)  # 답변 말풍선은 최대 폭까지 늘어남
        self.chat.addLayout(row)
        QTimer.singleShot(0, lambda: self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum()))

    def fit_bubbles(self):
        """말풍선 폭을 창 폭에 맞춘다. 줄바꿈 글상자는 Qt가 폭을 좁게 잡으므로 글이 들어갈 만큼은 넓힌다."""
        for label, kind, longest in self.bubbles:
            maxw = max(200, int((self.width() - 60) * (0.85 if kind == "me" else 1)))
            label.setMaximumWidth(maxw)
            label.setMinimumWidth(min(maxw, longest + 48))
            btn = label.findChild(QPushButton, "suggest")
            if btn:
                full = btn.property("full")
                room = maxw - 24 - 30  # 말풍선 안쪽 여백과 버튼 안쪽 여백
                btn.setText(btn.fontMetrics().elidedText(full, Qt.ElideRight, room))
                btn.setMaximumWidth(maxw - 24)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.fit_bubbles()

    def show_target(self, text, bad=False):
        self.target.setText(text)
        for w in (self.target, self.status):  # 대상 문구와 연결 점을 함께 빨갛게/원래대로
            w.setProperty("bad", bad)
            w.style().unpolish(w)
            w.style().polish(w)

    def refresh(self):
        """열린 파일을 다시 찾는다. 이전에 고른 파일이 아직 열려 있으면 그대로 유지한다."""
        prev = self.file_combo.currentText()
        self.file_combo.blockSignals(True)
        self.file_combo.clear()
        try:
            self.files = list_files()
        except ExcelError as e:
            self.files, self.sheet, self.sheet_key = [], None, None
            self.show_target(str(e), bad=True)
        else:
            self.file_combo.addItems([f["name"] for f in self.files])
            if prev in [f["name"] for f in self.files]:
                self.file_combo.setCurrentText(prev)
            self.status.setToolTip("Excel 연결됨")
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
            self.say(f"시트를 읽지 못했습니다. {excel_error(e)}", "err")
            return
        if self.planner is None:
            try:
                self.planner = ai.Planner()
            except Exception as e:
                self.say(str(e), "err")
                return
        self.entry.clear()
        self.say(text, "me")
        self.ask(text, t, info, before)

    def ask(self, text, t, info, before, retry=False):
        """AI에 계획을 요청한다. retry=True는 수식 오류를 고치는 두 번째 시도(한 번만)."""
        self.set_busy("AI가 수식 오류를 고치는 중…" if retry else "AI가 요청을 해석하는 중…")

        def work():
            try:
                plan, tokens = self.planner.plan(info, text)
                self.bridge.done.emit((plan, tokens, t, before, retry), None)
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
        if msg and not self.busy_msg:
            self.busy_tick, self.busy_start = 0, time.monotonic()
            self.spinner.start(150)
        elif not msg:
            self.spinner.stop()
            self.busy.setText("")
            self.dots.stop()
        self.busy_msg = msg
        if msg:
            self.spin()
        self.send_btn.setEnabled(not msg)
        self.undo_btn.setEnabled(not msg and self.last is not None)

    def spin(self):
        """작업 중 멘트: 파란 점 12개가 돌고, 단계 문구가 3초마다 바뀌며 경과 시간을 보여준다."""
        if not self.busy_msg:
            return
        self.busy_tick += 1
        self.dots.start()
        phrases = PHRASES.get(self.busy_msg, [self.busy_msg])
        secs = int(time.monotonic() - self.busy_start)
        phrase = phrases[(secs // 3) % len(phrases)]
        self.busy.setText(f"{phrase}  ({secs}초)")

    def on_plan(self, res, err):
        self.set_busy("")
        if err:
            self.say(err, "err")
            return
        plan, tokens, t, before, retry = res
        tail = f"\n(토큰 {tokens:,})"
        if plan["action"] == "ask":
            self.say((plan["question"] or plan["summary"]) + tail)
            return
        if plan["action"] == "unsupported":
            self.say("지원하지 않는 요청입니다. " + plan["summary"] + tail, "err")
            return
        bad = validate(plan, t.headers)
        if bad:
            self.say("AI 계획에 문제가 있어 실행하지 않았습니다:\n- " + "\n- ".join(bad), "err")
            return
        if plan["classify"]:
            self.start_classify(plan, t, before, tokens)
            return
        self.finish(plan, t, before, None, tail, retry)

    def start_classify(self, plan, t, before, tokens):
        """분류는 AI 호출이 많아 작업 스레드에서 하고, 끝나면 화면 스레드에서 Excel에 쓴다."""
        c = plan["classify"]
        try:
            values = column_values(t, c["column"])
        except Exception as e:
            self.say(f"'{c['column']}' 열을 읽지 못했습니다. {excel_error(e)}", "err")
            return
        self.set_busy(f"'{c['column']}' 열의 설명을 OpenAI로 보내 분류하는 중…")
        client = self.planner.client

        def work():
            try:
                res = classify.classify(client, values, c["categories"], c["criteria"], plan["summary"],
                                        lambda d, n: self.bridge.progress.emit(d, n))
                self.bridge.classified.emit((plan, t, before, tokens, res, values), None)
            except Exception as e:
                msg = ai.explain(e) or f"분류 중 오류: {e}"
                self.bridge.classified.emit(None, msg + " (Excel은 바꾸지 않았습니다)")

        threading.Thread(target=work, daemon=True).start()

    def on_progress(self, done, total):
        self.busy_msg = f"설명을 읽고 분류하는 중… {done:,}/{total:,}건"

    def on_classified(self, res, err):
        self.set_busy("")
        if err:
            self.say(err, "err")
            return
        plan, t, before, tokens, cls, values = res
        self.finish(plan, t, before, (cls, values), f"\n(토큰 {tokens + cls['tokens']:,})")

    def finish(self, plan, t, before, cls, tail, retry=True):
        fix = None
        try:
            changes = [] if plan.get("confirmed") else source_changes(t, plan)
            if changes:  # 원본 수정은 지시가 있어도 실행 전에 대화 안 버튼으로 한 번 더 확인 (제안 버튼은 누른 것이 확인)
                self.drop_suggestion()
                go = {**plan, "confirmed": True}
                self.say("⚠ 원본 표를 바꾸는 작업입니다. 진행할까요?\n"
                         + "\n".join("• " + c for c in changes[:15])
                         + (f"\n… 외 {len(changes) - 15}곳" if len(changes) > 15 else ""),
                         choices=[("원본 바꾸기", lambda: self.finish(go, t, before, cls, tail, retry)),
                                  ("취소", lambda: self.say("원본 수정을 취소했습니다. Excel은 바꾸지 않았습니다."))])
                return
            text = self.apply(plan, t, before, cls)
            action = None
            if self.suggestion:  # 2차: 원본의 글자 숫자를 숫자로 바꾸자는 제안 (1차 결과는 원본 그대로 계산됨)
                ws, cols = self.suggestion
                n = sum(k for _, _, k in cols)
                where = ", ".join(f"{s} {k}칸" for _, s, k in cols)
                action = (f"원본 글자 숫자 {n}칸을 숫자로 바꾸고 다시 계산", lambda: self.convert(ws, [c for c, _, _ in cols]),
                          f"바꿀 칸: {where}")
            self.say(text + tail, action=action)
        except FormulaError as e:
            if retry:
                self.say(str(e), "err")
            else:
                # 수식이 틀렸으면 이미 원래대로 되돌린 상태: 오류를 알려주고 AI가 한 번만 고쳐 다시 실행
                self.say(f"{e}\n→ AI에게 오류를 알려 수식을 고쳐 다시 실행합니다.", "err")
                fix = f"방금 계획을 실행했더니 수식 오류가 나서 되돌렸다:\n{e}\n같은 요청을 오류가 나지 않는 수식으로 다시 계획하라."
        except PlanError as e:
            self.say(str(e), "err")
        except Exception as e:  # 예상 못 한 Excel 오류에도 창이 '적용 중'에 멈추지 않게
            self.say(excel_error(e), "err")
        finally:
            self.set_busy("")
        if fix:
            try:
                info, before = t.info(), fingerprint(t.ws)
            except Exception as e2:
                self.say(f"시트를 다시 읽지 못해 재시도하지 않았습니다. {excel_error(e2)}", "err")
                return
            self.ask(fix, t, info, before, retry=True)

    def apply(self, plan, t, before, cls=None):
        """대상 재확인 → 되돌리기 사본 → 즉시 실행. 실패하면 원래대로 되돌리고 PlanError."""
        if fingerprint(t.ws) != before:
            raise PlanError("AI가 해석하는 동안 시트 내용이 바뀌어 실행하지 않았습니다. 다시 요청해주세요.")
        self.drop_suggestion()  # 다른 작업을 하면 이전 제안은 더 이상 그 결과에 맞지 않음
        self.set_busy("Excel에 적용하는 중…")
        QApplication.processEvents()
        u = Undo(t)
        self.last = None  # 새 사본을 뜨면 이전 사본은 지워지므로 이전 작업은 더 이상 되돌릴 수 없음
        u.before()
        # shortcut: Excel 적용은 화면 스레드에서 실행해 그동안 창이 잠깐 멈춤, 1만 행 측정 후 필요하면 분리
        t.xl.ScreenUpdating = False
        try:
            done = []
            sugg = suggest_to_number(t, plan)
            if cls:
                res, values = cls
                loc = write_classification(t, plan["classify"]["column"], res["rows"], plan["overwrite"],
                                           plan["classify"]["output"], plan["edit_source"])
                done += classify.report(res, values, loc)
                t = Target(t.ws)  # 분류 열이 붙은 표로 다시 읽어야 이어지는 정렬·필터가 그 열을 포함
            done += execute(t, plan)
        except Exception as e:
            try:
                u.undo(force=True)
                state = "원래 상태로 되돌렸습니다."
            except Exception as e2:
                state = f"자동 복구도 실패했습니다({excel_error(e2)}). 파일을 저장하지 말고 닫은 뒤 다시 여는 것을 권합니다."
            reason = str(e) if isinstance(e, PlanError) else excel_error(e)
            raise (FormulaError if isinstance(e, FormulaError) else PlanError)(f"{reason}\n→ {state}")
        finally:
            t.xl.ScreenUpdating = True
        u.after()
        self.last = u
        self.suggestion = (t.ws, sugg) if sugg else None
        steps = "\n".join("• " + d for d in done)
        head = "⚠ 일부만 분류됨" if cls and cls[0]["failed"] else "✔ 적용 완료"  # 일부 실패를 전체 성공으로 표시하지 않음
        return f"{head}\n{plan['summary']}\n\n대상: {t.wb.Name} › {t.ws.Name}\n{steps}"

    def drop_suggestion(self):
        self.suggestion = None
        for b in self.suggest_btn or []:
            b.setEnabled(False)
        self.suggest_btn = None

    def convert(self, ws, columns):
        """제안 버튼: 원본의 글자 숫자를 숫자로 바꾼다. 1차 결과의 수식은 다시 계산되고 피벗은 새로고침된다."""
        plan = dict(action="change", question=None, overwrite=False, edit_source=True, to_number=columns, filters=[],
                    sorts=[], writes=[], pivot=None, classify=None, new_sheet="", formats=[], confirmed=True,
                    summary="제안에 따라 원본의 글자로 저장된 숫자를 숫자로 바꿨습니다. 앞 결과가 바뀐 값 기준으로 다시 계산되었습니다.")
        try:
            t = Target(ws)
            before = fingerprint(ws)
        except Exception as e:
            self.say(f"시트를 다시 읽지 못했습니다. {excel_error(e)}", "err")
            return
        self.finish(plan, t, before, None, "")

    def undo(self):
        if not self.last:
            return
        self.drop_suggestion()
        try:
            self.last.undo()
        except PlanError as e:
            self.say(str(e), "err")
            return
        except Exception as e:
            self.say(f"되돌리지 못했습니다. {excel_error(e)}", "err")
            return
        self.last = None
        if self.planner:
            self.planner.undone()
        self.set_busy("")
        self.say("↩ 직전 AI 작업을 되돌렸습니다.")


def load_fonts():
    """함께 넣어 둔 프리텐다드를 셀리 안에서만 쓰도록 불러온다(PC에 설치 불필요).
    실패해도 STYLE의 다음 글꼴인 맑은 고딕으로 보인다. exe에는 app/fonts 폴더를 함께 넣어야 한다."""
    for path in glob.glob(os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts", "*.ttf")):
        QFontDatabase.addApplicationFont(path)
    # 힌팅: 획을 픽셀 경계에 맞춰 작은 글자도 또렷하게(끄면 13px에서 흐려 보임). TTF여야 효과가 큼
    f = QApplication.font()
    f.setHintingPreference(QFont.PreferFullHinting)
    f.setWeight(QFont.Medium)  # 본문 기본 굵기: Regular는 작은 글자에서 가늘어 보임
    QApplication.setFont(f)


def main():
    app = QApplication([])
    app.setStyleSheet(STYLE)
    load_fonts()  # 스타일시트 뒤에: 먼저 하면 확인 창·목록 같은 별도 창의 힌팅이 기본값으로 돌아가 흐려짐
    win = Window()
    # Excel 옆에 나란히 보이도록 화면 오른쪽에 세로로 길게 붙인다
    geo = app.primaryScreen().availableGeometry()
    win.resize(420, geo.height() - 16)
    win.move(geo.right() - win.width() - 8, geo.top() + 8)
    win.show()
    app.exec()
