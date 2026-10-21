"""요청 한 번의 흐름: 대상 고정 → AI 계획 → 검사 → 복구 사본 → 즉시 실행 → 결과. 화면과 분리해 테스트 가능."""
import win32com.client as w

from .ai import Planner, explain
from .excel_ops import PlanError, Target, Undo, execute, fingerprint, validate


class Session:
    def __init__(self, planner=None):
        self.connect()
        self.planner = planner or Planner()
        self.last = None  # (Undo, 설명)

    def connect(self):
        try:
            self.xl = w.gencache.EnsureDispatch(w.GetActiveObject("Excel.Application"))
        except Exception:
            raise PlanError("실행 중인 Excel을 찾지 못했습니다. Excel에서 파일을 먼저 열어주세요.")

    def targets(self):
        """열린 통합 문서·시트 목록과 현재 보고 있는 시트."""
        try:
            self.xl.Workbooks.Count
        except Exception:  # Excel을 껐다 켰으면 옛 연결이 끊김
            self.connect()
        items = [(wb.Name, ws.Name) for wb in self.xl.Workbooks for ws in wb.Worksheets if ws.Visible == -1]
        try:
            active = (self.xl.ActiveWorkbook.Name, self.xl.ActiveSheet.Name)
        except Exception:
            active = None
        return items, active

    def send(self, request, wb_name, ws_name):
        """돌려주는 값: (화면에 보여줄 글, 바뀐 게 있는지)."""
        try:
            t = Target(self.xl, wb_name, ws_name)
        except Exception:
            raise PlanError(f"'{wb_name} / {ws_name}'를 찾지 못했습니다. 대상을 새로고침해주세요.")
        t.check()
        before = fingerprint(t.ws)
        try:
            plan, tokens = self.planner.plan(t.info(), request)
        except Exception as e:
            msg = explain(e)
            if msg:
                raise PlanError(msg + " (Excel은 바꾸지 않았습니다)")
            raise
        tail = f"\n(사용 토큰 {tokens:,})"
        if plan["action"] == "ask":
            return (plan["question"] or plan["summary"]) + tail, False
        if plan["action"] == "unsupported":
            return "지원하지 않는 요청입니다. " + plan["summary"] + tail, False
        bad = validate(plan, t.headers)
        if bad:
            raise PlanError("AI 계획에 문제가 있어 실행하지 않았습니다:\n- " + "\n- ".join(bad))
        if fingerprint(t.ws) != before:
            raise PlanError("분석하는 동안 데이터가 바뀌어 실행하지 않았습니다. 다시 요청해주세요.")

        u = Undo(t)
        self.last = None  # 새 사본을 뜨면 이전 사본은 지워지므로 이전 작업은 더 이상 되돌릴 수 없음
        u.before()
        self.xl.ScreenUpdating = False
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
            self.xl.ScreenUpdating = True
        u.after()
        self.last = u
        return plan["summary"] + "\n" + "\n".join("  " + d for d in done) + tail, True

    def undo(self):
        if not self.last:
            raise PlanError("되돌릴 AI 작업이 없습니다.")
        self.last.undo()
        self.last = None
        return "직전 AI 작업을 되돌렸습니다."
