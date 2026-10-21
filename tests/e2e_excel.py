"""실제 Excel로 하는 종단 테스트. 500행 파일의 임시 복사본만 쓰고 원본은 건드리지 않는다.

python tests/e2e_excel.py        # 정해진 계획으로 실행·정답·되돌리기 확인
python tests/e2e_excel.py --api  # 위 + 실제 ChatGPT로 한국어 요청 처리
"""
import os, shutil, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import openpyxl
import win32com.client as w
from app.core import Session
from app.excel_ops import PlanError, ERR_BASE

SRC = os.path.join(os.path.dirname(__file__), "..", "거래내역_검증용_500행.xlsx")


class FakePlanner:
    def __init__(self):
        self.next = None
    def plan(self, info, request):
        return self.next, 0
    def reset(self):
        pass


def plan(**kw):
    p = {"action": "change", "summary": "테스트", "question": None, "overwrite": False,
         "filters": [], "sorts": [], "writes": [], "pivot": None}
    p.update(kw)
    return p


def state(ws):
    r = ws.UsedRange
    return r.Address, r.Formula2, r.Value2, ws.ListObjects(1).ShowAutoFilter and ws.ListObjects(1).AutoFilter.FilterMode, [s.Name for s in ws.Parent.Worksheets]


def answers():
    wb = openpyxl.load_workbook(SRC, read_only=True)
    return {(r[0], r[1]): r[3] for r in wb["검증 정답"].iter_rows(values_only=True)
            if isinstance(r[1], int) and 2000 < r[1] < 2100 and isinstance(r[3], int)}


def main():
    tmp = os.path.join(tempfile.mkdtemp(), "e2e_500.xlsx")
    shutil.copy(SRC, tmp)
    xl = w.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = True
    wb = xl.Workbooks.Open(tmp)
    ws = wb.Worksheets("더미 데이터")
    fake = FakePlanner()
    s = Session(planner=fake)
    name = wb.Name, ws.Name
    fails = []

    def check(label, cond):
        print(("OK  " if cond else "실패") + " " + label)
        if not cond:
            fails.append(label)

    def run(p, req="테스트"):
        fake.next = p
        return s.send(req, *name)

    try:
        orig = state(ws)

        run(plan(writes=[{"address": "J1", "content": '=SUMIFS($D$2:$D$501,$F$2:$F$501,"A상사",$B$2:$B$501,">="&DATE(2022,1,1),$B$2:$B$501,"<"&DATE(2023,1,1))'}]))
        check("수식: A상사 2022 합계가 정답(3,559,197,769)과 같음", ws.Range("J1").Value2 == 3559197769)
        check("수식: 셀에 실제 수식이 들어 있음", ws.Range("J1").HasFormula)
        s.undo(); check("수식 되돌리기", state(ws) == orig)

        msg, _ = run(plan(filters=[{"column": "INVOICE_AMOUNT", "op": ">", "value": "1000000000"}],
                          sorts=[{"column": "INVOICE_AMOUNT", "descending": True}]))
        vals = [v for v in ws.Range("D2:D501").Value2 if isinstance(v[0], float)]
        check("필터+정렬: 10억 초과 10행만 보임 (정답 10)", "보이는 행 10개" in msg)
        first_num = next(v[0] for v in ws.Range("D2:D501").Value2 if isinstance(v[0], float))
        check("정렬: 숫자 중 첫 값이 가장 큰 금액", first_num == max(v[0] for v in vals))
        check("정렬: 글자로 된 숫자 10칸 경고 표시", "글자로 저장된 칸이 10개" in msg)
        s.undo(); check("필터+정렬 되돌리기", state(ws) == orig)

        run(plan(pivot={"rows": ["VENDOR_NAME"], "columns": [], "values": [{"field": "INVOICE_AMOUNT", "func": "sum"}]}))
        pws = wb.Worksheets("AI 피벗1")
        check("피벗: 새 시트에 실제 피벗 테이블 생성", pws.PivotTables().Count == 1)
        s.undo(); check("피벗 되돌리기(시트 삭제)", state(ws) == orig)

        try:
            run(plan(writes=[{"address": "J1", "content": '=WEBSERVICE("http://example.com")'}]))
            check("외부 연결 수식 차단", False)
        except PlanError:
            check("외부 연결 수식 차단", state(ws) == orig)

        try:
            run(plan(writes=[{"address": "J1", "content": "=없는함수(1)"}]))
            check("잘못된 수식 → 실패로 표시", False)
        except PlanError as e:
            check("잘못된 수식 → 실패로 표시하고 자동 복구", "되돌렸습니다" in str(e) and state(ws) == orig)

        try:
            run(plan(writes=[{"address": "D2", "content": "0"}]))
            check("원본 덮어쓰기 거부", False)
        except PlanError:
            check("원본 덮어쓰기 거부(덮어쓰기 지시 없을 때)", state(ws) == orig)

        run(plan(writes=[{"address": "J1", "content": "AI 결과"}]))
        ws.Range("J2").Value2 = "사용자가 직접 입력"
        try:
            s.undo(); check("직접 수정 후 되돌리기 거부", False)
        except PlanError:
            check("직접 수정 후 되돌리기 거부", ws.Range("J2").Value2 == "사용자가 직접 입력")
        ws.Range("J1:J2").Clear()

        if "--api" in sys.argv:
            from app.ai import Planner
            s.planner = Planner()
            orig = state(ws)
            msg, changed = s.send("회사별 연도별 금액 합계를 수식으로 만들어줘", *name)
            print("\n[AI 답]", msg, "\n")
            ans = answers()
            used = ws.UsedRange.Value2
            found = {v for row in used for v in row if isinstance(v, float)}
            hit = sum(v in found for v in ans.values())
            check(f"실제 AI: 회사·연도 합계 40개 중 {hit}개가 정답과 일치", changed and hit == 40)
            formulas = sum(1 for row in ws.UsedRange.Formula2 for f in row if isinstance(f, str) and f.startswith("="))
            check(f"실제 AI: 결과가 실제 수식임 (수식 {formulas}칸)", formulas > 0)
            msg, changed = s.send("여기서 금액 큰 순서로 원본 표를 정렬해줘", *name)
            print("\n[AI 답]", msg, "\n")
            check("실제 AI 후속 요청: 정렬 적용", changed and "정렬" in msg)
            s.undo()
            msg, _ = s.send("이 파일 지우고 바탕화면 다른 파일 열어줘", *name)
            print("\n[AI 답]", msg, "\n")
            check("실제 AI: 허용되지 않은 요청은 실행 안 함", "지원하지 않는" in msg)
    finally:
        wb.Close(SaveChanges=False)
    print("\n전체 결과:", "모두 통과" if not fails else f"{len(fails)}개 실패")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
