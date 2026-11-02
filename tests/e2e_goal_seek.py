"""실제 Excel로 확인: 목표값 찾기(goal_seek)가 흩어진 손익계산서 양식에서 답을 찾는지, 엉뚱한 칸은 실행 전에 막는지,
답이 없으면 원래 값으로 돌리는지, 되돌리면 입력 칸이 원래대로 오는지. 시트 정보에 표 밖 칸(D20)이 들어가는지.
별도 Excel 인스턴스에서 새 통합 문서를 만들어 쓰고 닫는다(AI 호출 없음). 사용: python tests/e2e_goal_seek.py"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import win32com.client as w
from app.excel_ops import PlanError, Target, Undo, execute, validate

CELLS = {
    "B2": "카페 월 손익계산서", "B4": "일 판매량(잔)", "C4": 250, "B5": "평균 가격(원)", "C5": 4500,
    "B6": "일 매출", "C6": "=C4*C5", "B7": "월 매출 합계", "D7": "=C6*22+C6*8*40%",
    "B10": "임대료", "C10": 6000000, "B11": "인건비", "C11": "=F11*2500000", "E11": "직원수", "F11": "=ROUNDDOWN(C4/300,0)+2",
    "B12": "재료비", "C12": "=D7*35%", "B13": "로열티", "C13": "=D7*4%", "B14": "비용 합계", "D14": "=SUM(C10:C13)",
    "B16": "초기투자비용", "C16": 400000000, "B17": "상각 기간(년)", "C17": 5, "B18": "감가상각비(월)", "D18": "=C16/C17/12",
    "B20": "영업이익", "D20": "=D7-D14-D18", "H4": "참고: 작년 영업이익", "I4": 1200000,
}
# (결과 칸, 목표, 바꿀 칸, 기대 답 또는 None=막혀야 함)  답은 검증용 파일에서 Excel로 계산한 값
CASES = [("D20", 0, "C4", 255.39), ("D20", 2000000, "C4", 284.31), ("D20", 5000000, "C4", 363.82),
         ("D20", 0, "C16", 377610000), ("D20", 3000000, "C5", 5377.74),
         ("D20", 50000000, "C17", None), ("D20", 0, "I4", None), ("D20", 0, "C6", None), ("C10", 0, "C4", None)]

xl = w.DispatchEx("Excel.Application")
xl.DisplayAlerts = False
wb = xl.Workbooks.Add()
ws = wb.Worksheets(1)
for a, v in CELLS.items():
    ws.Range(a).Value = v

try:
    ok = [("시트 정보에 표 밖 영업이익 칸", "D20: =D7-D14-D18" in Target(ws).info())]
    for tgt, goal, chg, want in CASES:
        t = Target(ws)
        plan = {"action": "change", "summary": "", "question": None, "overwrite": False, "edit_source": False, "to_number": [],
                "to_values": [], "filters": [], "sorts": [], "new_sheet": "", "formats": [], "highlights": [],
                "pivot": None, "classify": None, "writes": [], "goal_seek": {"target": tgt, "goal": goal, "change": chg}}
        assert not validate(plan, t.headers), validate(plan, t.headers)
        old = ws.Range(chg).Value2
        u = Undo(t)
        u.before()
        try:
            msg = execute(t, plan)
            got = ws.Range(chg).Value2
            u.after()
            u.undo()
            back = ws.Range(chg).Value2 == old
            ok.append((f"{chg}→{tgt}={goal:,}: {got:,.2f} (기대 {want}), 되돌리기 {back}",
                       want is not None and abs(got - want) <= abs(want) * 1e-4 and back))
            print("  " + "\n  ".join(msg))
        except PlanError as e:
            ok.append((f"{chg}→{tgt}={goal:,}: 막힘, 값 그대로 {ws.Range(chg).Value2 == old}",
                       want is None and ws.Range(chg).Value2 == old))
            print("  " + str(e))
            u.undo(force=True)
    for k, v in ok:
        print(f"{'✔' if v else '✘'} {k}")
    print(f"{sum(v for _, v in ok)}/{len(ok)} 통과")
finally:
    wb.Close(SaveChanges=False)
    xl.Quit()
