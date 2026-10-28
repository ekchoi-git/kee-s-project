"""실제 Excel로 확인: 필터·Excel 표가 있는 시트에서 되돌리기가 모든 값·표 이름·보이는 행을 작업 전으로 돌리는지.
별도 Excel 인스턴스에서 새 통합 문서를 만들어 쓰고 닫는다(AI 호출 없음). 사용: python tests/e2e_undo_filters.py"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import win32com.client as w
from app.excel_ops import BACKUP, Target, Undo

xl = w.DispatchEx("Excel.Application")
xl.DisplayAlerts = False


def run(name, table, apply_filter, want_skipped=()):
    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)
    try:
        ws.Range("A1:C1").Value = ("이름", "부서", "금액")
        ws.Range("A2:C9").Value = [(f"사람{i}", ("영업", "구매", "회계")[i % 3], i * 100) for i in range(1, 9)]
        ws.Range("B3").Interior.Color = 255
        lo = ws.ListObjects.Add(1, ws.Range("A1:C9"), None, 1) if table else None
        if lo:
            lo.Name = "판매표"
        apply_filter(lo.Range if lo else ws.Range("A1:C9"))
        values = ws.Range("A1:C9").Value
        hidden = [r for r in range(1, 10) if ws.Rows(r).Hidden]
        tables = [x.Name for x in ws.ListObjects]

        u = Undo(Target(ws))
        u.before()
        ws.Range("C2").Value = 99999  # AI 작업 흉내
        u.after()
        skipped = u.undo()

        ok = {"값": ws.Range("A1:C9").Value == values,
              # 다시 못 건 필터는 모든 행이 보이는 게 맞음
              "숨은 행": [r for r in range(1, 10) if ws.Rows(r).Hidden] == ([] if want_skipped else hidden),
              "표 이름": [x.Name for x in ws.ListObjects] == tables,
              "사본 정리": BACKUP not in [s.Name for s in wb.Worksheets],
              "건너뛴 필터": skipped == list(want_skipped)}
        print(f"{'✔' if all(ok.values()) else '✘'} {name}: " + ", ".join(f"{k}={v}" for k, v in ok.items()))
        return all(ok.values())
    except Exception as e:
        print(f"✘ {name}: {type(e).__name__} {e}")
        return False
    finally:
        wb.Close(SaveChanges=False)


try:
    results = [
        run("필터 없음", False, lambda r: None),
        run("일반 범위 화살표만", False, lambda r: r.AutoFilter(Field=1)),
        run("일반 범위 단일 조건", False, lambda r: r.AutoFilter(Field=2, Criteria1="영업")),
        run("일반 범위 여러 값", False, lambda r: r.AutoFilter(Field=2, Criteria1=("영업", "회계"), Operator=7)),
        run("일반 범위 숫자 두 조건", False, lambda r: r.AutoFilter(Field=3, Criteria1=">=200", Operator=1, Criteria2="<=500")),
        run("일반 범위 색 필터(다시 못 걺)", False, lambda r: r.AutoFilter(Field=2, Criteria1=255, Operator=8), want_skipped=[2]),
        run("Excel 표 필터 없음", True, lambda r: None),
        run("Excel 표 단일 조건", True, lambda r: r.AutoFilter(Field=2, Criteria1="영업")),
        run("Excel 표 여러 값", True, lambda r: r.AutoFilter(Field=2, Criteria1=("영업", "회계"), Operator=7)),
    ]
    print(f"{sum(results)}/{len(results)} 통과")
finally:
    xl.Quit()
