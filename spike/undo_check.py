"""되돌리기 검증: 작업 전 시트를 숨김 사본으로 떠 두고, 작업 후 원래 상태로 되돌려지는지 확인.

10,000행 파일의 임시 복사본으로만 시험한다. 원본 파일은 열지도 저장하지도 않는다.
"""
import os, shutil, sys, tempfile, time, hashlib
import win32com.client as w

BACKUP = "_ai_undo"


class Undo:
    """직전 AI 작업 1단계 되돌리기."""

    def __init__(self, ws):
        self.ws, self.wb = ws, ws.Parent

    def fingerprint(self):
        r = self.ws.UsedRange
        raw = repr((r.Address, r.Formula2, r.Value2, filters(self.ws),
                    self.ws.PivotTables().Count, [s.Name for s in self.wb.Worksheets]))
        return hashlib.sha256(raw.encode()).hexdigest()

    def before(self):
        xl = self.wb.Application
        self.drop()
        self.sheets = [s.Name for s in self.wb.Worksheets]
        self.tables = [lo.Name for lo in self.ws.ListObjects]
        # shortcut: 숨김 사본이 사용자 문서 안에 있어 저장하면 같이 저장됨, 실제 앱에선 종료 시 삭제 안내 필요
        self.ws.Copy(After=self.wb.Worksheets(self.wb.Worksheets.Count))
        b = self.wb.Worksheets(self.wb.Worksheets.Count)
        b.Name, b.Visible = BACKUP, 2  # 2 = xlSheetVeryHidden
        self.ws.Activate()
        xl.ScreenUpdating = True

    def after(self):
        self.done = self.fingerprint()

    def undo(self):
        if self.fingerprint() != self.done:
            raise RuntimeError("AI 작업 뒤에 직접 수정한 내용이 있어 되돌리지 않음")
        xl = self.wb.Application
        xl.DisplayAlerts = False
        for s in list(self.wb.Worksheets):
            if s.Name not in self.sheets and s.Name != BACKUP:
                s.Delete()
        xl.DisplayAlerts = True
        for pt in list(self.ws.PivotTables()):
            pt.TableRange2.Clear()
        # 시트 전체를 비우고 사본을 붙이면 값·수식·서식·순서·필터 해제·표가 한 번에 돌아옴
        # shortcut: 작업 전 필터가 걸려 있었다면 그 조건은 복원 안 됨, 필터 위 작업 지원 시 보강
        for lo in list(self.ws.ListObjects):
            if lo.AutoFilter and lo.AutoFilter.FilterMode:
                lo.AutoFilter.ShowAllData()
            lo.Unlist()  # 표를 남긴 채 붙이면 충돌하므로 일반 범위로 바꾼 뒤 비움
        self.ws.Cells.Clear()  # Delete는 다른 시트 수식의 참조를 #REF!로 깨뜨려서 Clear를 씀
        self.wb.Worksheets(BACKUP).Cells.Copy(self.ws.Cells)
        for lo, name in zip(self.ws.ListObjects, self.tables):
            lo.Name = name  # 붙여넣으면 표 이름이 바뀌어 다른 수식의 표 참조가 깨지므로 원래 이름으로
        self.drop()

    def drop(self):
        xl = self.wb.Application
        xl.DisplayAlerts = False
        for s in list(self.wb.Worksheets):
            if s.Name == BACKUP:
                s.Visible = -1
                s.Delete()
        xl.DisplayAlerts = True


def filters(ws):
    return ws.AutoFilterMode, [lo.AutoFilter.FilterMode if lo.AutoFilter else None for lo in ws.ListObjects]


def state(ws):
    r = ws.UsedRange
    return (r.Address, r.Formula2, r.Value2, r.Interior.ColorIndex, filters(ws),
            [lo.Name for lo in ws.ListObjects],
            ws.PivotTables().Count, [s.Name for s in ws.Parent.Worksheets if s.Visible == -1],
            [ws.Columns(i).ColumnWidth for i in range(1, 12)])


# --- AI 작업 흉내 (실제 앱에서는 ChatGPT가 정한 계획을 이 함수들이 실행) ---
def op_sort(ws):
    lo = ws.ListObjects(1)
    lo.Sort.SortFields.Clear()
    lo.Sort.SortFields.Add(Key=lo.ListColumns("INVOICE_AMOUNT").Range, Order=2)
    lo.Sort.Header = 1
    lo.Sort.Apply()

def op_filter(ws):
    ws.ListObjects(1).Range.AutoFilter(Field=4, Criteria1=">100000000")

def op_classify(ws):
    n = ws.UsedRange.Rows.Count
    ws.Range("H1").Value = "AI 분류"
    ws.Range(f"H2:H{n}").Value = "확인 필요"
    ws.Range("H2:H50").Interior.Color = 0x00FFFF

def op_formula(ws):
    n = ws.UsedRange.Rows.Count
    ws.Range("I1").Value = "연도"
    ws.Range(f"I2:I{n}").Formula2 = "=IFERROR(YEAR(B2),\"\")"
    ws.Range("K1").Formula2 = f"=UNIQUE(F2:F{n})"

def op_pivot(ws):
    wb = ws.Parent
    n = ws.UsedRange.Rows.Count
    new = wb.Worksheets.Add(After=ws)
    new.Name = "AI 피벗"
    pt = wb.PivotCaches().Create(1, ws.Range(f"A1:G{n}")).CreatePivotTable(new.Range("A3"), "AI피벗")
    pt.PivotFields("VENDOR_NAME").Orientation = 1
    pt.AddDataField(pt.PivotFields("INVOICE_AMOUNT"), "금액 합계", -4157)
    ws.Activate()


def main():
    src = os.path.join(os.path.dirname(__file__), "..", "거래내역_대량검증용_10000행.xlsx")
    tmp = os.path.join(tempfile.mkdtemp(), "undo_test.xlsx")
    shutil.copy(src, tmp)
    xl = w.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = True
    wb = xl.Workbooks.Open(tmp)
    ws = wb.Worksheets("더미 데이터")
    ref = wb.Worksheets.Add(After=ws)
    ref.Name = "사용자 요약"
    ref.Range("A1").Formula2 = "=SUM('더미 데이터'!D2:D10001)"
    ws.Activate()
    ref_before = ref.Range("A1").Formula2, ref.Range("A1").Value2
    u = Undo(ws)
    ok = True
    try:
        for op in (op_sort, op_filter, op_classify, op_formula, op_pivot):
            orig = state(ws)
            t = time.perf_counter(); u.before(); t_snap = time.perf_counter() - t
            op(ws)
            u.after()
            assert state(ws) != orig, f"{op.__name__}: 작업이 아무것도 안 바꿈"
            t = time.perf_counter(); u.undo(); t_undo = time.perf_counter() - t
            same = state(ws) == orig
            ok &= same
            print(f"{op.__name__:12} 되돌리기 {'OK' if same else '실패'}  (사본 {t_snap:.2f}초 / 복원 {t_undo:.2f}초)")

        ref_ok = (ref.Range("A1").Formula2, ref.Range("A1").Value2) == ref_before
        ok &= ref_ok
        print("다른 시트 참조    ", "OK" if ref_ok else f"실패: {ref.Range('A1').Formula2}")

        orig = state(ws)
        u.before(); op_classify(ws); u.after()
        ws.Range("A2").Value = "사용자가 직접 고침"
        try:
            u.undo()
            print("직접 수정 충돌    실패: 사용자 수정을 덮어씀"); ok = False
        except RuntimeError as e:
            print("직접 수정 충돌    OK:", e)
    finally:
        wb.Close(SaveChanges=False)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
