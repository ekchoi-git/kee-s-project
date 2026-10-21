"""실행 중인 Excel(여러 개여도)에서 열린 통합 문서·시트와 데이터 범위를 읽는다. FR-01."""
import pythoncom
import win32com.client as w


class ExcelError(Exception):
    """사용자에게 그대로 보여줄 수 있는 문구를 담는다."""


def _open_workbooks():
    """Excel 프로세스가 여러 개여도 모두 찾는다 (파일로 저장된 문서는 실행 개체 표에서, 저장 전 문서는 기본 Excel에서)."""
    found = {}
    try:
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
        for mon in rot.EnumRunning():
            name = mon.GetDisplayName(ctx, None)
            if name.lower().endswith((".xlsx", ".xlsm", ".xls")):
                try:
                    wb = w.Dispatch(rot.GetObject(mon).QueryInterface(pythoncom.IID_IDispatch))
                    found[(wb.Application.Hwnd, wb.Name)] = wb
                except Exception:
                    pass  # 다른 프로그램이 연 파일이면 건너뜀
    except Exception:
        pass
    try:
        app = w.GetActiveObject("Excel.Application")
        for wb in app.Workbooks:
            found.setdefault((app.Hwnd, wb.Name), wb)
    except Exception:
        pass
    return list(found.values())


def list_files():
    """열린 파일 목록 [{'wb': 통합 문서, 'name': 파일명}]. Excel에서 지금 보고 있는 파일이 맨 앞에 온다."""
    wbs = _open_workbooks()
    if not wbs:
        raise ExcelError("열려 있는 Excel 파일을 찾지 못했습니다. Excel에서 파일을 연 뒤 새로고침을 눌러주세요.")
    try:
        current = wbs[0].Application.ActiveWorkbook.Name
    except Exception:
        current = None
    wbs.sort(key=lambda wb: wb.Name != current)
    return [{"wb": wb, "name": wb.Name} for wb in wbs]


def active_sheet(wb):
    """그 파일에서 지금 열려 있는(보이는) 시트. 일반 시트가 아니면(차트 시트 등) ExcelError."""
    ws = wb.ActiveSheet
    try:
        is_data = ws.Type == -4167  # xlWorksheet
    except Exception:  # 차트 시트에는 이 속성이 없음
        is_data = False
    if not is_data:
        raise ExcelError(f"'{ws.Name}'은 데이터 시트가 아닙니다. Excel에서 데이터가 있는 시트를 열어주세요.")
    return ws


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def sheet_info(ws, sample=5):
    """AI에 보낼 시트 설명: 열 제목·위치, 행 수, 앞쪽 몇 행, 결과를 둘 빈 열, 동적 배열 지원 여부.
    전체 데이터는 보내지 않는다 (PRD 8장)."""
    s = read_summary(ws)
    r0, c0, cols = s["first_row"], s["first_col"], s["cols"]
    last = r0 + s["rows"]
    cols_desc = [f"{col_letter(c0 + i)}열 '{h}' ({col_letter(c0 + i)}{r0 + 1}:{col_letter(c0 + i)}{last})"
                 for i, h in enumerate(s["headers"])]
    rows = []
    if s["rows"]:
        v = ws.Range(ws.Cells(r0 + 1, c0), ws.Cells(min(r0 + sample, last), c0 + cols - 1)).Value
        rows = [" | ".join("" if x is None else str(x) for x in (r if isinstance(r, tuple) else (r,)))
                for r in (v if isinstance(v, tuple) else ((v,),))]
    try:
        ws.Range("A1").Formula2
        dyn = "지원함"
    except Exception:
        dyn = "지원 안 함"
    return "\n".join([
        f"시트: {ws.Name} (파일 {ws.Parent.Name})",
        f"데이터 범위: {s['address']}, 머리글 {r0}행, 데이터 {s['rows']}행",
        "열: " + "; ".join(cols_desc),
        f"결과 시작 열: {col_letter(c0 + cols + 1)} (데이터 오른쪽 한 칸 띄움)",
        f"Excel 버전 {s['app_version']}, 동적 배열 함수 {dyn}",
        f"앞쪽 {len(rows)}행 예시:",
        *rows,
    ])


def read_summary(ws):
    """시트의 사용 범위, 행 수, 열 제목. 읽기 전용·보호 시트 여부도 함께 알려준다."""
    used = ws.UsedRange
    rows, cols = used.Rows.Count, used.Columns.Count
    first = used.Row
    v = ws.Range(ws.Cells(first, used.Column), ws.Cells(first, used.Column + cols - 1)).Value
    row = v[0] if isinstance(v, tuple) else (v,)  # 셀이 하나면 값이 튜플이 아님
    headers = [str(c if c is not None else "") for c in row]
    return {
        "first_row": first,
        "first_col": used.Column,
        "cols": cols,
        "address": used.Address.replace("$", ""),
        "rows": rows - 1,
        "headers": headers,
        "protected": bool(ws.ProtectContents),
        "readonly": bool(ws.Parent.ReadOnly),
        "app_version": ws.Application.Version,
    }
