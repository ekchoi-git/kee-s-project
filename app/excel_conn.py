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
    """열린 파일 목록. 각 항목: {'wb': 통합 문서, 'name': 파일명, 'sheets': 보이는 시트들, 'active': 지금 보고 있는 시트 이름 또는 None}.
    Excel에서 지금 보고 있는 파일이 맨 앞에 온다."""
    files = []
    for wb in _open_workbooks():
        sheets = [ws for ws in wb.Worksheets if ws.Visible == -1]
        try:
            active_wb = wb.Application.ActiveWorkbook
            active = wb.Application.ActiveSheet.Name if active_wb and active_wb.Name == wb.Name else None
        except Exception:
            active = None
        if sheets:
            files.append({"wb": wb, "name": wb.Name, "sheets": sheets, "active": active})
    if not files:
        raise ExcelError("열려 있는 Excel 파일을 찾지 못했습니다. Excel에서 파일을 연 뒤 새로고침을 눌러주세요.")
    files.sort(key=lambda f: f["active"] is None)
    return files


def read_summary(ws):
    """시트의 사용 범위, 행 수, 열 제목. 읽기 전용·보호 시트 여부도 함께 알려준다."""
    used = ws.UsedRange
    rows, cols = used.Rows.Count, used.Columns.Count
    first = used.Row
    v = ws.Range(ws.Cells(first, used.Column), ws.Cells(first, used.Column + cols - 1)).Value
    row = v[0] if isinstance(v, tuple) else (v,)  # 셀이 하나면 값이 튜플이 아님
    headers = [str(c if c is not None else "") for c in row]
    return {
        "address": used.Address.replace("$", ""),
        "rows": rows - 1,
        "headers": headers,
        "protected": bool(ws.ProtectContents),
        "readonly": bool(ws.Parent.ReadOnly),
        "app_version": ws.Application.Version,
    }
