"""실행 중인 Excel(여러 개여도)에서 열린 통합 문서·시트와 데이터 범위를 읽는다. FR-01."""
import glob, os, shutil
import pythoncom
import win32com
import win32com.client as w
import win32gui
import win32process

# pywin32의 Excel 형식 캐시(%TEMP%\gen_py)가 반쯤 지워지면(__init__.py 없음) Excel 연결이 전부
# "CLSIDToClassMap" AttributeError로 실패해 '파일을 찾지 못함'으로 보였음 → 깨진 폴더는 지우면 다시 만들어진다
for _d in glob.glob(os.path.join(win32com.__gen_path__, "*", "")):
    if os.path.basename(_d[:-1]) != "__pycache__" and not os.path.exists(os.path.join(_d, "__init__.py")):
        shutil.rmtree(_d, ignore_errors=True)


def _typed(obj):
    """형식 정보로 감싼다. 캐시 없이(새 PC·캐시 삭제 후) 감싸면 Copy(After=…) 같은 이름 붙인 인자가 무시돼
    시트가 새 통합 문서로 복사되는 등 동작이 달라짐 → 캐시를 만들어서라도 늘 같은 방식으로 연결한다."""
    try:
        return w.gencache.EnsureDispatch(obj)
    except Exception:
        return w.Dispatch(obj)


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
                    wb = _typed(rot.GetObject(mon).QueryInterface(pythoncom.IID_IDispatch))
                    found[(wb.Application.Hwnd, wb.Name)] = wb
                except Exception:
                    pass  # 다른 프로그램이 연 파일이면 건너뜀
    except Exception:
        pass
    try:
        app = _typed(w.GetActiveObject("Excel.Application"))
        for wb in app.Workbooks:
            found.setdefault((app.Hwnd, wb.Name), wb)
    except Exception:
        pass
    return [wb for wb in found.values() if _shown(wb)]


def _shown(wb):
    """숨은 Excel(다른 프로그램·시험이 자동화로 띄운 것)의 문서는 사용자가 연 파일이 아니라 뺀다."""
    try:
        return bool(wb.Application.Visible)
    except Exception:
        return True  # 셀 편집 중 등으로 바빠 못 물으면 보이는 것으로


def list_files():
    """열린 파일 목록 [{'wb': 통합 문서, 'name': 파일명}]. Excel에서 지금 보고 있는 파일이 맨 앞에 온다."""
    wbs = _open_workbooks()
    if not wbs:
        raise ExcelError("열려 있는 Excel 파일을 찾지 못했습니다. Excel에서 파일을 연 뒤 새로고침을 눌러주세요.")
    current = foreground_workbook(wbs)
    if current is None:  # 셀리 창이 앞에 있을 때(시작·새로고침): Excel이 보고 있는 파일을 먼저
        try:
            name = wbs[0].Application.ActiveWorkbook.Name
            current = next(wb for wb in wbs if wb.Name == name)
        except Exception:
            pass
    wbs.sort(key=lambda wb: current is None or
             (wb.Application.Hwnd, wb.Name) != (current.Application.Hwnd, current.Name))
    return [{"wb": wb, "name": wb.Name} for wb in wbs]


def foreground_workbook(wbs):
    """현재 앞에 있는 Excel 프로세스의 활성 통합 문서만 돌려준다."""
    hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        return None
    pid = win32process.GetWindowThreadProcessId(hwnd)[1]
    for wb in wbs:
        try:
            xl = wb.Application
            if win32process.GetWindowThreadProcessId(xl.Hwnd)[1] == pid:
                active = xl.ActiveWorkbook
                if active is not None:
                    return next((item for item in wbs if item.Application.Hwnd == xl.Hwnd
                                 and item.Name == active.Name), None)
        except Exception:
            continue
    return None


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
