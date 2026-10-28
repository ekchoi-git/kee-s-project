"""Excel 쪽: 대상 표 읽기, AI 계획 검사, 실행, 직전 작업 되돌리기."""
import hashlib, re, time

import pywintypes

BACKUP = "_ai_undo"


class PlanError(Exception):
    """사용자에게 그대로 보여줄 실패 이유."""


class FormulaError(PlanError):
    """AI가 만든 수식 자체가 잘못된 경우. 되돌린 뒤 AI에게 고치게 할 수 있다."""


# --- 계획 검사 (PRD 8장: 허용된 작업만, 외부 통신·다른 파일·매크로 차단) ---
ADDRESS = re.compile(r"^\$?[A-Z]{1,3}\$?[1-9]\d*(:\$?[A-Z]{1,3}\$?[1-9]\d*)?$")
BLOCKED_FN = re.compile(r"\b(WEBSERVICE|RTD|CALL|REGISTER(\.ID)?|HYPERLINK|IMAGE|STOCKHISTORY|CUBE[A-Z]*|EXEC)\s*\(", re.I)
EXTERNAL = re.compile(r"https?://|[A-Za-z]:\\|\\\\|\[[^\]]*\.(xl\w*|csv)\]|\w\|[^,)]*!", re.I)

CRITERIA = {
    "=": "={}", "<>": "<>{}", ">": ">{}", ">=": ">={}", "<": "<{}", "<=": "<={}",
    "contains": "=*{}*", "begins": "={}*", "blank": "=", "notblank": "<>",
}

PIVOT_FUNC = {"sum": -4157, "count": -4112, "average": -4106, "max": -4136, "min": -4139}
XL_TOP10_ITEMS = 3


def top_items_count(value):
    """Excel 상위 항목 필터의 개수. Excel 필터 대화상자와 같은 1~500 범위만 허용한다."""
    value = str(value).strip()
    if not re.fullmatch(r"[1-9]\d*", value):
        return None
    count = int(value)
    return count if count <= 500 else None


def apply_filter(region, field, op, value):
    """필터 한 개를 적용하고 사용자에게 보여줄 조건 문구를 돌려준다."""
    if op == "top_items":
        count = top_items_count(value)
        if count is None:
            raise PlanError("상위 항목 필터의 개수는 1~500 사이 정수여야 합니다.")
        region.AutoFilter(Field=field, Criteria1=str(count), Operator=XL_TOP10_ITEMS)
        return f"상위 {count}개"
    region.AutoFilter(Field=field, Criteria1=CRITERIA[op].format(value))
    return f"{op} {value}".rstrip()


def validate(plan, headers):
    """실행 전에 계획의 열 이름·주소·수식을 검사. 문제 목록을 돌려준다(비어 있으면 통과)."""
    bad = []
    cols = [f["column"] for f in plan["filters"]] + [s["column"] for s in plan["sorts"]]
    if plan["pivot"]:
        p = plan["pivot"]
        cols += p["rows"] + p["columns"] + [v["field"] for v in p["values"]]
        if not p["values"]:
            bad.append("피벗에 집계할 값이 없습니다.")
    if plan["classify"]:
        cols.append(plan["classify"]["column"])
        if not re.fullmatch(r"[A-Za-z]{0,3}", plan["classify"]["output"]):
            bad.append(f"'{plan['classify']['output']}'는 결과를 넣을 열 글자가 아닙니다.")
    bad += [f"'{c}' 열이 표에 없습니다." for c in dict.fromkeys(cols) if c not in headers]
    for f in plan["filters"]:
        if f["op"] == "top_items" and top_items_count(f["value"]) is None:
            bad.append(f"'{f['value']}'는 상위 항목 필터에 쓸 수 있는 개수가 아닙니다. 1~500 사이 정수를 사용해주세요.")
    for w in plan["writes"]:
        if not ADDRESS.match(w["address"].upper()):
            bad.append(f"'{w['address']}'는 이 시트의 셀 주소가 아닙니다.")
        if BLOCKED_FN.search(w["content"]) or EXTERNAL.search(w["content"]):
            bad.append(f"'{w['content'][:60]}'는 외부 연결·다른 파일·추가 기능이 필요한 수식이라 실행하지 않습니다.")
    bad += [f"'{c}' 열이 표에 없습니다." for c in plan["to_number"] if c not in headers]
    if plan["to_number"] and not plan["edit_source"]:
        bad.append("원본 수정 지시 없이 원본 값을 바꾸려 했습니다.")
    for f in plan["formats"]:
        if not ADDRESS.match(f["address"].upper()):
            bad.append(f"'{f['address']}'는 이 시트의 셀 주소가 아닙니다.")
        if not f["format"] or len(f["format"]) > 60:
            bad.append(f"'{f['format']}'는 표시 형식이 아닙니다.")
    for h in plan["highlights"]:
        if not ADDRESS.match(h["address"].upper()):
            bad.append(f"'{h['address']}'는 이 시트의 셀 주소가 아닙니다.")
        if not h["formula"].startswith("=") or len(h["formula"]) > 255:  # 조건부 서식 수식 한도 255자
            bad.append(f"'{h['formula'][:60]}'는 조건부 서식 수식이 아닙니다.")
        if BLOCKED_FN.search(h["formula"]) or EXTERNAL.search(h["formula"]):
            bad.append(f"'{h['formula'][:60]}'는 외부 연결·다른 파일이 필요한 수식이라 실행하지 않습니다.")
    if plan["action"] == "change" and not (plan["filters"] or plan["sorts"] or plan["writes"] or plan["pivot"] or plan["classify"]
                                           or plan["to_number"] or plan["formats"] or plan["highlights"]):
        bad.append("실행할 작업이 없습니다.")
    return bad


# --- 오류 값: COM은 #N/A 같은 오류를 음수 정수로 돌려줌 ---
ERR_BASE = -2146828288
ERR_NAMES = {2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!", 2029: "#NAME?", 2036: "#NUM!",
             2042: "#N/A", 2045: "#SPILL!", 2046: "#CONNECT!", 2047: "#BLOCKED!", 2048: "#UNKNOWN!",
             2049: "#FIELD!", 2050: "#CALC!"}
# 데이터 때문에 일부 행에서 생길 수 있는 오류는 경고만, 나머지는 수식 자체가 잘못된 것으로 보고 취소
DATA_ERRORS = {"#DIV/0!", "#VALUE!", "#N/A", "#NUM!"}


def errors_in(values):
    if not isinstance(values, tuple):
        values = ((values,),)
    found = {}
    for row in values:
        for v in row:
            if type(v) is int and ERR_BASE < v < ERR_BASE + 3000:
                name = ERR_NAMES.get(v - ERR_BASE, "#오류")
                found[name] = found.get(name, 0) + 1
    return found


def _num(v):
    """글자로 저장된 숫자와 (△151) 같은 음수 표기면 그 값, 아니면 None."""
    if not isinstance(v, str):
        return None
    cleaned = v.strip().replace(",", "")
    negative = re.fullmatch(r"(?:\(\s*△\s*(\d+(?:\.\d+)?)\s*\)|△\s*(\d+(?:\.\d+)?))", cleaned)
    if negative:
        return -float(negative.group(1) or negative.group(2))
    try:
        return float(cleaned)
    except ValueError:
        return None


def rows_where(t, j, pred):
    """표의 j번째(0부터) 열에서 조건에 맞는 칸의 (행 번호, 값) 목록."""
    v = t.region.Columns(j + 1).Value2
    v = v if isinstance(v, tuple) else ((v,),)
    return [(t.region.Row + i, x) for i, (x,) in enumerate(v) if i and pred(x)]


def text_numbers(t, j):
    """진짜 숫자 사이에 섞인 '글자로 저장된 숫자' 칸. 열 전체가 글자(공급사번호 같은 ID)면 실수가 아니므로 빈 목록."""
    v = t.region.Columns(j + 1).Value2
    v = v if isinstance(v, tuple) else ((v,),)
    if not any(isinstance(x, (int, float)) and not isinstance(x, bool) for (x,) in v[1:]):
        return []
    return [(t.region.Row + i, x) for i, (x,) in enumerate(v) if i and _num(x) is not None]


def bad_text(t, j):
    """대부분 날짜·숫자인 열에 섞인, 숫자로도 읽을 수 없는 글자 칸(예: '날짜미상', '2023/02/30'). YEAR·SUM에서 오류나 누락이 된다."""
    v = t.region.Columns(j + 1).Value2
    v = v if isinstance(v, tuple) else ((v,),)
    vals = [x for (x,) in v[1:] if not blank(x)]
    nums = sum(1 for x in vals if isinstance(x, (int, float)) and not isinstance(x, bool))
    if nums * 2 <= len(vals):
        return []
    return [(t.region.Row + i, x) for i, (x,) in enumerate(v) if i and isinstance(x, str) and not blank(x) and _num(x) is None]


def spots(letter, rows, limit=5):
    """[452..461, 480] → 'D452~D461, D480'. 묶음이 많으면 앞 몇 개만."""
    runs = []
    for r in rows:
        if runs and r == runs[-1][1] + 1:
            runs[-1][1] = r
        else:
            runs.append([r, r])
    text = ", ".join(f"{letter}{a}" if a == b else f"{letter}{a}~{letter}{b}" for a, b in runs[:limit])
    return text + (f" 외 {sum(b - a + 1 for a, b in runs[limit:])}칸" if len(runs) > limit else "")


def blank(x):
    return x is None or (isinstance(x, str) and not x.strip())


def source_changes(t, plan):
    """원본 표 안을 바꾸는 내용(확인 창에 보여줄 줄). 원본 수정 지시가 없으면 빈 목록(실행 단계에서 막힘)."""
    if not plan["edit_source"]:
        return []
    lines = []
    for c in plan["to_number"]:
        j = t.headers.index(c)
        hits = rows_where(t, j, lambda x: _num(x) is not None)
        if hits:
            lines.append(f"{c} 열 글자로 저장된 숫자 {len(hits)}칸을 숫자로 바꿈 ({spots(col_letter(t.region.Column + j), [r for r, _ in hits])})")
    if not plan["new_sheet"]:
        if rebuilds(t, plan):
            lines.append(f"기존 표 {t.region.Address.replace('$', '')}를 비우고 새로 씀(수식 결과 표라 남은 칸이 새 결과를 막음)")
        for w in plan["writes"]:
            if ADDRESS.match(w["address"].upper()) and t.xl.Intersect(t.ws.Range(w["address"]), t.region) is not None:
                lines.append(f"{w['address']} ← {w['content'][:60]}")
    c = plan["classify"]
    if c and c["output"] and _in_table(t, _col_index(c["output"])) \
            and t.ws.Cells(t.region.Row, _col_index(c["output"])).Value2 != f"{c['column']} 분류":
        lines.append(f"{c['output'].upper()}열을 분류 결과로 덮어씀")
    return lines


def rebuilds(t, plan):
    """원본 수정으로 수식 결과(펼쳐진 배열) 표 안에 새로 쓰면 표를 다시 짜는 것: 남은 옛 칸이 새 결과를 막아 #SPILL이 나므로 먼저 비운다."""
    return bool(plan["edit_source"] and not plan["new_sheet"] and t.spills() and any(
        ADDRESS.match(w["address"].upper()) and t.xl.Intersect(t.ws.Range(w["address"]), t.region) is not None
        for w in plan["writes"]))


SAME_SHEET_REF = re.compile(r"(?<![!\w$'.:])\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?(?![\w(])")


def stale_refs(t, writes, ranges):
    """표를 비우고 다시 쓸 때, 비운 뒤 다시 채워지지 않는 옛 표 칸을 읽는 수식 참조 목록.
    새로 쓰는 칸(및 그 펼침 B1# 등)을 가리키는 참조는 괜찮다."""
    xl, bad = t.xl, []
    for w in writes:
        if not w["content"].startswith("="):
            continue
        body = re.sub(r'"[^"]*"', '""', w["content"])  # 따옴표 안 글자는 참조가 아님
        for ref in SAME_SHEET_REF.findall(body):
            rng = t.ws.Range(ref)
            if xl.Intersect(rng, t.region) is not None and not any(xl.Intersect(rng, r) is not None for r in ranges):
                bad.append(f"{w['address']}의 {ref}")
    return bad


def _col_index(letters):
    return sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(letters.upper())))


def _in_table(t, c):
    return t.region.Column <= c < t.region.Column + t.region.Columns.Count


def _ref(letter, same_sheet):
    """수식이 원본의 그 열을 참조하는지 찾는 패턴. 새 시트 수식은 '시트명'! 뒤의 참조만 원본이다."""
    return rf"(?<![A-Za-z_])\$?{letter}\$?\d" if same_sheet else rf"!\$?{letter}\$?\d"


def suggest_to_number(t, plan):
    """1차 결과가 쓴 원본 열 중 글자로 저장된 숫자가 있는 열: [(머리글, 칸 위치, 칸 수)]. 제안 버튼에 쓴다."""
    if plan["edit_source"]:
        return []
    text = " ".join(w["content"] for w in plan["writes"] if w["content"].startswith("="))
    js = {j for j in range(len(t.headers)) if re.search(_ref(col_letter(t.region.Column + j), not plan["new_sheet"]), text)}
    if plan["pivot"]:
        js |= {t.headers.index(v["field"]) for v in plan["pivot"]["values"] if v["func"] != "count"}
    js |= {t.headers.index(s["column"]) for s in plan["sorts"]}
    js |= {t.headers.index(f["column"]) for f in plan["filters"] if f["op"] in "<>=<=>="}
    out = []
    for j in sorted(js):
        tn = text_numbers(t, j)
        if tn:
            out.append((t.headers[j], spots(col_letter(t.region.Column + j), [r for r, _ in tn]), len(tn)))
    return out


def assumption_notes(t, formulas, same_sheet):
    """수식이 참조한 원본 열에 글자 숫자·빈 칸이 있으면 셀리가 어떻게 다뤘는지 건수·위치·금액으로 적는다."""
    # shortcut: 열 참조와 '--'/VALUE 포함 여부를 글자로 판단, 수식이 복잡해져 오판이 보이면 Excel로 직접 검산
    text = " ".join(formulas)
    converts = bool(re.search(r"--|\bN?VALUE\(|NUMBERVALUE\(", text, re.I))
    notes = []
    for j, h in enumerate(t.headers):
        letter = col_letter(t.region.Column + j)
        if not re.search(_ref(letter, same_sheet), text):
            continue
        tn = text_numbers(t, j)
        if tn:
            amt = f"{sum(_num(x) for _, x in tn):,.0f}"
            where = spots(letter, [r for r, _ in tn])
            notes.append(f"⚠ 임의 판단: {letter}열 '{h}'의 글자로 저장된 숫자 {len(tn)}칸({where})을 "
                         + (f"숫자로 바꿔 포함 → 원본 {letter}열 SUM보다 {amt} 많음" if converts else f"제외 → 합계에서 {amt} 빠짐"))
        junk = bad_text(t, j)
        if junk:
            where = spots(letter, [r for r, _ in junk])
            notes.append(f"⚠ 임의 판단: {letter}열 '{h}'의 날짜·숫자가 아닌 글자 {len(junk)}칸({where})"
                         + ("을 '(날짜 오류)' 묶음으로 따로 집계" if "(날짜 오류)" in text else "은 계산할 수 없어 오류나 누락이 생길 수 있음"))
        bl = rows_where(t, j, blank)
        grouped = "(빈 값)" in text and re.search(rf"\$?{letter}\$?\d+(:\$?{letter}\$?\d+)?\s*=\s*\"\"", text)
        if bl and grouped:
            notes.append(f"⚠ 임의 판단: {letter}열 '{h}' 빈 칸 {len(bl)}칸({spots(letter, [r for r, _ in bl])})을 '(빈 값)' 묶음으로 따로 집계")
        elif bl:
            notes.append(f"참고: {letter}열 '{h}' 빈 칸 {len(bl)}칸({spots(letter, [r for r, _ in bl])})")
    return notes


def pivot_notes(t, p):
    """피벗은 글자 숫자를 합계에 넣지 못하고 빈 칸은 '(비어 있음)'으로 묶는다. 그 영향을 건수·위치·금액으로 적는다."""
    notes = []
    for v in p["values"]:
        if v["func"] == "count":
            continue
        j = t.headers.index(v["field"])
        letter = col_letter(t.region.Column + j)
        tn = text_numbers(t, j)
        if tn:
            notes.append(f"⚠ 임의 판단: {letter}열 '{v['field']}'의 글자로 저장된 숫자 {len(tn)}칸({spots(letter, [r for r, _ in tn])})은 "
                         f"피벗 합계에서 빠짐 → 원본 값 기준 {sum(_num(x) for _, x in tn):,.0f} 적음")
    for c in p["rows"] + p["columns"]:
        j = t.headers.index(c)
        letter = col_letter(t.region.Column + j)
        bl = rows_where(t, j, blank)
        if bl:
            notes.append(f"⚠ 임의 판단: {letter}열 '{c}' 빈 칸 {len(bl)}칸({spots(letter, [r for r, _ in bl])})은 피벗에서 '(비어 있음)' 묶음으로 표시")
    return notes


def numeric_text(values):
    """숫자처럼 보이지만 글자로 저장된 칸 수. 정렬·필터에서 숫자로 취급되지 않는다."""
    n = 0
    for (v,) in values:
        if isinstance(v, str):
            try:
                float(v.replace(",", ""))
                n += 1
            except ValueError:
                pass
    return n


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


TOTAL_LABELS = {"합계", "총합계"}


class Target:
    """요청 시점에 고정한 통합 문서·시트와 그 안의 데이터 표."""

    def __init__(self, ws):
        self.ws = ws
        retry_busy(self._load)

    def _load(self):
        ws = self.ws
        self.wb, self.xl = ws.Parent, ws.Application
        self.lo = self.ws.ListObjects(1) if self.ws.ListObjects.Count else None
        # shortcut: 시트의 첫 표(없으면 첫 피벗, 그것도 없으면 위에서부터 처음 만나는 2행 이상 범위)만 대상으로 봄, 여러 표 지원 시 선택 범위로 확장
        if self.lo:
            self.region = self.lo.Range
        elif self.ws.PivotTables().Count:
            self.region = self.ws.PivotTables(1).TableRange1
        else:
            self.region = first_block(self.ws)
        # 표 바로 아래 붙은 합계 줄(피벗 '총합계' 포함)은 데이터가 아님: 넣으면 다음 합계에서 두 번 더해지고 정렬에 섞인다
        r = self.region
        last = r.Cells(r.Rows.Count, 1).Value2
        self.total_row = None
        if r.Rows.Count > 2 and isinstance(last, str) and last.replace(" ", "") in TOTAL_LABELS:
            self.total_row = r.Row + r.Rows.Count - 1
            self.region = ws.Range(r.Cells(1, 1), r.Cells(r.Rows.Count - 1, r.Columns.Count))  # pywin32의 Resize(n)은 n번째 칸이 됨
        row1 = self.region.Rows(1).Value2
        self.headers = [str(h) if h is not None else "" for h in (row1[0] if isinstance(row1, tuple) else (row1,))]

    def check(self):
        if self.wb.ReadOnly:
            raise PlanError("읽기 전용 문서라 작업하지 않습니다. 편집 가능한 상태로 열어주세요.")
        if self.ws.ProtectContents:
            raise PlanError("보호된 시트라 작업하지 않습니다. 보호는 직접 해제해주세요.")
        if self.region.MergeCells is not False:
            raise PlanError("병합된 셀이 있는 표는 지원하지 않습니다.")
        if self.region.Rows.Count < 2 or not any(self.headers):
            raise PlanError("머리글 있는 표를 찾지 못했습니다. 머리글 한 줄과 데이터가 붙어 있는 표가 필요합니다.")

    def spills(self):
        """표의 첫 데이터 행이 펼쳐진 수식(동적 배열)에서 온 것이면 [(시작 칸, 수식, 펼쳐진 범위)]. 이런 표는 Excel 정렬·필터가 거부한다."""
        if not dynamic_arrays(self.ws):
            return []
        out, seen = [], set()
        for j in range(1, self.region.Columns.Count + 1):
            cell = self.region.Cells(2, j)
            if not cell.HasSpill:
                continue
            top = cell.SpillParent
            a = top.Address.replace("$", "")
            if a not in seen:
                seen.add(a)
                out.append((a, top.Formula2, top.SpillingToRange.Address.replace("$", "")))
        return out

    def info(self):
        r = self.region
        first, last = r.Row, r.Row + r.Rows.Count - 1
        cols = [f"{col_letter(r.Column + i)}={h}" for i, h in enumerate(self.headers)]
        n = min(5, r.Rows.Count - 1)
        sample = [" | ".join(str(r.Cells(2 + i, j + 1).Text) for j in range(len(self.headers))) for i in range(n)]
        used = self.ws.UsedRange
        free = col_letter(max(r.Column + r.Columns.Count, used.Column + used.Columns.Count) + 1)
        return "\n".join([
            f"통합 문서: {self.wb.Name} / 시트: {self.ws.Name}",
            f"데이터 범위: {r.Address} (머리글 {first}행, 데이터 {first + 1}~{last}행, {last - first}행)",
            *([f"합계 줄: {self.total_row}행 (표 바로 아래 붙은 합계라 데이터 범위에 넣지 않음)"] if self.total_row else []),
            *[f"이 표는 수식 결과(펼쳐진 배열): {a}: {f}  → {rng}에 펼쳐짐" for a, f, rng in self.spills()],
            *self.column_formulas(),
            "열: " + ", ".join(cols),
            f"표 바로 아래 행: {last + 1}행, 표 바로 오른쪽 열: {col_letter(r.Column + r.Columns.Count)}열 (합계 줄·합계 열은 여기에)",
            f"결과 시작 열: {free} (따로 떨어진 결과를 놓을 때만, 그 왼쪽은 사용 중)",
            f"Excel 버전 {self.xl.Version}, 동적 배열 함수 {'지원함' if dynamic_arrays(self.ws) else '지원 안 함'}",
            *self.quality(),
            "샘플 5행:", *sample,
            *self.outside(),
            *self.other_sheets(),
        ])

    def other_sheets(self, limit=30):
        """같은 파일의 다른 시트 표(이름·범위·머리글). 이게 없으면 AI는 여러 시트를 합치거나 참조할 때 시트 이름을 되묻는다."""
        lines = []
        for i in range(1, self.wb.Worksheets.Count + 1):
            ws = self.wb.Worksheets(i)
            if ws.Name == self.ws.Name:
                lines.append(f"'{ws.Name}': (현재 시트)")  # "2번째 시트" 같은 순서 지시를 풀 수 있게 자리만 표시
                continue
            if ws.Name == BACKUP or ws.Type != -4167 or ws.UsedRange.Count < 2:  # -4167 = xlWorksheet
                continue
            r = ws.ListObjects(1).Range if ws.ListObjects.Count else first_block(ws)
            row1 = r.Rows(1).Value2
            heads = [str(h) if h is not None else "" for h in (row1[0] if isinstance(row1, tuple) else (row1,))]
            f = str(r.Cells(1, 1).Formula)  # 이미 합친 결과 시트를 또 합치면 행이 두 번 들어감
            lines.append(f"'{ws.Name}': 데이터 범위 {r.Address} (머리글 {r.Row}행, 데이터 {r.Rows.Count - 1}행) 열: {', '.join(heads)}"
                         + (f"  ← 수식 결과 시트: {f[:80]}" if f.startswith("=") else ""))
            if len(lines) >= limit:
                break
        return ["같은 파일의 시트(파일 안 순서대로):", *lines] if len(lines) > 1 else []

    def column_formulas(self):
        """칸마다 채운 수식 열(펼쳐진 배열이 아닌 것)의 첫 데이터 행 수식. 이게 없으면 AI는 '구매비 합계'가
        어느 원본 열에서 계산되는지 몰라 표를 옮기거나 다시 만들지 못하고 원본 열을 되묻는다."""
        r, row = self.region, self.region.Row + 1
        if r.Rows.Count < 2:
            return []
        dyn = dynamic_arrays(self.ws)
        lines = []
        for j in range(r.Columns.Count):
            cell = self.ws.Cells(row, r.Column + j)
            f = str(cell.Formula2 if dyn else cell.Formula)
            if not f.startswith("=") or (dyn and cell.HasSpill):
                continue  # 펼쳐진 배열의 시작 칸은 spills()가 이미 알려줌
            below = r.Rows.Count > 2 and str(self.ws.Cells(row + 1, r.Column + j).Formula).startswith("=")
            lines.append(f"{col_letter(r.Column + j)}{row}: {f[:300]}" + ("  (아래 행에도 같은 식으로 채워져 있음)" if below else ""))
        return ["표 열의 수식(첫 데이터 행):", *lines] if lines else []

    def quality(self):
        """열별 빈 칸·글자로 저장된 숫자 수. AI가 이를 모르면 집계에서 그 행을 조용히 빠뜨린다."""
        v = self.region.Value2[1:]
        notes = []
        for j, h in enumerate(self.headers):
            col = [(row[j],) for row in v]
            blank = sum(1 for (x,) in col if x is None or (isinstance(x, str) and not x.strip()))
            text_num = numeric_text(col)
            junk = bad_text(self, j)
            junk = junk and f"날짜·숫자가 아닌 글자 {len(junk)}개({spots(col_letter(self.region.Column + j), [r for r, _ in junk])})"
            if blank or text_num or junk:
                notes.append(f"{h}: " + ", ".join(p for p in (blank and f"빈 칸 {blank}개", text_num and f"숫자처럼 보이는 글자 {text_num}개", junk) if p))
        return ["열별 주의할 값: " + "; ".join(notes)] if notes else []

    def outside(self, limit=40):
        """표 오른쪽에 이미 있는 내용(이전 작업 결과 등). 펼쳐진 수식은 실제 범위까지 알려줘야
        '방금 만든 표 맨 아래' 같은 후속 요청의 위치를 AI가 짐작하지 않는다."""
        r, used = self.region, self.ws.UsedRange
        c0, c1 = r.Column + r.Columns.Count, used.Column + used.Columns.Count - 1
        if c0 > c1:
            return []
        block = self.ws.Range(self.ws.Cells(used.Row, c0), self.ws.Cells(used.Row + used.Rows.Count - 1, c1))
        dyn = dynamic_arrays(self.ws)
        f = block.Formula2 if dyn else block.Formula
        f = f if isinstance(f, tuple) else ((f,),)
        lines = []
        for i, row in enumerate(f):
            for j, v in enumerate(row):
                if v in ("", None):
                    continue  # 펼쳐진 결과 칸은 수식이 비어 있어 여기서 빠지고 원래 수식 칸에서 범위로 알려줌
                cell = block.Cells(i + 1, j + 1)
                addr = cell.Address.replace("$", "")
                if dyn and str(v).startswith("=") and cell.HasSpill:
                    lines.append(f"{addr}: {v}  → 결과가 {cell.SpillingToRange.Address.replace('$', '')}에 펼쳐져 있음")
                else:
                    lines.append(f"{addr}: {v}")
                if len(lines) >= limit:
                    return ["표 오른쪽에 이미 있는 내용(일부):", *lines]
        return ["표 오른쪽에 이미 있는 내용:", *lines] if lines else []


def first_block(ws):
    """위에서부터 훑어 처음 만나는 2행 이상 이어진 범위. 맨 위 제목 한 줄은 건너뛴다. 없으면 A1의 범위."""
    used = ws.UsedRange
    v = used.Value2
    v = v if isinstance(v, tuple) else ((v,),)
    seen = []
    for i, row in enumerate(v):
        for j, x in enumerate(row):
            if x is None or x == "":
                continue
            r, c = used.Row + i, used.Column + j
            if any(r0 <= r < r1 and c0 <= c < c1 for r0, r1, c0, c1 in seen):
                continue
            block = ws.Cells(r, c).CurrentRegion
            if block.Rows.Count >= 2:
                return block
            seen.append((block.Row, block.Row + block.Rows.Count, block.Column, block.Column + block.Columns.Count))
    return ws.Range("A1").CurrentRegion


def dynamic_arrays(ws):
    """버전 문자열 대신 실제 속성으로 확인 (PRD FR-01)."""
    try:
        ws.Range("A1").Formula2
        return True
    except Exception:
        return False


BUSY = -2147418111  # RPC_E_CALL_REJECTED: 재계산·피벗 새로고침 직후 Excel이 바빠 COM 호출을 잠깐 거부함


def retry_busy(fn, *args):
    """읽기처럼 다시 해도 되는 호출만: Excel이 바빠 거부하면 0.25초씩 최대 약 10초 기다렸다 다시."""
    for _ in range(40):
        try:
            return fn(*args)
        except pywintypes.com_error as e:
            if e.hresult != BUSY:
                raise
            time.sleep(0.25)
    return fn(*args)


def fingerprint(ws):
    return retry_busy(_fingerprint, ws)


def _fingerprint(ws):
    r = ws.UsedRange
    # 컬렉션을 for로 돌면 Excel이 바쁠 때 pywin32가 TypeError로 바꿔 버려 재시도가 안 되므로 번호로 꺼냄
    los = [ws.ListObjects(i) for i in range(1, ws.ListObjects.Count + 1)]
    filt = ws.AutoFilterMode, [lo.AutoFilter.FilterMode if lo.AutoFilter else None for lo in los]
    raw = repr((r.Address, r.Formula2, r.Value2, filt, ws.PivotTables().Count,
                [ws.Parent.Worksheets(i).Name for i in range(1, ws.Parent.Worksheets.Count + 1)]))
    return hashlib.sha256(raw.encode()).hexdigest()


def execute(t, plan, progress=None):
    """검사를 통과한 계획을 실행하고 바뀐 내용을 설명 목록으로 돌려준다. 실패하면 PlanError."""
    stage = ["준비"]
    try:
        return _execute(t, plan, stage, progress)
    except PlanError:
        raise
    except Exception as e:
        # COM 오류 문구는 사용자가 읽을 수 없으므로 어느 단계였는지와 뜻을 알린다
        raise PlanError(f"{stage[0]} 단계에서 문제가 생겼습니다. {excel_error(e)}")


BUSY_CODES = {0x800AC472, 0x8001010A, 0x80010001}  # Excel이 셀 편집·대화상자·계산 중이라 호출을 안 받음
GONE_CODES = {0x80010108, 0x800706BA, 0x800706BE, 0x80010114}  # Excel이 닫혔거나 멈춰 연결이 끊김


def excel_error(e):
    """Excel(COM) 오류를 사용자가 할 일을 알 수 있는 문장으로. COM 오류가 아니면 원래 문구."""
    if isinstance(e, AttributeError) and re.fullmatch(r"\w+\.\w+", str(e)):
        # 셀 편집 중 Excel이 이름 조회를 거부하면 pywin32가 'PivotTables.Count' 같은 AttributeError로 바꿔 버림
        return "Excel에서 셀을 편집 중입니다. 입력 중인 셀에서 Enter나 Esc를 누른 뒤 다시 요청해주세요."
    if not isinstance(e, pywintypes.com_error):
        return str(e)
    info = e.excepinfo if isinstance(getattr(e, "excepinfo", None), tuple) else None
    code = (info[5] if info and len(info) > 5 and info[5] else e.hresult) & 0xFFFFFFFF
    if code in BUSY_CODES:
        return "Excel이 다른 작업 중입니다. 셀 편집이나 열린 창을 끝낸 뒤 다시 요청해주세요."
    if code in GONE_CODES:
        return "Excel과 연결이 끊겼습니다. Excel이 닫혔거나 멈췄는지 확인하고 새로고침을 눌러주세요."
    if code == 0x800A03EC:
        return "Excel이 이 작업을 받아들이지 않았습니다(수식, 주소, 범위 문제)."
    if code == 0x80070005:
        return "Excel이 접근을 거부했습니다. 시트 보호나 읽기 전용 여부를 확인해주세요."
    detail = info[2] if info and info[2] else ""
    return f"Excel 내부 오류(0x{code:08X}){': ' + detail if detail else ''}. 같은 문제가 계속되면 이 번호를 알려주세요."


def _execute(t, plan, stage, progress=None):
    ws, xl, done = t.ws, t.xl, []
    idx = {h: i + 1 for i, h in enumerate(t.headers)}
    progress = progress or (lambda _: None)

    # 원본 수정(사용자가 지시한 경우만): 이어지는 정렬·합계가 바뀐 숫자를 쓰도록 맨 먼저
    stage[0] = "숫자 변환"
    progress(stage[0])
    converted = 0
    for c in plan["to_number"]:
        if not plan["edit_source"]:
            raise PlanError("원본 수정 지시가 없어 원본 값을 바꾸지 않았습니다.")
        j = idx[c] - 1
        hits = rows_where(t, j, lambda x: _num(x) is not None)
        col = t.region.Column + j
        for n, (r, x) in enumerate(hits):
            if n % 100 == 0:
                progress(f"숫자 변환 {n + 1:,}/{len(hits):,}")
            cell = ws.Cells(r, col)
            if cell.NumberFormat == "@":
                cell.NumberFormat = "General"  # 텍스트 서식이면 숫자를 넣어도 다시 글자가 됨
            cell.Value2 = _num(x)
        converted += len(hits)
        if hits:
            done.append(f"원본 수정: '{c}' 열 글자로 저장된 숫자 {len(hits)}칸을 숫자로 바꿈"
                        + f" ({spots(col_letter(col), [r for r, _ in hits])}, {sum(_num(x) for _, x in hits):,.0f})")
    if plan["to_number"]:
        if not converted:
            raise PlanError("지정한 열에서 변환할 글자 숫자를 찾지 못해 원본을 바꾸지 않았습니다.")
        # 이 숫자를 쓰는 수식은 Excel이 다시 계산하지만 피벗은 새로고침해야 바뀐 값이 들어감
        pts = [pt for s in t.wb.Worksheets for pt in s.PivotTables()]
        for pt in pts:
            pt.PivotCache().Refresh()
        done.append("→ 이 값을 쓰는 수식은 다시 계산" + (f", 피벗 {len(pts)}개 새로고침" if pts else ""))

    # 필터를 먼저 걸면 Excel이 보이는 행만 정렬하므로 정렬이 먼저
    stage[0] = "정렬"
    progress(stage[0])
    if (plan["sorts"] or plan["filters"]) and t.spills():
        a = ", ".join(s[0] for s in t.spills())
        raise PlanError(f"이 표는 {a} 수식이 펼쳐진 결과라 Excel 정렬·필터를 쓸 수 없습니다. "
                        "'수식을 고쳐서 정렬해줘'처럼 요청하면 수식 자체를 바꿔 순서를 맞춥니다.")
    if plan["sorts"]:
        sorter = t.lo.Sort if t.lo else ws.Sort
        sorter.SortFields.Clear()
        for s in plan["sorts"]:
            key = t.lo.ListColumns(s["column"]).Range if t.lo else t.region.Columns(idx[s["column"]])
            sorter.SortFields.Add(Key=key, Order=2 if s["descending"] else 1)
        if not t.lo:
            sorter.SetRange(t.region)
        sorter.Header = 1  # 머리글 행은 정렬하지 않음, 행 단위로 함께 움직임
        sorter.Apply()
        done.append("정렬: " + ", ".join(f"{s['column']} {'내림차순' if s['descending'] else '오름차순'}" for s in plan["sorts"]))

    # 원본 값은 바꾸지 않고, 숫자로 취급되지 않은 칸이 있으면 알림 (PRD 5장: 사용한 가정 표시)
    numeric_ops = {"=", "<>", ">", ">=", "<", "<=", "top_items"}
    for c in dict.fromkeys([s["column"] for s in plan["sorts"]] + [f["column"] for f in plan["filters"] if f["op"] in numeric_ops]):
        hits = text_numbers(t, idx[c] - 1)
        if hits:
            where = spots(col_letter(t.region.Column + idx[c] - 1), [r for r, _ in hits])
            done.append(f"⚠ 임의 판단: '{c}' 열의 글자로 저장된 숫자 {len(hits)}칸({where})은 정렬·비교에서 숫자로 취급되지 않았습니다.")

    stage[0] = "필터"
    progress(stage[0])
    for f in plan["filters"]:
        value = f["value"]
        if f["op"] == "=" and isinstance(value, str):
            # 사용자가 띄어쓰기를 빼고 쓴 값('확인필요')이 칸의 실제 값('확인 필요')과 다르면 실제 값으로
            cells = {x for _, x in rows_where(t, idx[f["column"]] - 1, lambda x: isinstance(x, str))}
            squash = lambda s: "".join(s.split()).casefold()
            if value not in cells:
                same = [x for x in cells if squash(x) == squash(value)]
                if len(same) == 1:
                    value = same[0]
                    done.append(f"임의 판단: '{f['value']}'은 '{value}'로 보고 걸렀습니다(띄어쓰기 차이).")
        condition = apply_filter(t.region, idx[f["column"]], f["op"], value)
        done.append(f"필터: {f['column']} {condition}")
    if plan["filters"]:
        shown = int(xl.WorksheetFunction.Subtotal(103, t.region.Columns(1))) - 1
        if shown <= 0:  # 전부 숨기면 데이터가 사라진 것처럼 보이므로 걸지 않음(호출한 쪽이 되돌림)
            raise PlanError("필터 조건에 맞는 행이 없어 필터를 걸지 않았습니다. 열에 실제로 있는 값과 정확히 같은지 확인해주세요.")
        done.append(f"→ 조건에 맞아 보이는 행 {shown}개")

    stage[0] = "셀 입력"
    progress(stage[0])
    if plan["writes"]:
        out_ws = new_sheet(t, plan["new_sheet"]) if plan["new_sheet"] else ws
        ranges = [out_ws.Range(w["address"]) for w in plan["writes"]]
        if out_ws is ws:
            for w, r in zip(plan["writes"], ranges):
                if xl.Intersect(r, t.region) is not None:
                    if not plan["edit_source"]:  # 원본 표 안은 사용자가 원본 수정을 지시했을 때만
                        raise PlanError(f"{w['address']}는 원본 표 안이라 쓰지 않았습니다. 원본을 고치려면 '원본을 수정해줘'라고 말씀해주세요.")
                elif not plan["overwrite"] and xl.WorksheetFunction.CountA(r):
                    raise PlanError(f"{w['address']}에 이미 데이터가 있어 덮어쓰지 않았습니다. 덮어쓰려면 그렇게 말씀해주세요.")
        dyn = dynamic_arrays(ws)
        rebuilt = out_ws is ws and rebuilds(t, plan)
        if rebuilt:
            stale = stale_refs(t, plan["writes"], ranges)
            if stale:  # 곧 비울 옛 표를 읽는 수식은 빈 칸·0만 낸다 (예: A2에 =TRANSPOSE(A2:C5))
                raise FormulaError(f"새 수식이 곧 비울 기존 표({t.region.Address.replace('$', '')})를 참조합니다: "
                                   + ", ".join(stale) + ". 기존 표 대신 원본 데이터에서 다시 계산해야 합니다.")
            t.region.ClearContents()
            done.append(f"기존 표 {t.region.Address.replace('$', '')}를 비우고 새로 씀")
        for w, r in zip(plan["writes"], ranges):
            progress(f"셀 입력: {w['address']}")
            if w["content"].startswith("="):
                if len(w["content"]) > 8000:  # Excel 수식 길이 한도 8192자
                    raise FormulaError(f"{w['address']} 수식이 {len(w['content']):,}자로 Excel 한도를 넘습니다. 더 짧은 수식으로 만들어야 합니다.")
                if out_ws is ws and xl.Intersect(r, t.region) is not None and in_place(ws, xl, r, w["content"]):
                    continue
                try:
                    if dyn:
                        r.Formula2 = w["content"]
                    else:
                        r.Formula = w["content"]  # 동적 배열이 없는 Excel
                except pywintypes.com_error as e:
                    # 0x800A03EC 등: Excel이 수식 자체를 거부(문법 오류·없는 함수·잘못된 참조). 번호 대신 칸과 수식을 알려 AI가 고치게 함
                    raise FormulaError(f"Excel이 {w['address']} 수식을 받아들이지 않았습니다(문법 오류, 지원하지 않는 함수, 잘못된 참조 등).\n"
                                       f"  {w['address']} {w['content'][:200]}") from e
            elif "\t" in w["content"] or "\n" in w["content"]:
                # 머리글처럼 탭·줄바꿈으로 이어 붙인 글자는 칸마다 나눠 씀(범위 전체에 같은 글자가 들어가지 않게)
                grid = [line.split("\t") for line in w["content"].split("\n")]
                width = max(len(g) for g in grid)
                grid = tuple(tuple(g + [""] * (width - len(g))) for g in grid)
                block = out_ws.Range(r.Cells(1, 1), r.Cells(len(grid), width))
                if any(c.startswith("=") for g in grid for c in g):
                    # 글자와 수식이 섞였으면 칸마다 수식으로 넣어야 펼쳐진다(Value2로 넣으면 한 칸 값만 남음)
                    for i, g in enumerate(grid):
                        for j, c in enumerate(g):
                            cell = block.Cells(i + 1, j + 1)
                            if c.startswith("=") and dyn:
                                cell.Formula2 = c
                            else:
                                cell.Formula = c
                else:
                    block.Value2 = grid
            else:
                r.Value2 = w["content"]
        xl.Calculate()
        progress("수식 계산 완료")
        try:
            circ = out_ws.CircularReference
        except Exception:
            circ = None
        if circ is not None:  # 순환 참조는 오류 대신 0을 내서 아래 오류 검사에 걸리지 않음
            # Excel은 순환 고리 중 한 칸만 알려주므로 그 칸 수식이 원인이라고 단정하지 않는다
            raise FormulaError(f"새 수식들이 서로 또는 자기 결과를 다시 참조합니다(순환 참조, Excel이 표시한 칸: "
                               f"{circ.Address.replace('$', '')}).\n  쓴 수식: "
                               + "; ".join(f"{w['address']} {w['content'][:120]}" for w in plan["writes"] if w["content"].startswith("=")))
        errs, broken = {}, []
        for w, r in zip(plan["writes"], ranges):
            out = r.SpillingToRange if r.Count == 1 and dyn and r.HasSpill else r
            found = errors_in(out.Value2)
            for k, v in found.items():
                errs[k] = errs.get(k, 0) + v
            # 결과 칸이 전부 오류면 데이터 탓이 아니라 수식이 잘못된 것
            if w["content"].startswith("=") and sum(found.values()) == out.Count:
                broken.append(f"{w['address']} {w['content'][:80]} → {', '.join(found)}")
        fatal = {k: v for k, v in errs.items() if k not in DATA_ERRORS}
        if fatal or broken:
            detail = ", ".join(f"{k} {v}칸" for k, v in errs.items())
            raise FormulaError(f"수식이 오류를 냈습니다({detail}). 결과 칸이 모두 오류이거나 지원하지 않는 함수, "
                               "결과가 들어갈 자리가 막혀 있을 수 있습니다." + "".join(f"\n  {b}" for b in broken))
        if out_ws is not ws:
            out_ws.Columns.AutoFit()
            ws.Activate()
            done.append(f"새 시트 '{out_ws.Name}'에 작성")
        else:  # 결과가 열 폭보다 길면 5.24E+10·#### 로 보이므로 넓히기만 함(원래 더 넓으면 그대로)
            for r in ranges:
                out = r.SpillingToRange if r.Count == 1 and dyn and r.HasSpill else r
                inherit_format(t, out)
                for k in range(1, out.Columns.Count + 1):
                    col = out.Columns(k)
                    old = col.ColumnWidth
                    col.AutoFit()
                    if col.ColumnWidth < old:
                        col.ColumnWidth = old
        if errs:
            done.append("⚠ 일부 칸 오류(데이터 때문일 수 있음): " + ", ".join(f"{k} {v}칸" for k, v in errs.items()))
        if plan["edit_source"]:
            done += [f"원본 수정: {w['address']} ← {w['content'][:60]}" for w, r in zip(plan["writes"], ranges)
                     if out_ws is ws and xl.Intersect(r, t.region) is not None]
        if not rebuilt:  # 비운 옛 표의 빈 칸을 원본 문제로 적지 않게
            done += assumption_notes(t, [w["content"] for w in plan["writes"] if w["content"].startswith("=")], out_ws is ws)

    stage[0] = "표시 형식"
    progress(stage[0])
    for f in plan["formats"]:
        ws.Range(f["address"]).NumberFormat = f["format"]  # 값·수식은 그대로, 보이는 모양만 바뀜
        done.append(f"표시 형식: {f['address']} → {f['format']} (실제 값은 그대로)")

    stage[0] = "색 표시"
    progress(stage[0])
    for h in plan["highlights"]:
        done += highlight(t, h)

    stage[0] = "피벗"
    progress(stage[0])
    if plan["pivot"]:
        p = plan["pivot"]
        names = {s.Name for s in t.wb.Worksheets}
        name = next(f"AI 피벗{i}" for i in range(1, 1000) if f"AI 피벗{i}" not in names)
        new = t.wb.Worksheets.Add(After=ws)
        new.Name = name
        pt = t.wb.PivotCaches().Create(1, t.region).CreatePivotTable(new.Range("A3"), name.replace(" ", ""))
        for c in p["rows"]:
            pt.PivotFields(c).Orientation = 1
        for c in p["columns"]:
            pt.PivotFields(c).Orientation = 2
        for v in p["values"]:
            pt.AddDataField(pt.PivotFields(v["field"]), f"{v['field']} {v['func']}", PIVOT_FUNC[v["func"]])
        ws.Activate()
        done.append(f"피벗: 새 시트 '{name}' (행 {', '.join(p['rows']) or '-'} / 열 {', '.join(p['columns']) or '-'})")
        done += pivot_notes(t, p)
    return done


COLORS = {"yellow": 0x99FFFF, "red": 0x9999FF, "green": 0x99E6B3, "blue": 0xF7D6B4, "orange": 0x80C0FF, "gray": 0xD9D9D9}  # 연한 색 (BGR)
COLOR_NAMES = {"yellow": "노란색", "red": "빨간색", "green": "초록색", "blue": "파란색", "orange": "주황색", "gray": "회색"}


def highlight(t, h):
    """조건부 서식: 수식이 TRUE인 칸에 배경색. 값은 그대로. 칠해질 칸을 미리 계산해 칸 수·위치를 알린다
    (기준 칸이 어긋난 수식은 오류 없이 엉뚱한 칸을 칠하므로 사용자가 위치로 확인하게)."""
    ws, xl, f = t.ws, t.xl, h["formula"]
    r = ws.Range(h["address"])
    used = ws.UsedRange
    c0 = used.Column + used.Columns.Count + 1
    tmp = ws.Range(ws.Cells(r.Row, c0), ws.Cells(r.Row + r.Rows.Count - 1, c0 + r.Columns.Count - 1))
    try:
        try:
            tmp.Formula = f  # 첫 칸 기준 상대 참조라 다른 열에 써도 조건부 서식과 같은 칸을 읽음
        except pywintypes.com_error as e:
            raise FormulaError(f"Excel이 {h['address']} 색 표시 조건 수식을 받아들이지 않았습니다.\n  {f[:200]}") from e
        xl.Calculate()
        v = tmp.Value2
    finally:
        tmp.Clear()
    v = v if isinstance(v, tuple) else ((v,),)
    errs = errors_in(v)
    if errs and sum(errs.values()) == r.Count:
        raise FormulaError(f"색 표시 조건 수식이 모든 칸에서 오류를 냈습니다({', '.join(errs)}).\n  {h['address']} {f[:200]}")
    hits = {}
    for i, row in enumerate(v):
        for j, x in enumerate(row):
            if x is True or (type(x) in (int, float) and x and not errors_in(x)):
                hits.setdefault(r.Column + j, []).append(r.Row + i)
    # A1 수식은 범위 첫 칸이 아니라 '현재 선택 칸' 기준으로 해석되므로, 첫 칸 기준 R1C1로 바꿔 R1C1 모드에서 넣음
    rc = xl.ConvertFormula(f, 1, -4150, None, r.Cells(1, 1))
    style = xl.ReferenceStyle
    xl.ReferenceStyle = -4150
    try:
        fc = r.FormatConditions.Add(Type=2, Formula1=rc)  # 2 = 수식 조건
    finally:
        xl.ReferenceStyle = style
    fc.Interior.Color = COLORS[h["color"]]
    n = sum(len(rows) for rows in hits.values())
    out = [f"색 표시: {h['address']} 중 조건 {f} 에 맞는 {n}칸을 {COLOR_NAMES[h['color']]}으로 (조건부 서식, 값은 그대로)"]
    out += [f"  {col_letter(c)}열 {len(rows)}칸: {spots(col_letter(c), rows)}" for c, rows in list(hits.items())[:6]]
    if len(hits) > 6:
        out.append(f"  외 {len(hits) - 6}개 열")
    if not n:
        out.append("⚠ 지금은 조건에 맞는 칸이 없습니다. 조건이 맞는지 확인해주세요.")
    return out


def in_place(ws, xl, r, formula):
    """원본 칸을 자기 값으로 고치는 수식(예: I2:I877 =IF($I2="확인 필요",...,$I2))은 그 자리에 두면 순환 참조라,
    빈 열에서 계산해 값만 붙인다. 자기 칸을 읽지 않는 수식이면 False(그대로 수식으로 씀)."""
    used = ws.UsedRange
    c0 = used.Column + used.Columns.Count + 1
    tmp = ws.Range(ws.Cells(r.Row, c0), ws.Cells(r.Row + r.Rows.Count - 1, c0 + r.Columns.Count - 1))
    try:
        try:
            tmp.Formula = formula  # 글자 주소는 첫 칸 기준이라 다른 열에 써도 같은 칸을 읽음
        except pywintypes.com_error as e:
            raise FormulaError(f"Excel이 {r.Address.replace('$', '')} 수식을 받아들이지 않았습니다.\n  {formula[:200]}") from e
        xl.Calculate()
        try:
            own = xl.Intersect(tmp.Cells(1, 1).DirectPrecedents, r) is not None
        except pywintypes.com_error:  # 참조하는 칸이 없음
            own = False
        if own:
            tmp.Copy()
            r.PasteSpecial(Paste=-4163)  # 값만 (오류 값도 그대로)
            xl.CutCopyMode = False
    finally:
        tmp.Clear()
    return own


def inherit_format(t, out):
    """표 바로 아래 행(열별 합계)·바로 오른쪽 열(행별 합계)에 붙여 쓴 칸은 표 숫자 형식을 이어받는다.
    (예: 표가 1,000 단위로 보이는데 합계만 원래 숫자로 보이면 합계가 안 맞아 보임) 이미 형식이 있는 칸은 그대로."""
    r, ws = t.region, t.ws
    top, bottom, left, right = r.Row + 1, r.Row + r.Rows.Count - 1, r.Column, r.Column + r.Columns.Count - 1
    if bottom < top:
        return
    o_top, o_left = out.Row, out.Column
    o_bottom, o_right = o_top + out.Rows.Count - 1, o_left + out.Columns.Count - 1

    general = {"General", "G/표준", t.wb.Styles("Normal").NumberFormat}  # 한국어 Excel은 기본 형식 이름이 'G/표준'

    def copy(cell_from, row, col):
        cell = ws.Cells(row, col)
        if cell.NumberFormat in general:
            cell.NumberFormat = cell_from.NumberFormat

    if o_top <= bottom + 1 <= o_bottom:  # 열별 합계 행: 각 열은 그 열 첫 데이터 칸의 형식
        for c in range(max(left, o_left), min(right, o_right) + 1):
            copy(ws.Cells(top, c), bottom + 1, c)
    if o_left <= right + 1 <= o_right:  # 행별 합계 열(과 아래 모서리): 마지막 데이터 열의 형식
        for row in range(max(top, o_top), min(bottom + 1, o_bottom) + 1):
            copy(ws.Cells(min(row, bottom), right), row, right + 1)


def new_sheet(t, name):
    """작업 시트 바로 뒤에 새 시트를 만든다. 이름은 Excel 규칙(31자, []:*?/\\ 금지)에 맞추고 겹치면 (2)…를 붙인다."""
    base = re.sub(r"[\[\]:*?/\\']", "", name).strip()[:28] or "AI 결과"
    names = {s.Name for s in t.wb.Worksheets}
    name = next(n for n in [base] + [f"{base}({i})" for i in range(2, 1000)] if n not in names)
    new = t.wb.Worksheets.Add(After=t.ws)
    new.Name = name
    return new


class Undo:
    """직전 AI 작업 1단계 되돌리기 (spike/undo_check.py에서 1만 행으로 검증한 방식)."""

    def __init__(self, t):
        self.ws, self.wb, self.xl = t.ws, t.wb, t.xl

    @staticmethod
    def _filters(af):
        if af is None:
            return []
        saved = []
        for i in range(1, af.Filters.Count + 1):
            f = af.Filters(i)
            if not f.On:
                continue
            args = {"Field": i, "Criteria1": f.Criteria1}
            try:
                op = f.Operator
                if op:
                    args["Operator"] = op
            except Exception:
                pass
            try:
                args["Criteria2"] = f.Criteria2
            except Exception:
                pass
            saved.append(args)
        return saved

    def before(self):
        self.drop()
        self.sheets = [s.Name for s in self.wb.Worksheets]
        try:
            self.table_specs = [(lo.Name, lo.Range.Address, lo.TableStyle) for lo in self.ws.ListObjects]
            self.sheet_filter = (self.ws.AutoFilter.Range.Address,
                                 self._filters(self.ws.AutoFilter)) if self.ws.AutoFilterMode else None
            self.table_filters = {lo.Name: self._filters(lo.AutoFilter) for lo in self.ws.ListObjects}
            # shortcut: 숨김 사본이 사용자 문서 안에 있어 그 사이 저장하면 같이 저장됨, 다음 작업·되돌리기 때 지워짐
            self.ws.Copy(After=self.wb.Worksheets(self.wb.Worksheets.Count))
            b = self.wb.Worksheets(self.wb.Worksheets.Count)
            b.Name, b.Visible = BACKUP, 2  # 2 = 사용자에게 안 보이는 숨김
            self.ws.Activate()
        except Exception as e:
            raise PlanError(f"되돌리기용 사본을 만들지 못해 실행하지 않았습니다. {excel_error(e)}")

    def after(self):
        self.done = fingerprint(self.ws)

    def undo(self, force=False):
        if not force and fingerprint(self.ws) != self.done:
            raise PlanError("AI 작업 뒤에 직접 수정한 내용이 있어 되돌리지 않았습니다. 직접 수정한 내용을 덮어쓰지 않기 위해서입니다.")
        self.xl.DisplayAlerts = False
        try:
            for s in list(self.wb.Worksheets):
                if s.Name not in self.sheets and s.Name != BACKUP:
                    s.Delete()
        finally:
            self.xl.DisplayAlerts = True
        for pt in list(self.ws.PivotTables()):
            pt.TableRange2.Clear()
        for lo in list(self.ws.ListObjects):
            if lo.AutoFilter and lo.AutoFilter.FilterMode:
                lo.AutoFilter.ShowAllData()
            lo.Unlist()  # 표를 남긴 채 붙이면 충돌하므로 일반 범위로 바꾼 뒤 비움
        self.ws.AutoFilterMode = False
        self.ws.Cells.Clear()  # Delete는 다른 시트 수식의 참조를 #REF!로 깨뜨려서 Clear를 씀
        backup = self.wb.Worksheets(BACKUP)
        # 필터가 걸린 채 사본을 떴으면 사본의 숨은 행은 Copy가 건너뛰어 되돌리면 보이던 행만 남으므로 모두 펼쳐서 붙임.
        # 표(ListObject)의 필터는 AutoFilterMode·Rows.Hidden으로 안 풀려서 표마다 따로 푼다(안 풀면 숨은 행 데이터가 사라짐)
        for lo in backup.ListObjects:
            if lo.AutoFilter and lo.AutoFilter.FilterMode:
                lo.AutoFilter.ShowAllData()
        backup.AutoFilterMode = False
        backup.Rows.Hidden = False
        backup.Cells.Copy(self.ws.Cells)
        skipped = []
        for name, address, style in self.table_specs:
            # 붙이면 표가 새 이름으로 따라오므로 같은 자리의 표를 원래 이름으로(다른 수식의 표 참조가 깨지지 않게), 없으면 다시 만듦
            lo = next((item for item in self.ws.ListObjects if item.Range.Address == address), None)
            if lo is None:
                lo = self.ws.ListObjects.Add(1, self.ws.Range(address), None, 1)
                lo.TableStyle = style
            lo.Name = name
            skipped += self._reapply(lo.Range, self.table_filters.get(name, []))
        if self.sheet_filter:
            address, filters = self.sheet_filter
            region = self.ws.Range(address)
            if not self.ws.AutoFilterMode:
                region.AutoFilter(Field=1)  # 조건 없이 화살표만 켬(인수 없는 AutoFilter()는 pywin32에서 오류)
            skipped += self._reapply(region, filters)
        self.drop()
        return skipped

    @staticmethod
    def _reapply(rng, filters):
        """필터 조건을 다시 건다. 색·아이콘 필터처럼 다시 걸 수 없는 조건은 건너뛰고 그 열 번호를 돌려준다(값 복원은 이미 끝남)."""
        skipped = []
        for args in filters:
            try:
                rng.AutoFilter(**args)
            except pywintypes.com_error:
                skipped.append(args["Field"])
        return skipped

    def drop(self):
        self.xl.DisplayAlerts = False
        try:
            for s in list(self.wb.Worksheets):
                if s.Name == BACKUP:
                    s.Visible = -1
                    s.Delete()
        finally:
            self.xl.DisplayAlerts = True


def column_values(t, column):
    """표의 한 열 데이터(머리글 제외)를 행 순서대로."""
    v = t.region.Columns(t.headers.index(column) + 1).Value2
    return [r[0] for r in v[1:]]


def write_classification(t, column, rows, overwrite, output="", edit_source=False):
    """분류 결과를 표 바로 오른쪽 두 열(분류, 분류 이유)에 쓴다. 표에 붙여야 정렬할 때 행과 함께 움직인다.
    output(열 글자)을 주면 그 한 열에만 범주를 쓴다. 같은 이름의 열이 이미 있으면(다시 분류) 그 열을 갱신한다.
    돌려주는 값: 기록한 위치 설명."""
    from .classify import CHECK, FAILED
    names = [f"{column} 분류", f"{column} 분류 이유"]
    ws, n = t.ws, len(rows)
    if output:
        c = _col_index(output)
        names, cols = names[:1], [c]
        block = ws.Range(ws.Cells(t.region.Row, c), ws.Cells(t.region.Row + n, c))
        own = ws.Cells(t.region.Row, c).Value2 == names[0]  # 셀리가 전에 만든 분류 열이면 다시 분류해 갱신
        if _in_table(t, c) and not (edit_source or own):
            raise PlanError(f"{col_letter(c)}열은 원본 표 안이라 분류 결과를 쓰지 않았습니다. "
                            "빈 열을 지정하거나, 원본을 수정해도 된다고 말씀해주세요.")
        if not (overwrite or edit_source or own) and t.xl.WorksheetFunction.CountA(block):
            raise PlanError(f"{col_letter(c)}열에 이미 내용이 있어 분류 결과를 쓰지 않았습니다. "
                            "비어 있는 열을 지정하거나, 덮어써도 된다고 말씀해주세요.")
        ws.Cells(t.region.Row, c).Value2 = names[0]
        top = t.region.Row + 1
        ws.Range(ws.Cells(top, c), ws.Cells(top + n - 1, c)).Value2 = tuple((r[0],) for r in rows)
        ws.Range(ws.Cells(top, c), ws.Cells(top + n - 1, c)).Interior.ColorIndex = -4142
        for i, (cat, _) in enumerate(rows):
            if cat in (CHECK, FAILED):
                ws.Cells(top + i, c).Interior.Color = 0x99E6FF if cat == CHECK else 0xC0C0FF
        return f"{col_letter(c)}열 '{names[0]}'"
    if all(h in t.headers for h in names):
        cols = [t.region.Column + t.headers.index(h) for h in names]
    else:
        start = t.region.Column + t.region.Columns.Count
        cols = [start, start + 1]
        block = ws.Range(ws.Cells(t.region.Row, start), ws.Cells(t.region.Row + n, start + 1))
        if not overwrite and t.xl.WorksheetFunction.CountA(block):
            raise PlanError(f"표 오른쪽 {col_letter(start)}~{col_letter(start + 1)}열에 이미 내용이 있어 분류 결과를 쓰지 않았습니다. "
                            "그 열을 비우거나, 덮어써도 된다고 말씀해주세요.")
        if t.lo:  # 표(ListObject)는 열을 추가해야 표 범위에 포함된다
            for h in names:
                t.lo.ListColumns.Add().Name = h
        else:
            for c, h in zip(cols, names):
                ws.Cells(t.region.Row, c).Value2 = h
    top = t.region.Row + 1
    for c, k in zip(cols, (0, 1)):
        ws.Range(ws.Cells(top, c), ws.Cells(top + n - 1, c)).Value2 = tuple((r[k],) for r in rows)
    # 확인 필요·실패 행은 눈에 띄게 (원래 서식은 되돌리기 사본에 있음)
    target = ws.Range(ws.Cells(top, cols[0]), ws.Cells(top + n - 1, cols[0]))
    target.Interior.ColorIndex = -4142
    for i, (cat, _) in enumerate(rows):
        if cat in (CHECK, FAILED):
            ws.Cells(top + i, cols[0]).Interior.Color = 0x99E6FF if cat == CHECK else 0xC0C0FF  # 연노랑 / 연빨강 (BGR)
    ws.Columns(cols[1]).ColumnWidth = 36
    return f"{col_letter(cols[0])}열 '{names[0]}', {col_letter(cols[1])}열 '{names[1]}'"
