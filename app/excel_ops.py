"""Excel 쪽: 대상 표 읽기, AI 계획 검사, 실행, 직전 작업 되돌리기."""
import hashlib, re

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
    for w in plan["writes"]:
        if not ADDRESS.match(w["address"].upper()):
            bad.append(f"'{w['address']}'는 이 시트의 셀 주소가 아닙니다.")
        if BLOCKED_FN.search(w["content"]) or EXTERNAL.search(w["content"]):
            bad.append(f"'{w['content'][:60]}'는 외부 연결·다른 파일·추가 기능이 필요한 수식이라 실행하지 않습니다.")
    bad += [f"'{c}' 열이 표에 없습니다." for c in plan["to_number"] if c not in headers]
    if plan["to_number"] and not plan["edit_source"]:
        bad.append("원본 수정 지시 없이 원본 값을 바꾸려 했습니다.")
    if plan["action"] == "change" and not (plan["filters"] or plan["sorts"] or plan["writes"] or plan["pivot"] or plan["classify"]
                                           or plan["to_number"]):
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
    """글자로 저장된 숫자면 그 값, 아니면 None."""
    if not isinstance(v, str):
        return None
    try:
        return float(v.replace(",", ""))
    except ValueError:
        return None


def rows_where(t, j, pred):
    """표의 j번째(0부터) 열에서 조건에 맞는 칸의 (행 번호, 값) 목록."""
    v = t.region.Columns(j + 1).Value2
    v = v if isinstance(v, tuple) else ((v,),)
    return [(t.region.Row + i, x) for i, (x,) in enumerate(v) if i and pred(x)]


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
        for w in plan["writes"]:
            if ADDRESS.match(w["address"].upper()) and t.xl.Intersect(t.ws.Range(w["address"]), t.region) is not None:
                lines.append(f"{w['address']} ← {w['content'][:60]}")
    c = plan["classify"]
    if c and c["output"] and _in_table(t, _col_index(c["output"])) \
            and t.ws.Cells(t.region.Row, _col_index(c["output"])).Value2 != f"{c['column']} 분류":
        lines.append(f"{c['output'].upper()}열을 분류 결과로 덮어씀")
    return lines


def _col_index(letters):
    return sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(letters.upper())))


def _in_table(t, c):
    return t.region.Column <= c < t.region.Column + t.region.Columns.Count


def assumption_notes(t, formulas, same_sheet):
    """수식이 참조한 원본 열에 글자 숫자·빈 칸이 있으면 셀리가 어떻게 다뤘는지 건수·위치·금액으로 적는다."""
    # shortcut: 열 참조와 '--'/VALUE 포함 여부를 글자로 판단, 수식이 복잡해져 오판이 보이면 Excel로 직접 검산
    text = " ".join(formulas)
    converts = bool(re.search(r"--|\bN?VALUE\(|NUMBERVALUE\(", text, re.I))
    notes = []
    for j, h in enumerate(t.headers):
        letter = col_letter(t.region.Column + j)
        ref = rf"(?<![A-Za-z_])\$?{letter}\$?\d" if same_sheet else rf"!\$?{letter}\$?\d"
        if not re.search(ref, text):
            continue
        tn = rows_where(t, j, lambda x: _num(x) is not None)
        if tn:
            amt = f"{sum(_num(x) for _, x in tn):,.0f}"
            where = spots(letter, [r for r, _ in tn])
            notes.append(f"⚠ 임의 판단: {letter}열 '{h}'의 글자로 저장된 숫자 {len(tn)}칸({where})을 "
                         + (f"숫자로 바꿔 포함 → 원본 {letter}열 SUM보다 {amt} 많음" if converts else f"제외 → 합계에서 {amt} 빠짐"))
        bl = rows_where(t, j, blank)
        grouped = "(빈 값)" in text and re.search(rf"\$?{letter}\$?\d+(:\$?{letter}\$?\d+)?\s*=\s*\"\"", text)
        if bl and grouped:
            notes.append(f"⚠ 임의 판단: {letter}열 '{h}' 빈 칸 {len(bl)}칸({spots(letter, [r for r, _ in bl])})을 '(빈 값)' 묶음으로 따로 집계")
        elif bl:
            notes.append(f"참고: {letter}열 '{h}' 빈 칸 {len(bl)}칸({spots(letter, [r for r, _ in bl])})")
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


class Target:
    """요청 시점에 고정한 통합 문서·시트와 그 안의 데이터 표."""

    def __init__(self, ws):
        self.ws, self.wb, self.xl = ws, ws.Parent, ws.Application
        self.lo = self.ws.ListObjects(1) if self.ws.ListObjects.Count else None
        # shortcut: 시트의 첫 표(없으면 첫 피벗, 그것도 없으면 위에서부터 처음 만나는 2행 이상 범위)만 대상으로 봄, 여러 표 지원 시 선택 범위로 확장
        if self.lo:
            self.region = self.lo.Range
        elif self.ws.PivotTables().Count:
            self.region = self.ws.PivotTables(1).TableRange1
        else:
            self.region = first_block(self.ws)
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
            "열: " + ", ".join(cols),
            f"결과 시작 열: {free} (그 왼쪽은 사용 중)",
            f"Excel 버전 {self.xl.Version}, 동적 배열 함수 {'지원함' if dynamic_arrays(self.ws) else '지원 안 함'}",
            *self.quality(),
            "샘플 5행:", *sample,
            *self.outside(),
        ])

    def quality(self):
        """열별 빈 칸·글자로 저장된 숫자 수. AI가 이를 모르면 집계에서 그 행을 조용히 빠뜨린다."""
        v = self.region.Value2[1:]
        notes = []
        for j, h in enumerate(self.headers):
            col = [(row[j],) for row in v]
            blank = sum(1 for (x,) in col if x is None or (isinstance(x, str) and not x.strip()))
            text_num = numeric_text(col)
            if blank or text_num:
                notes.append(f"{h}: " + ", ".join(p for p in (blank and f"빈 칸 {blank}개", text_num and f"숫자처럼 보이는 글자 {text_num}개") if p))
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


def fingerprint(ws):
    r = ws.UsedRange
    filt = ws.AutoFilterMode, [lo.AutoFilter.FilterMode if lo.AutoFilter else None for lo in ws.ListObjects]
    raw = repr((r.Address, r.Formula2, r.Value2, filt, ws.PivotTables().Count, [s.Name for s in ws.Parent.Worksheets]))
    return hashlib.sha256(raw.encode()).hexdigest()


def execute(t, plan):
    """검사를 통과한 계획을 실행하고 바뀐 내용을 설명 목록으로 돌려준다. 실패하면 PlanError."""
    stage = ["준비"]
    try:
        return _execute(t, plan, stage)
    except PlanError:
        raise
    except Exception as e:
        # COM 오류 문구는 사용자가 읽을 수 없으므로 어느 단계였는지 알린다
        raise PlanError(f"{stage[0]} 단계에서 Excel이 작업을 거부했습니다. ({getattr(e, 'excepinfo', None) and e.excepinfo[2] or e})")


def _execute(t, plan, stage):
    ws, xl, done = t.ws, t.xl, []
    idx = {h: i + 1 for i, h in enumerate(t.headers)}

    # 원본 수정(사용자가 지시한 경우만): 이어지는 정렬·합계가 바뀐 숫자를 쓰도록 맨 먼저
    stage[0] = "숫자 변환"
    for c in plan["to_number"]:
        if not plan["edit_source"]:
            raise PlanError("원본 수정 지시가 없어 원본 값을 바꾸지 않았습니다.")
        j = idx[c] - 1
        hits = rows_where(t, j, lambda x: _num(x) is not None)
        col = t.region.Column + j
        for r, x in hits:
            cell = ws.Cells(r, col)
            if cell.NumberFormat == "@":
                cell.NumberFormat = "General"  # 텍스트 서식이면 숫자를 넣어도 다시 글자가 됨
            cell.Value2 = _num(x)
        done.append(f"원본 수정: '{c}' 열 글자로 저장된 숫자 {len(hits)}칸을 숫자로 바꿈"
                    + (f" ({spots(col_letter(col), [r for r, _ in hits])})" if hits else ""))

    # 필터를 먼저 걸면 Excel이 보이는 행만 정렬하므로 정렬이 먼저
    stage[0] = "정렬"
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
    for c in dict.fromkeys([s["column"] for s in plan["sorts"]] + [f["column"] for f in plan["filters"] if f["op"] in "<>=<=>="]):
        hits = rows_where(t, idx[c] - 1, lambda x: _num(x) is not None)
        if hits:
            where = spots(col_letter(t.region.Column + idx[c] - 1), [r for r, _ in hits])
            done.append(f"⚠ 임의 판단: '{c}' 열의 글자로 저장된 숫자 {len(hits)}칸({where})은 정렬·비교에서 숫자로 취급되지 않았습니다. "
                        "원본을 숫자로 바꾸려면 '원본 숫자로 바꿔줘'라고 말씀해주세요.")

    stage[0] = "필터"
    for f in plan["filters"]:
        crit = CRITERIA[f["op"]].format(f["value"])
        t.region.AutoFilter(Field=idx[f["column"]], Criteria1=crit)
        done.append(f"필터: {f['column']} {f['op']} {f['value']}".rstrip())
    if plan["filters"]:
        shown = int(xl.WorksheetFunction.Subtotal(103, t.region.Columns(1))) - 1
        done.append(f"→ 조건에 맞아 보이는 행 {shown}개")

    stage[0] = "셀 입력"
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
        for w, r in zip(plan["writes"], ranges):
            if w["content"].startswith("=") and dyn:
                r.Formula2 = w["content"]
            elif w["content"].startswith("="):
                r.Formula = w["content"]  # 동적 배열이 없는 Excel
            else:
                r.Value2 = w["content"]
        xl.Calculate()
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
        if errs:
            done.append("⚠ 일부 칸 오류(데이터 때문일 수 있음): " + ", ".join(f"{k} {v}칸" for k, v in errs.items()))
        if plan["edit_source"]:
            done += [f"원본 수정: {w['address']} ← {w['content'][:60]}" for w, r in zip(plan["writes"], ranges)
                     if out_ws is ws and xl.Intersect(r, t.region) is not None]
        done += assumption_notes(t, [w["content"] for w in plan["writes"] if w["content"].startswith("=")], out_ws is ws)

    stage[0] = "피벗"
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
    return done


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

    def before(self):
        self.drop()
        self.sheets = [s.Name for s in self.wb.Worksheets]
        self.tables = [lo.Name for lo in self.ws.ListObjects]
        try:
            # shortcut: 숨김 사본이 사용자 문서 안에 있어 그 사이 저장하면 같이 저장됨, 다음 작업·되돌리기 때 지워짐
            self.ws.Copy(After=self.wb.Worksheets(self.wb.Worksheets.Count))
            b = self.wb.Worksheets(self.wb.Worksheets.Count)
            b.Name, b.Visible = BACKUP, 2  # 2 = 사용자에게 안 보이는 숨김
            self.ws.Activate()
        except Exception as e:
            raise PlanError(f"되돌리기용 사본을 만들지 못해 실행하지 않았습니다: {e}")

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
        # shortcut: 작업 전 필터가 걸려 있었다면 그 조건은 복원 안 됨, 필터 위 작업 지원 시 보강
        self.ws.Cells.Clear()  # Delete는 다른 시트 수식의 참조를 #REF!로 깨뜨려서 Clear를 씀
        self.wb.Worksheets(BACKUP).Cells.Copy(self.ws.Cells)
        for lo, name in zip(self.ws.ListObjects, self.tables):
            lo.Name = name  # 붙여넣으면 표 이름이 바뀌어 다른 수식의 표 참조가 깨지므로 원래 이름으로
        self.drop()

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
