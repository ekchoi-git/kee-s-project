"""Excel 없이 도는 계획 검사 테스트: python tests/test_validate.py"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from app.excel_ops import validate, errors_in, col_letter, ERR_BASE

H = ["INVOICE_NUM", "INVOICE_DATE", "INVOICE_AMOUNT", "VENDOR_NAME"]

def plan(**kw):
    p = {"action": "change", "summary": "", "question": None, "overwrite": False,
         "filters": [], "sorts": [], "writes": [], "pivot": None}
    p.update(kw)
    return p

ok = plan(sorts=[{"column": "INVOICE_AMOUNT", "descending": True}],
          writes=[{"address": "J2:J501", "content": '=SUMIFS($C$2:$C$501,$D$2:$D$501,"A상사")'},
                  {"address": "K1", "content": "=UNIQUE($D$2:$D$501)"},
                  {"address": "L1", "content": '=TEXTJOIN("|",TRUE,K1#)'}])
assert validate(ok, H) == [], validate(ok, H)
assert validate(plan(sorts=[{"column": "금액", "descending": True}]), H) == ["'금액' 열이 표에 없습니다."]
assert validate(plan(pivot={"rows": ["VENDOR_NAME"], "columns": [], "values": []}), H) == ["피벗에 집계할 값이 없습니다."]
assert validate(plan(), H) == ["실행할 작업이 없습니다."]
assert validate(plan(action="ask"), H) == []
for bad in ["Sheet2!A1", "A0", "다른곳", "[x.xlsx]A1"]:
    assert validate(plan(writes=[{"address": bad, "content": "1"}]), H), bad
for f in ['=WEBSERVICE("http://x")', '=HYPERLINK("a")', r"='C:\a\[b.xlsx]S'!A1", "=[b.xlsx]S!A1",
          '=cmd|"/c calc"!A0', "=RTD(1)", "=CUBEVALUE(1)", '=IMAGE("u")', r"=SUM('\\srv\x'!A1)"]:
    assert validate(plan(writes=[{"address": "J1", "content": f}]), H), f
assert errors_in(((1.0, "x", True, ERR_BASE + 2042),)) == {"#N/A": 1}
assert errors_in(ERR_BASE + 2045) == {"#SPILL!": 1}
assert errors_in(5.0) == {}
assert [col_letter(n) for n in (1, 26, 27, 702, 703)] == ["A", "Z", "AA", "ZZ", "AAA"]
print("계획 검사 테스트 모두 통과")
