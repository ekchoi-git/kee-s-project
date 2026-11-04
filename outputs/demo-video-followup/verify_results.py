"""Verify the filmed results against the source, without changing Excel."""
from pathlib import Path
import json
import math
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parent


def load(name):
    return json.loads((ROOT / name).read_text(encoding="utf-8"))


def close(actual, expected, label):
    if not isinstance(actual, (float, int)) or not math.isclose(actual, expected, rel_tol=0, abs_tol=0.01):
        raise AssertionError(f"{label}: expected {expected}, got {actual}")


def company(value):
    return "" if value in (None, "", "(빈 값)") else str(value)


def verify_table(state, expected):
    rows = state["values"]
    year_cols = {str(int(year)): col for col, year in enumerate(rows[0]) if isinstance(year, (int, float)) and 2000 <= year <= 2100}
    assert sorted(year_cols) == expected["years"], "Year headers differ"
    data_rows = {company(row[0]): index for index, row in enumerate(rows[1:], start=1) if company(row[0]) in expected["companies"]}
    assert sorted(data_rows) == expected["companies"], "Company rows differ"
    for cell in expected["cells"]:
        close(rows[data_rows[cell["company"]]][year_cols[cell["year"]]], cell["amount"], str(cell))
    return data_rows, year_cols


def verify(stage):
    expected = load("expected.json")
    state = load(f"excel-{stage}.json")
    data_rows, year_cols = verify_table(state, expected)
    checked = len(expected["cells"])
    if stage in ("totals", "formatted", "saved"):
        rows = state["values"]
        last_row, last_col = len(rows) - 1, len(rows[0]) - 1
        assert last_row not in data_rows.values() and last_col not in year_cols.values(), "Missing total row or column"
        for year, col in year_cols.items():
            close(rows[last_row][col], sum(c["amount"] for c in expected["cells"] if c["year"] == year), f"Annual total {year}")
        for name, row in data_rows.items():
            close(rows[row][last_col], sum(c["amount"] for c in expected["cells"] if c["company"] == name), f"Company total {name}")
        close(rows[last_row][last_col], expected["total"], "Grand total")
        checked += len(data_rows) + len(year_cols) + 1
    if stage in ("formatted", "saved"):
        previous = load("excel-totals.json")
        assert previous["sheet"] == state["sheet"], "Result sheet changed"
        assert previous["values"] == state["values"], "Formatting changed values"
        assert previous["formulas"] == state["formulas"], "Formatting changed formulas"
        assert previous["formats"][0] == state["formats"][0], "Year/header formats changed"
        for r in range(1, len(state["values"])):
            assert previous["formats"][r][0] == state["formats"][r][0], "Company label format changed"
            for c in range(1, len(state["values"][0])):
                assert state["formats"][r][c] == "#,##0,", f"Unexpected format at {r+1},{c+1}"
    if stage == "saved":
        assert state["saved"], "Excel workbook is not saved"
    report = {"passed": True, "stage": stage, "sheet": state["sheet"], "checkedCells": checked, "expectedTotal": expected["total"]}
    (ROOT / f"verification-{stage}.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))


if __name__ == "__main__":
    verify(sys.argv[1])
