"""상위 N개 Excel 필터 계획과 COM 호출 인자를 확인한다."""
import sys
import types
import unittest

try:
    import pywintypes  # noqa: F401
except ModuleNotFoundError:
    sys.modules["pywintypes"] = types.SimpleNamespace(com_error=Exception)

from app.excel_ops import XL_TOP10_ITEMS, apply_filter, top_items_count, validate


class FakeRange:
    def __init__(self):
        self.call = None

    def AutoFilter(self, **kwargs):
        self.call = kwargs


class TopItemsFilterTest(unittest.TestCase):
    def test_top_five_uses_native_excel_operator(self):
        region = FakeRange()
        text = apply_filter(region, 4, "top_items", "5")
        self.assertEqual(text, "상위 5개")
        self.assertEqual(region.call, {"Field": 4, "Criteria1": "5", "Operator": XL_TOP10_ITEMS})

    def test_count_validation(self):
        self.assertEqual(top_items_count("1"), 1)
        self.assertEqual(top_items_count("500"), 500)
        self.assertIsNone(top_items_count("0"))
        self.assertIsNone(top_items_count("5건"))
        self.assertIsNone(top_items_count("501"))

    def test_invalid_plan_is_rejected_before_excel(self):
        plan = {
            "filters": [{"column": "접수건수", "op": "top_items", "value": "다섯"}],
            "sorts": [], "pivot": None, "classify": None, "writes": [], "to_number": [],
            "edit_source": False, "formats": [], "highlights": [], "action": "change",
        }
        self.assertTrue(any("1~500" in message for message in validate(plan, ["접수건수"])))


if __name__ == "__main__":
    unittest.main()
