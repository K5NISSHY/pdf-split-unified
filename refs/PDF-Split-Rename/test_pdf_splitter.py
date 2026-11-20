import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import call, patch

import pdf_splitter


class FakeSheet:
    def __init__(self, cell):
        self.cell = cell

    def __getitem__(self, address):
        return self.cell


class FakeWorkbook:
    sheetnames = ["Sheet1"]

    def __init__(self, value, data_type="n"):
        self.cell = SimpleNamespace(value=value, data_type=data_type)
        self.closed = False

    def __getitem__(self, sheet_name):
        return FakeSheet(self.cell)

    def close(self):
        self.closed = True


class ReadBaseNameFormulaTests(unittest.TestCase):
    def read_with(self, formula_cell, cached_cell):
        formula_book = FakeWorkbook(formula_cell.value, formula_cell.data_type)
        value_book = FakeWorkbook(cached_cell.value, cached_cell.data_type)
        with patch.object(pdf_splitter, "load_workbook", side_effect=[formula_book, value_book]) as loader:
            result = pdf_splitter.read_base_name(Path("book.xlsx"), "Sheet1", "A1")
        self.assertEqual(loader.call_args_list, [
            call(Path("book.xlsx"), read_only=True, data_only=False),
            call(Path("book.xlsx"), read_only=True, data_only=True),
        ])
        self.assertTrue(formula_book.closed)
        self.assertTrue(value_book.closed)
        return result

    def test_ordinary_value_cell_keeps_using_its_value(self):
        self.assertEqual(self.read_with(
            SimpleNamespace(value="ordinary-name", data_type="s"),
            SimpleNamespace(value="ordinary-name", data_type="s"),
        ), "ordinary-name")

    def test_formula_cell_uses_saved_cached_value(self):
        self.assertEqual(self.read_with(
            SimpleNamespace(value="=A2&\"-name\"", data_type="f"),
            SimpleNamespace(value="cached-name", data_type="s"),
        ), "cached-name")

    def test_formula_without_cached_value_explains_recalculation(self):
        formula_book = FakeWorkbook("=A2", "f")
        value_book = FakeWorkbook(None)
        with patch.object(pdf_splitter, "load_workbook", side_effect=[formula_book, value_book]):
            with self.assertRaisesRegex(ValueError, "Excelでブックを開き、数式を再計算してから保存"):
                pdf_splitter.read_base_name(Path("book.xlsx"), "Sheet1", "A1")
        self.assertTrue(formula_book.closed)
        self.assertTrue(value_book.closed)


if __name__ == "__main__":
    unittest.main()
