"""pdf-split-unified の統合テスト（FR-12）.

参考コード2本（refs/PDF-Split-Rename, refs/PDF-Split-Rename-By-Cell）
の単体テストを移植し、共通コア・2モード・プレビューのテストを追加する。
実行: `python -m unittest -v`
"""

from __future__ import annotations

import os
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import call, patch

from openpyxl import Workbook
from pypdf import PdfReader, PdfWriter

import modes
import preview
import split_core


class FakeSheet:
    """openpyxlのシートを模す。既知の番地だけ値を返す."""

    def __init__(self, cells):
        self.cells = cells

    def __getitem__(self, address):
        return self.cells.get(address, SimpleNamespace(value=None, data_type="n"))


class FakeWorkbook:
    """load_workbook の戻り値を模すダブリング."""

    sheetnames = ["Sheet1"]

    def __init__(self, cells):
        self.sheet = FakeSheet(cells)
        self.closed = False

    def __getitem__(self, sheet_name):
        return self.sheet

    def close(self):
        self.closed = True


class RealDataMixin:
    """実データ（ページ数分の空白PDF・A1から並ぶExcel・空の出力先）を作る."""

    @staticmethod
    def make_data(root: Path, values=("あいう", "かきく", "さしす")):
        pdf_path = root / "input.pdf"
        writer = PdfWriter()
        for _ in values:
            writer.add_blank_page(width=72, height=72)
        with pdf_path.open("wb") as stream:
            writer.write(stream)

        excel_path = root / "names.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        for offset, value in enumerate(values):
            if value is not None:  # None はセルを作らない（空欄扱い）
                sheet[f"A{offset + 1}"] = value
        workbook.save(excel_path)
        workbook.close()

        output_dir = root / "output"
        output_dir.mkdir()
        return pdf_path, excel_path, output_dir


class ValidateFilenameBaseTests(unittest.TestCase):
    def test_ordinary_value_is_kept(self):
        self.assertEqual(split_core.validate_filename_base("ABC123", "A1"), "ABC123")

    def test_invalid_windows_characters_rejected(self):
        for bad in ("bad/name", 'bad:name', "bad*name", "bad?name", 'bad"name',
                    "bad<name", "bad>name", "bad|name", "bad\\name"):
            with self.subTest(bad=bad):
                with self.assertRaisesRegex(ValueError, "ファイル名に使えない文字"):
                    split_core.validate_filename_base(bad, "B2")

    def test_control_characters_rejected(self):
        with self.assertRaisesRegex(ValueError, "ファイル名に使えない文字"):
            split_core.validate_filename_base("bad\x07name", "B2")

    def test_trailing_space_or_dot_rejected(self):
        with self.assertRaisesRegex(ValueError, "末尾に空白またはピリオド"):
            split_core.validate_filename_base("name. ", "A1")
        with self.assertRaisesRegex(ValueError, "末尾に空白またはピリオド"):
            split_core.validate_filename_base("name.", "A1")

    def test_windows_reserved_names_rejected_case_insensitively(self):
        for reserved in ("CON", "con", "PRN", "aux", "NUL", "com1", "COM9", "lpt1"):
            with self.subTest(reserved=reserved):
                with self.assertRaisesRegex(ValueError, "予約名"):
                    split_core.validate_filename_base(reserved, "A1")

    def test_names_starting_like_reserved_are_allowed(self):
        self.assertEqual(split_core.validate_filename_base("CONSOLE1", "A1"), "CONSOLE1")
        self.assertEqual(split_core.validate_filename_base("COM10", "A1"), "COM10")
class ParseCellAddressTests(unittest.TestCase):
    def test_plain_and_absolute_forms(self):
        self.assertEqual(split_core.parse_cell_address("a1"), ("A", 1))
        self.assertEqual(split_core.parse_cell_address("$B$3"), ("B", 3))
        self.assertEqual(split_core.parse_cell_address(" AA123 "), ("AA", 123))

    def test_invalid_forms_use_japanese_label(self):
        for bad in ("1A", "A0", "A-1", "", "A1B", "ABBBC1"):
            with self.subTest(bad=bad):
                with self.assertRaisesRegex(ValueError, "開始セルを正しく入力してください"):
                    split_core.parse_cell_address(bad, "開始セル")

    def test_label_follows_mode_wording(self):
        with self.assertRaisesRegex(ValueError, "セル番地を正しく入力してください"):
            split_core.parse_cell_address("!!", "セル番地")


class ReadCellValuesTests(unittest.TestCase):
    """参考コード2本の数式キャッシュ検査を移植（load_workbook は2回読みの順も確認）."""

    def read(self, formula_cells, value_cells, addresses=("A1",), messages=None):
        formula_book = FakeWorkbook(formula_cells)
        value_book = FakeWorkbook(value_cells)
        with patch.object(split_core, "load_workbook", side_effect=[formula_book, value_book]) as loader:
            kwargs = {} if messages is None else {"messages": messages}
            values = split_core.read_cell_values(Path("data.xlsx"), "Sheet1", addresses, **kwargs)
        self.assertEqual(loader.call_args_list, [
            call(Path("data.xlsx"), read_only=True, data_only=False),
            call(Path("data.xlsx"), read_only=True, data_only=True),
        ])
        self.assertTrue(formula_book.closed)
        self.assertTrue(value_book.closed)
        return values

    def test_ordinary_value_cells(self):
        cells = {
            "A1": SimpleNamespace(value="あいう", data_type="s"),
            "A2": SimpleNamespace(value="かきく", data_type="s"),
        }
        self.assertEqual(self.read(cells, cells, ("A1", "A2")), {"A1": "あいう", "A2": "かきく"})

    def test_cached_formula_results_are_used_without_calculation(self):
        formulas = {
            "A1": SimpleNamespace(value='="あいう"', data_type="f"),
            "A2": SimpleNamespace(value='="かきく"', data_type="f"),
        }
        values = {
            "A1": SimpleNamespace(value="あいう", data_type="s"),
            "A2": SimpleNamespace(value="かきく", data_type="s"),
        }
        self.assertEqual(self.read(formulas, values, ("A1", "A2")), {"A1": "あいう", "A2": "かきく"})

    def test_missing_formula_cache_prompts_excel_recalculation(self):
        formulas = {"A1": SimpleNamespace(value="=B1", data_type="f")}
        values = {"A1": SimpleNamespace(value=None, data_type="n")}
        with self.assertRaisesRegex(ValueError, "Excelでブックを開き、数式を再計算してから保存"):
            self.read(formulas, values)

    def test_error_value_cell_rejected(self):
        formulas = {"A1": SimpleNamespace(value="=B1/0", data_type="f")}
        values = {"A1": SimpleNamespace(value="#DIV/0!", data_type="e")}
        with self.assertRaisesRegex(ValueError, "数式でエラーが発生しています"):
            self.read(formulas, values)

    def test_empty_cell_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "が空です"):
            self.read({}, {})

    def test_empty_cell_message_follows_style(self):
        with self.assertRaisesRegex(ValueError, "指定したセル（Sheet1!A1）が空です。ベースファイル名を入力してください。"):
            self.read({}, {}, messages=split_core.BASE_CELL_MESSAGES)

    def test_missing_sheet_is_rejected(self):
        cells = {"A1": SimpleNamespace(value="x", data_type="s")}
        formula_book = FakeWorkbook(cells)
        value_book = FakeWorkbook(cells)
        with patch.object(split_core, "load_workbook", side_effect=[formula_book, value_book]):
            with self.assertRaisesRegex(ValueError, "指定したシート「Other」がExcelブックにありません。"):
                split_core.read_cell_values(Path("data.xlsx"), "Other", ["A1"])

    def test_invalid_filename_value_is_rejected(self):
        cells = {"A1": SimpleNamespace(value="bad/name", data_type="s")}
        with self.assertRaisesRegex(ValueError, "ファイル名に使えない文字"):
            self.read(cells, cells)

    def test_unreadable_excel_raises_japanese_error(self):
        with patch.object(split_core, "load_workbook", side_effect=OSError("broken")):
            with self.assertRaisesRegex(ValueError, "Excelファイルを読み込めませんでした"):
                split_core.read_cell_values(Path("data.xlsx"), "Sheet1", ["A1"])
class PerCellModeTests(unittest.TestCase):
    """参考コード test_pdf_splitter_by_cell.py の読取テストを移植."""

    def build(self, formula_cells, value_cells, start="A1", page_count=2):
        formula_book = FakeWorkbook(formula_cells)
        value_book = FakeWorkbook(value_cells)
        with patch.object(split_core, "load_workbook", side_effect=[formula_book, value_book]) as loader:
            names = modes.build_filenames_per_cell(Path("data.xlsx"), "Sheet1", start, page_count)
        self.assertEqual(loader.call_args_list, [
            call(Path("data.xlsx"), read_only=True, data_only=False),
            call(Path("data.xlsx"), read_only=True, data_only=True),
        ])
        self.assertTrue(formula_book.closed)
        self.assertTrue(value_book.closed)
        return names

    def test_reads_downward_from_selected_start_cell_without_sequence(self):
        cells = {
            "B3": SimpleNamespace(value="あいう", data_type="s"),
            "B4": SimpleNamespace(value="かきく", data_type="s"),
        }
        self.assertEqual(self.build(cells, cells, start="B3"), ["あいう.pdf", "かきく.pdf"])

    def test_uses_cached_formula_results(self):
        formulas = {
            "A1": SimpleNamespace(value='="あいう"', data_type="f"),
            "A2": SimpleNamespace(value='="かきく"', data_type="f"),
        }
        values = {
            "A1": SimpleNamespace(value="あいう", data_type="s"),
            "A2": SimpleNamespace(value="かきく", data_type="s"),
        }
        self.assertEqual(self.build(formulas, values), ["あいう.pdf", "かきく.pdf"])

    def test_missing_formula_cache_prompts_excel_recalculation(self):
        formulas = {"A1": SimpleNamespace(value="=B1", data_type="f")}
        values = {"A1": SimpleNamespace(value=None, data_type="n")}
        with self.assertRaisesRegex(ValueError, "Excelでブックを開き、数式を再計算してから保存"):
            self.build(formulas, values, page_count=1)

    def test_empty_cell_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "が空です"):
            self.build({}, {}, page_count=1)

    def test_duplicate_names_are_rejected_case_insensitively(self):
        cells = {
            "A1": SimpleNamespace(value="Report", data_type="s"),
            "A2": SimpleNamespace(value="report", data_type="s"),
        }
        with self.assertRaisesRegex(ValueError, "同じファイル名"):
            self.build(cells, cells)

    def test_invalid_windows_filename_is_rejected(self):
        cells = {"A1": SimpleNamespace(value="bad/name", data_type="s")}
        with self.assertRaisesRegex(ValueError, "ファイル名に使えない文字"):
            self.build(cells, cells, page_count=1)

    def test_reading_past_excel_last_row_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "Excelの最終行を超えます"):
            modes.build_filenames_per_cell(Path("data.xlsx"), "Sheet1", "A1048576", 2)

    def test_page_count_below_one_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "PDFにページがありません"):
            modes.build_filenames_per_cell(Path("data.xlsx"), "Sheet1", "A1", 0)


class BaseSeqModeTests(unittest.TestCase):
    """参考コード test_pdf_splitter.py の数式キャッシュ検査を移植."""

    def build(self, formula_cell, cached_cell, page_count=2, cell="A1"):
        formula_book = FakeWorkbook({"A1": formula_cell})
        value_book = FakeWorkbook({"A1": cached_cell})
        with patch.object(split_core, "load_workbook", side_effect=[formula_book, value_book]) as loader:
            names = modes.build_filenames_base_seq(Path("book.xlsx"), "Sheet1", cell, page_count)
        self.assertEqual(loader.call_args_list, [
            call(Path("book.xlsx"), read_only=True, data_only=False),
            call(Path("book.xlsx"), read_only=True, data_only=True),
        ])
        self.assertTrue(formula_book.closed)
        self.assertTrue(value_book.closed)
        return names

    def test_ordinary_value_cell_keeps_using_its_value(self):
        names = self.build(
            SimpleNamespace(value="ordinary-name", data_type="s"),
            SimpleNamespace(value="ordinary-name", data_type="s"),
            page_count=3,
        )
        self.assertEqual(names, [
            "ordinary-name_001.pdf", "ordinary-name_002.pdf", "ordinary-name_003.pdf",
        ])

    def test_formula_cell_uses_saved_cached_value(self):
        names = self.build(
            SimpleNamespace(value='=A2&"-name"', data_type="f"),
            SimpleNamespace(value="cached-name", data_type="s"),
            page_count=1,
        )
        self.assertEqual(names, ["cached-name_001.pdf"])

    def test_formula_without_cached_value_explains_recalculation(self):
        with self.assertRaisesRegex(ValueError, "Excelでブックを開き、数式を再計算してから保存"):
            self.build(SimpleNamespace(value="=A2", data_type="f"), SimpleNamespace(value=None), page_count=1)

    def test_sequence_digit_grows_beyond_999(self):
        names = self.build(
            SimpleNamespace(value="x", data_type="s"),
            SimpleNamespace(value="x", data_type="s"),
            page_count=1001,
        )
        self.assertEqual(names[999], "x_1000.pdf")
        self.assertEqual(names[1000], "x_1001.pdf")

    def test_address_is_normalized_before_reading(self):
        names = self.build(
            SimpleNamespace(value="x", data_type="s"),
            SimpleNamespace(value="x", data_type="s"),
            page_count=1, cell="$a$1",
        )
        self.assertEqual(names, ["x_001.pdf"])

    def test_empty_cell_message_keeps_base_mode_wording(self):
        with self.assertRaisesRegex(ValueError, "指定したセル（Sheet1!A1）が空です。ベースファイル名を入力してください。"):
            self.build(
                SimpleNamespace(value=None, data_type="n"),
                SimpleNamespace(value=None, data_type="n"),
                page_count=1,
            )

    def test_missing_sheet_keeps_base_mode_wording(self):
        formula_book = FakeWorkbook({})
        value_book = FakeWorkbook({})
        with patch.object(split_core, "load_workbook", side_effect=[formula_book, value_book]):
            with self.assertRaisesRegex(ValueError, "指定したシート「Other」がExcelにありません。"):
                modes.build_filenames_base_seq(Path("book.xlsx"), "Other", "A1", 1)

    def test_page_count_below_one_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "PDFにページがありません"):
            modes.build_filenames_base_seq(Path("book.xlsx"), "Sheet1", "A1", 0)

    def test_dispatcher_unknown_mode(self):
        with self.assertRaisesRegex(ValueError, "未対応のモードです"):
            modes.build_filenames("no-such-mode", Path("book.xlsx"), "Sheet1", "A1", 1)
class LoadPdfTests(RealDataMixin, unittest.TestCase):
    def make_pdf(self, root: Path) -> Path:
        pdf_path, _, _ = self.make_data(root)
        return pdf_path

    def test_encrypted_pdf_requires_unlocked_file(self):
        with TemporaryDirectory() as temporary:
            path = Path(temporary) / "secret.pdf"
            writer = PdfWriter()
            writer.add_blank_page(width=72, height=72)
            writer.encrypt("hunter2")
            with path.open("wb") as stream:
                writer.write(stream)
            with self.assertRaisesRegex(ValueError, "パスワードで保護されています"):
                split_core.load_pdf(path)

    def test_broken_pdf_raises_japanese_error(self):
        with TemporaryDirectory() as temporary:
            path = Path(temporary) / "broken.pdf"
            path.write_bytes(b"this is not a pdf")
            with self.assertRaisesRegex(ValueError, "PDFを読み込めませんでした"):
                split_core.load_pdf(path)

    def test_zero_page_pdf_is_rejected(self):
        with TemporaryDirectory() as temporary:
            path = Path(temporary) / "empty.pdf"
            with path.open("wb") as stream:
                PdfWriter().write(stream)
            with self.assertRaisesRegex(ValueError, "PDFにページがありません"):
                split_core.load_pdf(path)

    def test_input_path_checks(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root)
            missing = root / "nope.pdf"
            notes = root / "notes.txt"
            notes.write_text("plain text", encoding="utf-8")
            cases = [
                (missing, excel_path, output_dir, "指定したPDFファイルが見つかりません"),
                (pdf_path, root / "nope.xlsx", output_dir, "指定したExcelファイルが見つかりません"),
                (excel_path, excel_path, output_dir, "PDFファイルを選択してください"),
                (pdf_path, notes, output_dir, "Excelファイル（.xlsx または .xlsm）を選択してください"),
                (pdf_path, excel_path, root / "no-dir", "出力先フォルダーが見つかりません"),
            ]
            for pdf, excel, out, message in cases:
                with self.subTest(pdf=pdf, excel=excel, out=out):
                    with self.assertRaisesRegex(ValueError, message):
                        split_core.check_input_paths(pdf, excel, out)


class WritePagesTests(RealDataMixin, unittest.TestCase):
    def make_reader_and_names(self, root: Path, names: list[str]):
        pdf_path, _, _ = self.make_data(root)
        reader = split_core.load_pdf(pdf_path)
        return reader, names

    def test_existing_output_blocks_without_overwrite(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            reader, names = self.make_reader_and_names(root, ["あいう.pdf", "かきく.pdf", "さしす.pdf"])
            output_dir = root / "output"
            guard = output_dir / "かきく.pdf"
            guard.write_bytes(b"ORIGINAL")
            with self.assertRaisesRegex(FileExistsError, "同名の出力ファイルが既にあります。上書きは行いません"):
                split_core.write_pages(reader, names, output_dir)
            self.assertEqual(guard.read_bytes(), b"ORIGINAL")
            self.assertFalse((output_dir / "あいう.pdf").exists())
            self.assertFalse((output_dir / "さしす.pdf").exists())
            self.assertEqual([p.name for p in output_dir.iterdir()], ["かきく.pdf"])  # 一時フォルダーも残さない

    def test_existing_list_is_capped_with_remainder_count(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            names = [f"file{i}.pdf" for i in range(1, 11)]
            pdf_path = root / "input.pdf"
            writer = PdfWriter()
            for _ in names:
                writer.add_blank_page(width=72, height=72)
            with pdf_path.open("wb") as stream:
                writer.write(stream)
            reader = split_core.load_pdf(pdf_path)
            output_dir = root / "output"
            output_dir.mkdir()
            for name in names[:9]:
                (output_dir / name).write_bytes(b"x")
            with self.assertRaisesRegex(FileExistsError, "ほか 1 件"):
                split_core.write_pages(reader, names, output_dir)

    def test_page_name_count_mismatch_is_rejected(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            reader, names = self.make_reader_and_names(root, ["only.pdf"])
            output_dir = root / "output"
            with self.assertRaisesRegex(ValueError, "ページ数とファイル名リストの数が一致しません"):
                split_core.write_pages(reader, names, output_dir)

    def test_staging_failure_leaves_nothing_behind(self):
        class FailingWriter:
            def __init__(self):
                pass

            def add_page(self, page):
                pass

            def write(self, stream):
                raise OSError("simulated disk failure")

        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            reader, names = self.make_reader_and_names(root, ["a.pdf", "b.pdf", "c.pdf"])
            output_dir = root / "output"
            with patch.object(split_core, "PdfWriter", FailingWriter):
                with self.assertRaisesRegex(ValueError, "分割ファイルを保存できませんでした"):
                    split_core.write_pages(reader, names, output_dir)
            self.assertEqual(list(output_dir.iterdir()), [])

    def test_move_failure_rolls_back_created_outputs(self):
        real_rename = os.rename
        state = {"calls": 0}

        def flaky_rename(src, dst):
            state["calls"] += 1
            if state["calls"] == 1:
                real_rename(src, dst)
            else:
                raise OSError("simulated move failure")

        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            reader, names = self.make_reader_and_names(root, ["a.pdf", "b.pdf", "c.pdf"])
            output_dir = root / "output"
            with patch("os.rename", side_effect=flaky_rename):
                with self.assertRaisesRegex(ValueError, "分割ファイルを保存できませんでした"):
                    split_core.write_pages(reader, names, output_dir)
            self.assertGreaterEqual(state["calls"], 2)
            self.assertEqual(list(output_dir.iterdir()), [])  # 移動済み1件も消えて片付いている
class SplitEndToEndTests(RealDataMixin, unittest.TestCase):
    def check_one_page_each(self, results):
        self.assertEqual([len(PdfReader(str(path)).pages) for path in results], [1, 1, 1])

    def test_base_seq_split_end_to_end(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root)
            before = (pdf_path.read_bytes(), excel_path.read_bytes())
            results = split_core.split_pdf(
                pdf_path, excel_path, output_dir, "Sheet", "A1", modes.build_filenames_base_seq,
            )
            self.assertEqual([p.name for p in results], ["あいう_001.pdf", "あいう_002.pdf", "あいう_003.pdf"])
            self.check_one_page_each(results)
            self.assertEqual(sorted(p.name for p in output_dir.iterdir()), [
                "あいう_001.pdf", "あいう_002.pdf", "あいう_003.pdf",
            ])
            self.assertEqual((pdf_path.read_bytes(), excel_path.read_bytes()), before)  # 元ファイルは無変更

    def test_per_cell_split_end_to_end(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root)
            results = split_core.split_pdf(
                pdf_path, excel_path, output_dir, "Sheet", "A1", modes.build_filenames_per_cell,
            )
            self.assertEqual([p.name for p in results], ["あいう.pdf", "かきく.pdf", "さしす.pdf"])
            self.check_one_page_each(results)

    def test_dispatcher_mode_selection_same_names(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root)
            results = split_core.split_pdf(
                pdf_path, excel_path, output_dir, "Sheet", "A1",
                lambda excel, sheet, cell, page_count: modes.build_filenames(
                    modes.PER_CELL, excel, sheet, cell, page_count,
                ),
            )
            self.assertEqual([p.name for p in results], ["あいう.pdf", "かきく.pdf", "さしす.pdf"])

    def test_blank_cell_stops_before_any_output(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root, values=("あいう", None, "さしす"))
            with self.assertRaisesRegex(ValueError, "が空です"):
                split_core.split_pdf(
                    pdf_path, excel_path, output_dir, "Sheet", "A1", modes.build_filenames_per_cell,
                )
            self.assertEqual(list(output_dir.iterdir()), [])

    def test_duplicate_names_stop_before_any_output(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root, values=("Report", "report", "third"))
            with self.assertRaisesRegex(ValueError, "同じファイル名"):
                split_core.split_pdf(
                    pdf_path, excel_path, output_dir, "Sheet", "A1", modes.build_filenames_per_cell,
                )
            self.assertEqual(list(output_dir.iterdir()), [])

    def test_mode_labels_match_specification(self):
        self.assertEqual(modes.MODE_LABELS[modes.BASE_SEQ], "セル値＋連番")
        self.assertEqual(modes.MODE_LABELS[modes.PER_CELL], "セル別（連番なし）")


class PreviewTests(RealDataMixin, unittest.TestCase):
    def test_existing_files_marked_and_nothing_written(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root)
            guard = output_dir / "かきく.pdf"
            guard.write_bytes(b"KEEP")

            rows = preview.build_preview(excel_path, "Sheet", "A1", modes.PER_CELL, output_dir, pdf_path)

            self.assertEqual([row.page_no for row in rows], [1, 2, 3])
            self.assertEqual([row.filename for row in rows], ["あいう.pdf", "かきく.pdf", "さしす.pdf"])
            self.assertEqual([row.status for row in rows], ["", "[既存]", ""])
            self.assertEqual(guard.read_bytes(), b"KEEP")  # 既存ファイルは無変更
            self.assertEqual(sorted(p.name for p in output_dir.iterdir()), ["かきく.pdf"])  # 書き出しなし

    def test_duplicate_rows_marked_without_error(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root, values=("Report", "report", "third"))
            rows = preview.build_preview(excel_path, "Sheet", "A1", modes.PER_CELL, output_dir, pdf_path)
            self.assertEqual([row.status for row in rows], ["[重複]", "[重複]", ""])
            self.assertEqual(list(output_dir.iterdir()), [])

    def test_base_seq_preview_names_and_no_duplicate_marks(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root)
            rows = preview.build_preview(excel_path, "Sheet", "A1", modes.BASE_SEQ, output_dir, pdf_path)
            self.assertEqual([row.filename for row in rows], [
                "あいう_001.pdf", "あいう_002.pdf", "あいう_003.pdf",
            ])
            self.assertTrue(all(not row.duplicate and not row.exists for row in rows))
            self.assertEqual(list(output_dir.iterdir()), [])

    def test_invalid_cells_propagate_japanese_errors(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root, values=("bad/name", "かきく", "さしす"))
            with self.assertRaisesRegex(ValueError, "ファイル名に使えない文字"):
                preview.build_preview(excel_path, "Sheet", "A1", modes.PER_CELL, output_dir, pdf_path)

    def test_truncate_display_keeps_all_data_and_announces_remainder(self):
        rows = [preview.PreviewRow(page_no, f"f{page_no}.pdf", False, False) for page_no in range(1, 151)]
        shown, extra = preview.truncate_display(rows)
        self.assertEqual(len(shown), preview.MAX_ANNOUNCED)
        self.assertEqual(extra, 50)
        self.assertEqual(rows[149].display_cells(), ("150", "f150.pdf", ""))
        short, extra2 = preview.truncate_display(rows[:preview.MAX_ANNOUNCED])
        self.assertEqual(extra2, 0)

    def test_status_joins_both_marks(self):
        row = preview.PreviewRow(1, "a.pdf", exists=True, duplicate=True)
        self.assertEqual(row.status, "[既存] [重複]")


class GuiWordingTests(unittest.TestCase):
    """GUI は表示のみ。ウィンドウを開かずに、文言とインポート確認だけ行う（FR-15）."""

    def test_gui_module_imports_without_opening_window(self):
        import pdf_split_unified

        self.assertTrue(hasattr(pdf_split_unified, "UnifiedPdfSplitApp"))
        self.assertTrue(hasattr(pdf_split_unified, "main"))
        self.assertEqual(pdf_split_unified.PREVIEW_VISIBLE_ROWS, 12)
        self.assertEqual(pdf_split_unified.STATE_STALE, "入力変更あり: プレビューを更新してください")


if __name__ == "__main__":
    unittest.main()




