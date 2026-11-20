"""独立検証スクリプト（FR-13）。

プロジェクト内の一時フォルダー（`.pdf_split_verify_*`）に実データ
（3ページ空白PDF・A1〜A3 に あいう/かきく/さしす のExcel）を作り、
両モードで分割を実行して機械的に合否を出す。表示用数値は手で書かない。

実行: `.venv\\Scripts\\python.exe tools\\verify_independent.py`
出力: PASS/FAIL を1行ずつ、最終行に `TOTAL: <成功>/<総数>`。
終了コードは全件成功時 0、そうでない時 0 以外。
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook, load_workbook
from pypdf import PdfReader, PdfWriter

import modes
import preview
import split_core

RESULTS: list[tuple[str, bool]] = []


def check(label: str, ok: bool) -> None:
    RESULTS.append((label, bool(ok)))


def make_pdf(path: Path, pages: int) -> None:
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=72, height=72)
    with path.open("wb") as stream:
        writer.write(stream)


def make_excel(path: Path, cells: dict[str, str]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    for address, value in cells.items():
        sheet[address] = value
    workbook.save(path)
    workbook.close()


def page_count(path: Path) -> int:
    return len(PdfReader(str(path)).pages)


def top_level_pdf_names(output_dir: Path) -> list[str]:
    return sorted(p.name for p in output_dir.glob("*.pdf"))


def all_dir_names(output_dir: Path) -> list[str]:
    return sorted(p.name for p in output_dir.iterdir())
def verify(root: Path) -> None:
    # 1. 実データの生成
    pdf_path = root / "input.pdf"
    excel_path = root / "names.xlsx"
    make_pdf(pdf_path, 3)
    make_excel(excel_path, {"A1": "あいう", "A2": "かきく", "A3": "さしす"})
    book = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        sheet = book["Sheet"]
        excel_values = [sheet[f"A{row}"].value for row in (1, 2, 3)]
        sheet_names = book.sheetnames
    finally:
        book.close()
    data_ok = page_count(pdf_path) == 3 and sheet_names == ["Sheet"] and excel_values == ["あいう", "かきく", "さしす"]
    check("1. 3ページの空白PDFとA1〜A3にあいう/かきく/さしすを持つExcelを生成", data_ok)

    # 2. base_seq（A1）
    out_base = root / "out_base_seq"
    out_base.mkdir()
    split_core.split_pdf(pdf_path, excel_path, out_base, "Sheet", "A1", modes.build_filenames_base_seq)
    base_names = top_level_pdf_names(out_base)
    check(
        "2. base_seq（A1）→ あいう_001.pdf / あいう_002.pdf / あいう_003.pdf が生成される",
        base_names == ["あいう_001.pdf", "あいう_002.pdf", "あいう_003.pdf"],
    )
    check(
        "3. base_seqの各PDFはちょうど1ページ",
        len(base_names) == 3 and all(page_count(out_base / name) == 1 for name in base_names),
    )

    # 3. per_cell（A1）
    out_cell = root / "out_per_cell"
    out_cell.mkdir()
    split_core.split_pdf(pdf_path, excel_path, out_cell, "Sheet", "A1", modes.build_filenames_per_cell)
    cell_names = top_level_pdf_names(out_cell)
    check(
        "4. per_cell（A1）→ あいう.pdf / かきく.pdf / さしす.pdf が生成される",
        cell_names == sorted(["あいう.pdf", "かきく.pdf", "さしす.pdf"]),
    )
    check(
        "5. per_cellの各PDFはちょうど1ページ",
        len(cell_names) == 3 and all(page_count(out_cell / name) == 1 for name in cell_names),
    )

    # 4. 同じ出力先で再実行 → FileExistsError で中止・既存の中身は無変更
    before_bytes = {name: (out_cell / name).read_bytes() for name in cell_names}
    guard_base = False
    try:
        split_core.split_pdf(pdf_path, excel_path, out_base, "Sheet", "A1", modes.build_filenames_base_seq)
    except FileExistsError:
        guard_base = True
    check("6. base_seqの再実行はFileExistsErrorで中止（上書きなし）", guard_base)
    guard_cell = False
    try:
        split_core.split_pdf(pdf_path, excel_path, out_cell, "Sheet", "A1", modes.build_filenames_per_cell)
    except FileExistsError:
        guard_cell = True
    check("7. per_cellの再実行はFileExistsErrorで中止（上書きなし）", guard_cell)
    after_bytes = {name: (out_cell / name).read_bytes() for name in top_level_pdf_names(out_cell)}
    tidy_ok = (
        after_bytes == before_bytes
        and all_dir_names(out_base) == [f"あいう_{n:03d}.pdf" for n in (1, 2, 3)]
        and all_dir_names(out_cell) == cell_names  # 一時フォルダー（.pdf_split_*）も残さない
    )
    check("8. 再実行後も既存ファイルの中身が変わっておらず、後片付けされている", tidy_ok)
    # 5. 空欄セル（タブ2でA2を空）→ 実行前に中止し、出力は0件
    excel_blank = root / "blank.xlsx"
    make_excel(excel_blank, {"A1": "p1", "A3": "p3"})  # A2 は作らない（空欄）
    out_blank = root / "out_blank"
    out_blank.mkdir()
    blank_abort = False
    try:
        split_core.split_pdf(pdf_path, excel_blank, out_blank, "Sheet", "A1", modes.build_filenames_per_cell)
    except ValueError as exc:
        blank_abort = "が空です" in str(exc)
    check("9. 空欄セル（A2）は実行前のエラーで中止", blank_abort)
    check("10. 空欄セルのとき出力ファイルは1つも作られていない", all_dir_names(out_blank) == [])

    # 6. 重複名（Report / report）→ 重複エラーで中止
    excel_dup = root / "dup.xlsx"
    make_excel(excel_dup, {"A1": "Report", "A2": "report", "A3": "solo"})
    out_dup = root / "out_dup"
    out_dup.mkdir()
    dup_abort = False
    try:
        split_core.split_pdf(pdf_path, excel_dup, out_dup, "Sheet", "A1", modes.build_filenames_per_cell)
    except ValueError as exc:
        dup_abort = "同じファイル名" in str(exc)
    check("11. 重複名（Report / report）は重複エラーで中止", dup_abort)
    check("12. 重複名のとき出力ファイルは1つも作られていない", all_dir_names(out_dup) == [])

    # 7. プレビュー: 4の状況で [既存]、6の状況で [重複]、書き出しなし
    dup_rows = preview.build_preview(excel_dup, "Sheet", "A1", modes.PER_CELL, out_dup, pdf_path)
    dup_marks = [row.status for row in dup_rows] == ["[重複]", "[重複]", ""]
    check(
        "13. プレビューは重複状況で [重複] を付け、書き出しが発生していない",
        dup_marks and all_dir_names(out_dup) == [],
    )
    exist_rows = preview.build_preview(excel_path, "Sheet", "A1", modes.PER_CELL, out_cell, pdf_path)
    exist_marks = len(exist_rows) == 3 and all("[既存]" in row.status for row in exist_rows)
    preview_untouched = {name: (out_cell / name).read_bytes() for name in top_level_pdf_names(out_cell)} == before_bytes
    check(
        "14. プレビューは既存ファイル状況で [既存] を付け、既存の中身と一覧は無変更",
        exist_marks and preview_untouched and all_dir_names(out_cell) == cell_names,
    )


def main() -> int:
    root = Path(tempfile.mkdtemp(prefix=".pdf_split_verify_", dir=str(PROJECT_ROOT)))
    try:
        verify(root)
    except Exception as exc:  # 想定外例外も合否として計上する
        check(f"検証スクリプトが想定外のエラーで終了した: {exc}", False)
    finally:
        shutil.rmtree(root, ignore_errors=True)

    succeeded = sum(1 for _label, ok in RESULTS if ok)
    for label, ok in RESULTS:
        print(("PASS: " if ok else "FAIL: ") + label)
    print(f"TOTAL: {succeeded}/{len(RESULTS)}")
    return 0 if succeeded == len(RESULTS) else 1


if __name__ == "__main__":
    sys.exit(main())


