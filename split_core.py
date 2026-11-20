"""pdf-split-unified 共通コア.

Excel読み取り・ファイル名検証・PDF読み込み・一時書き出し/ロールバック・
既存ファイル検査をこの1箇所にまとめ、2モード（modes.py）は
「出力ファイル名リストを作る関数」だけを差し替えて共用する。
エラー文言は参考コード（refs/PDF-Split-Rename, refs/PDF-Split-Rename-By-Cell）
と同じ意味・同じ表現を維持する。
"""

from __future__ import annotations

import os
import re
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Sequence

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
from pypdf import PdfReader, PdfWriter

INVALID_FILENAME_CHARS = set('<>:"/\\|?*')
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
EXCEL_MAX_ROW = 1_048_576


@dataclass(frozen=True)
class CellMessageStyle:
    """セル読み取りエラーの文言スタイル（参考コード2本の文言をそれぞれ継承）."""

    cell_label: str  # セル位置の表記。書式は {sheet} と {address}
    empty_hint: str  # 空欄セルのときの後文
    sheet_missing: str  # シートが無いときの文言。書式は {sheet_name}


# 「セル値＋連番」モード（参考: pdf_splitter.py）の文言
BASE_CELL_MESSAGES = CellMessageStyle(
    cell_label="指定したセル（{sheet}!{address}）",
    empty_hint="ベースファイル名を入力してください。",
    sheet_missing="指定したシート「{sheet_name}」がExcelにありません。",
)

# 「セル別（連番なし）」モード（参考: pdf_splitter_by_cell.py）の文言
COLUMN_CELL_MESSAGES = CellMessageStyle(
    cell_label="セル {sheet}!{address} ",
    empty_hint="ページごとにファイル名を入力してください。",
    sheet_missing="指定したシート「{sheet_name}」がExcelブックにありません。",
)


def validate_filename_base(name: str, cell_address: str) -> str:
    """ファイル名ベースとして不正なら ValueError（セル位置入りの日本語）."""
    if any(ch in INVALID_FILENAME_CHARS for ch in name) or any(ord(ch) < 32 for ch in name):
        raise ValueError(
            f"セル {cell_address} の値「{name}」には、ファイル名に使えない文字（\\ / : * ? \" < > | など）が含まれています。"
        )
    if name.endswith((" ", ".")):
        raise ValueError(f"セル {cell_address} の値の末尾に空白またはピリオドは使えません。")
    if name.split(".", 1)[0].upper() in WINDOWS_RESERVED_NAMES:
        raise ValueError(f"セル {cell_address} の値「{name}」はWindowsで使用できない予約名です。")
    return name
def parse_cell_address(cell: str, label: str = "セル") -> tuple[str, int]:
    """`A1` / `$B$3` 形式を (列字母, 行番号) に分解。不正形式は ValueError."""
    address = cell.strip().upper()
    if not re.fullmatch(r"\$?[A-Z]{1,3}\$?[1-9][0-9]*", address):
        raise ValueError(f"{label}を正しく入力してください（例: A1、B3）。")
    address = address.replace("$", "")
    try:
        column, row = coordinate_from_string(address)
        column_index_from_string(column)
    except Exception as exc:
        raise ValueError(f"{label}を正しく入力してください（例: A1、B3）。") from exc
    return column, row


def read_cell_values(
    excel_path: Path,
    sheet_name: str,
    addresses: Sequence[str],
    messages: CellMessageStyle = COLUMN_CELL_MESSAGES,
) -> dict[str, str]:
    """指定セルの値を検証済み文字列として返す ({番地: 値}).

    openpyxlは数式を計算しないため、数式本文 (data_only=False) と
    Excelが保存した計算結果 (data_only=True) を別々に読み比べる。
    数式の計算結果が無い・エラー値・空欄は ValueError。
    """
    requested = [address.strip().upper() for address in addresses]
    formula_workbook = None
    value_workbook = None
    try:
        # openpyxl は数式を計算しない。数式本文とExcelの保存済み結果を分けて読む。
        formula_workbook = load_workbook(excel_path, read_only=True, data_only=False)
        value_workbook = load_workbook(excel_path, read_only=True, data_only=True)
        if sheet_name not in formula_workbook.sheetnames:
            raise ValueError(messages.sheet_missing.format(sheet_name=sheet_name))

        formula_sheet = formula_workbook[sheet_name]
        value_sheet = value_workbook[sheet_name]
        values: dict[str, str] = {}
        for address in requested:
            label = messages.cell_label.format(sheet=sheet_name, address=address)
            formula_cell = formula_sheet[address]
            value_cell = value_sheet[address]
            value = value_cell.value

            if formula_cell.data_type == "f" and (value is None or not str(value).strip()):
                raise ValueError(
                    f"{label}は数式セルですが、保存済みの計算結果がありません。\n"
                    "Excelでブックを開き、数式を再計算してから保存し、もう一度お試しください。"
                )
            if value_cell.data_type == "e":
                raise ValueError(f"{label}の数式でエラーが発生しています: {value}")
            if value is None or not str(value).strip():
                raise ValueError(f"{label}が空です。{messages.empty_hint}")

            values[address] = validate_filename_base(str(value).strip(), address)
        return values
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(f"Excelファイルを読み込めませんでした。\n{exc}") from exc
    finally:
        if formula_workbook is not None:
            formula_workbook.close()
        if value_workbook is not None:
            value_workbook.close()
def load_pdf(pdf_path: Path) -> PdfReader:
    """暗号化・ページ数を確認済みの PdfReader を返す（読み取り専用）."""
    try:
        reader = PdfReader(str(pdf_path), strict=True)
        if reader.is_encrypted:
            try:
                unlocked = reader.decrypt("")
            except Exception:
                unlocked = 0
            if not unlocked:
                raise ValueError("このPDFはパスワードで保護されています。解除したPDFを選択してください。")
        page_count = len(reader.pages)
        if page_count == 0:
            raise ValueError("PDFにページがありません。")
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(f"PDFを読み込めませんでした。破損または未対応の可能性があります。\n{exc}") from exc
    return reader


def check_input_paths(pdf_path: Path, excel_path: Path, output_dir: Path) -> None:
    """PDF・Excel・出力先フォルダーの存在と拡張子を確認する."""
    if not pdf_path.is_file():
        raise ValueError("指定したPDFファイルが見つかりません。")
    if not excel_path.is_file():
        raise ValueError("指定したExcelファイルが見つかりません。")
    if pdf_path.suffix.lower() != ".pdf":
        raise ValueError("PDFファイルを選択してください。")
    if excel_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Excelファイル（.xlsx または .xlsm）を選択してください。")
    if not output_dir.is_dir():
        raise ValueError("出力先フォルダーが見つかりません。")


def write_pages(reader: PdfReader, filenames: Sequence[str], output_dir: Path) -> list[Path]:
    """全ページを一時フォルダーへ書き出してから出力先へ移動する.

    出力予定ファイルが1つでも既存なら FileExistsError で開始前に中止。
    失敗した場合は作成済みの出力を削除して後片付けする。
    元のPDF（reader）は一切変更しない。戻り値は出力Pathのリスト。
    """
    pages = list(reader.pages)
    if len(pages) != len(filenames):
        raise ValueError("ページ数とファイル名リストの数が一致しません。")
    destinations = [output_dir / filename for filename in filenames]
    existing = [path.name for path in destinations if path.exists()]
    if existing:
        preview = "\n".join(existing[:8])
        extra = f"\nほか {len(existing) - 8} 件" if len(existing) > 8 else ""
        raise FileExistsError("同名の出力ファイルが既にあります。上書きは行いません。\n" + preview + extra)

    created: list[Path] = []
    staging_dir = Path(tempfile.mkdtemp(prefix=".pdf_split_", dir=output_dir))
    try:
        for page, filename in zip(pages, filenames):
            writer = PdfWriter()
            writer.add_page(page)
            staged = staging_dir / filename
            with staged.open("xb") as stream:
                writer.write(stream)

        for staged in sorted(staging_dir.glob("*.pdf")):
            destination = output_dir / staged.name
            os.rename(staged, destination)
            created.append(destination)
    except Exception as exc:
        for path in created:
            try:
                path.unlink()
            except OSError:
                pass
        raise ValueError(f"分割ファイルを保存できませんでした。\n{exc}") from exc
    finally:
        shutil.rmtree(staging_dir, ignore_errors=True)

    return destinations


def split_pdf(
    pdf_path: Path,
    excel_path: Path,
    output_dir: Path,
    sheet_name: str,
    cell: str,
    build_filenames: Callable[[Path, str, str, int], list[str]],
) -> list[Path]:
    """実行の共通流れ: 入力確認 → PDF読み込み → ファイル名生成 → 書き出し.

    2モードの違いは注入される build_filenames（modes.py）だけ。
    """
    check_input_paths(pdf_path, excel_path, output_dir)
    reader = load_pdf(pdf_path)
    filenames = build_filenames(excel_path, sheet_name, cell, len(reader.pages))
    return write_pages(reader, filenames, output_dir)


