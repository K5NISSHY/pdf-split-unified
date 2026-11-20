"""2モードの出力ファイル名生成のみを担う.

- `base_seq` (セル値＋連番): 指定セル1つの値をベース名に `{ベース名}_{ページ番号:03d}.pdf`
  （先頭は 001、1000ページ超では桁数が伸びる）
- `per_cell` (セル別・連番なし): 開始セルから同じ列を下方向へページ数分読み `{セル値}.pdf`
  （大文字小文字を区別しない重複名はエラー）

読み取り・検証・書き出しなどの共通処理は split_core.py を使う。
"""

from __future__ import annotations

from pathlib import Path

from split_core import (
    BASE_CELL_MESSAGES,
    COLUMN_CELL_MESSAGES,
    EXCEL_MAX_ROW,
    parse_cell_address,
    read_cell_values,
)

BASE_SEQ = "base_seq"
PER_CELL = "per_cell"

MODE_LABELS = {
    BASE_SEQ: "セル値＋連番",
    PER_CELL: "セル別（連番なし）",
}


def build_filenames_base_seq(
    excel_path: Path,
    sheet_name: str,
    cell: str,
    page_count: int,
    *,
    abort_on_duplicates: bool = True,
) -> list[str]:
    """セル値＋連番モード。1つのセル値を全ページのベース名にする.

    連番で全ファイル名が一意になるため重複は起こり得ず、abort_on_duplicates は受け取るだけで影響しない
    （共通形を揃えるためだけにある）。
    """
    if page_count < 1:
        raise ValueError("PDFにページがありません。")
    column, row = parse_cell_address(cell, "セル番地")
    address = f"{column}{row}"
    values = read_cell_values(excel_path, sheet_name, [address], BASE_CELL_MESSAGES)
    base_name = values[address]
    return [f"{base_name}_{page_no:03d}.pdf" for page_no in range(1, page_count + 1)]
def build_filenames_per_cell(
    excel_path: Path,
    sheet_name: str,
    cell: str,
    page_count: int,
    *,
    abort_on_duplicates: bool = True,
) -> list[str]:
    """セル別モード。開始セルから同列を下行きページ数分のセル値をそのまま名前にする.

    実行時は abort_on_duplicates=True（既定）で大文字小文字無視の重複名をエラーにする。
    プレビューのみ False を渡して重複行を残す（FR-14 の `[重複]` 印は preview.py が付ける）。
    """
    if page_count < 1:
        raise ValueError("PDFにページがありません。")
    column, first_row = parse_cell_address(cell, "開始セル")
    last_row = first_row + page_count - 1
    if last_row > EXCEL_MAX_ROW:
        raise ValueError(f"開始セルから{page_count}ページ分のセルを読み取ると、Excelの最終行を超えます。")

    addresses = [f"{column}{first_row + offset}" for offset in range(page_count)]
    values = read_cell_values(excel_path, sheet_name, addresses, COLUMN_CELL_MESSAGES)
    names = [f"{values[address]}.pdf" for address in addresses]

    if abort_on_duplicates:
        seen: set[str] = set()
        duplicates: list[str] = []
        for filename in names:
            key = filename.casefold()
            if key in seen and filename not in duplicates:
                duplicates.append(filename)
            seen.add(key)
        if duplicates:
            raise ValueError("Excelに同じファイル名が複数あります。重複を解消してください。\n" + "\n".join(duplicates))
    return names


_BUILDERS = {
    BASE_SEQ: build_filenames_base_seq,
    PER_CELL: build_filenames_per_cell,
}


def build_filenames(
    mode: str,
    excel_path: Path,
    sheet_name: str,
    cell: str,
    page_count: int,
    *,
    abort_on_duplicates: bool = True,
) -> list[str]:
    """モードIDでファイル名生成を差し替える（両モードで共通形）.

    abort_on_duplicates=True（既定）は実行時の検査。プレビューのみ
    重複行を残すため False を渡す（preview.py 参照）。
    """
    try:
        builder = _BUILDERS[mode]
    except KeyError:
        raise ValueError(f"未対応のモードです: {mode}") from None
    return builder(excel_path, sheet_name, cell, page_count, abort_on_duplicates=abort_on_duplicates)

