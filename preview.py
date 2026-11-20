"""ファイル名プレビューの生成（読み取り専用。書き出しは一切行わない）.

FR-14 の判定ロジック（`[既存]` / `[重複]`）はすべてこのファイルに置き、
GUI（pdf_split_unified.py）は受け取った行を表示するだけにする。
ファイル名生成は実行時と同一の関数（modes.build_filenames）を共有する。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from modes import build_filenames
from split_core import load_pdf

MAX_ANNOUNCED = 100  # 一覧にそのまま並べる件数。超分は「ほか N 件」にまとめる


@dataclass(frozen=True)
class PreviewRow:
    """プレビュー1行. exists=出力先に既存, duplicate=大文字小文字無視で同一名が複数."""

    page_no: int
    filename: str
    exists: bool
    duplicate: bool

    @property
    def status(self) -> str:
        """表示用の状態文字列（`[既存]` `[重複]`）."""
        marks = []
        if self.exists:
            marks.append("[既存]")
        if self.duplicate:
            marks.append("[重複]")
        return " ".join(marks)

    def display_cells(self) -> tuple[str, str, str]:
        """Treeview へ渡す (ページ番号, ファイル名, 状態) タプル."""
        return (str(self.page_no), self.filename, self.status)


def truncate_display(rows: list[PreviewRow]) -> tuple[list[PreviewRow], int]:
    """表示用の先頭 MAX_ANNOUNCED 件と、まとめられた「ほか」件数を返す.

    rows（データ）自体は全件保持され、この関数は書き出しや破壊的変更をしない。
    """
    if len(rows) <= MAX_ANNOUNCED:
        return rows, 0
    return rows[:MAX_ANNOUNCED], len(rows) - MAX_ANNOUNCED


def build_preview(
    excel_path: Path,
    sheet_name: str,
    cell: str,
    mode: str,
    output_dir: Path,
    pdf_path: Path,
) -> list[PreviewRow]:
    """現在の入力からプレビュー行を全件組み立てる.

    PDFとExcelを読むだけで、ファイルの書き出し・既存ファイルの変更は行わない。
    FR-6 相当の不正（空欄・不正文字・予約名・数式キャッシュ無し）は
    ValueError をそのまま伝搬する（GUI のエラー欄に表示される）。
    """
    reader = load_pdf(pdf_path)
    # 実行時と同一の生成関数を使う。プレビューでは重複をエラーにせず、下の重複判定で `[重複]` 印を付ける。
    filenames = build_filenames(mode, excel_path, sheet_name, cell, len(reader.pages), abort_on_duplicates=False)

    counts: dict[str, int] = {}
    for filename in filenames:
        key = filename.casefold()
        counts[key] = counts.get(key, 0) + 1

    return [
        PreviewRow(
            page_no=page_no,
            filename=filename,
            exists=(output_dir / filename).exists(),
            duplicate=counts[filename.casefold()] > 1,
        )
        for page_no, filename in enumerate(filenames, start=1)
    ]
