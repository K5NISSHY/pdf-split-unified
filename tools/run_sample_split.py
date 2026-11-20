r"""サンプル資料で両モードのPDF分割を実行し、結果を samples\検証結果 に保存する検証スクリプト.

実行: .venv\Scripts\python.exe tools\run_sample_split.py

- samples\sample_5pages.pdf（5ページ・本文に文字あり）と samples\sample_names.xlsx を使う
  （両ファイルは tools\make_samples.py で生成され、無ければこのスクリプトが生成する）
- 「セル値＋連番」は A1、「セル別（連番なし）」は B1 を開始セルとして実際の処理系（split_core / modes）で分割する
- 結果は samples\検証結果\連番 と samples\検証結果\セル別 に保存（実行前に作り直すため再実行可能）
- 各出力が1ページか・本文文字がそのまま残っているか・元PDF/Excelが無変更かを機械判定し、
  最後に TOTAL: <成功>/<総数> を出す（全成功で終了コード0）
"""

from __future__ import annotations

import hashlib
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pypdf import PdfReader  # noqa: E402

import make_samples  # noqa: E402
import modes  # noqa: E402
import split_core  # noqa: E402

SAMPLE_DIR = ROOT / "samples"
PDF_PATH = SAMPLE_DIR / "sample_5pages.pdf"
XLSX_PATH = SAMPLE_DIR / "sample_names.xlsx"
RESULT_ROOT = SAMPLE_DIR / "検証結果"
SEQ_DIR = RESULT_ROOT / "連番"
CELL_DIR = RESULT_ROOT / "セル別"

RESULTS: list[tuple[bool, str]] = []


def check(ok: bool, message: str) -> None:
    RESULTS.append((bool(ok), message))
    print(("PASS: " if ok else "FAIL: ") + message)


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def prepare_dir(folder: Path) -> None:
    """結果フォルダーを作り直す（同名既存ファイルで中止する規則を避けるため）."""
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir(parents=True)


def inspect_outputs(files: list[Path], expect_count: int, label: str) -> None:
    check(len(files) == expect_count and all(p.is_file() for p in files),
          f"{label}: {expect_count}個の実ファイルが出力先に保存されている")
    texts = []
    page_counts = []
    for path in files:
        reader = PdfReader(str(path))
        page_counts.append(len(reader.pages))
        texts.append((reader.pages[0].extract_text() or "").replace("\n", " / "))
    check(page_counts == [1] * expect_count, f"{label}: 各PDFはちょうど1ページ {page_counts}")
    for path, text in zip(files, texts):
        print(f"      - {path.name}: {text[:78]}")
    keep = all(f"Page {n} of 5" in text for text, n in zip(texts, _page_order(files, label)))
    check(keep, f"{label}: 本文のページ指定文字列が正しく各ファイルに残っている")


def _page_order(files: list[Path], label: str) -> list[int]:
    """ファイル名から対応する元ページ番号を復元する（検証用の手掛かり）."""
    if label == "連番":
        return [int(p.stem.rsplit("_", 1)[1]) for p in files]
    index = {name: number + 1 for number, name in enumerate(make_samples.PAGE_NAMES)}
    return [index[p.stem] for p in files]


def main() -> int:
    if not PDF_PATH.is_file() or not XLSX_PATH.is_file():
        make_samples.main()
    before = (digest(PDF_PATH), digest(XLSX_PATH))

    # --- モード1: セル値＋連番（A1 の値をベース名に 001..）
    prepare_dir(SEQ_DIR)
    seq_files = split_core.split_pdf(
        PDF_PATH, XLSX_PATH, SEQ_DIR, "サンプル", "A1",
        lambda excel, sheet, cell, pages: modes.build_filenames(modes.BASE_SEQ, excel, sheet, cell, pages),
    )
    check([p.name for p in seq_files] == [f"{make_samples.BASE_NAME}_{n:03d}.pdf" for n in range(1, 6)],
          f"FR-3 連番モードのファイル名（A1=「{make_samples.BASE_NAME}」）: {seq_files[0].name}〜{seq_files[-1].name}")
    inspect_outputs(seq_files, 5, "連番")

    # --- モード2: セル別（連番なし・B1から下行き）
    prepare_dir(CELL_DIR)
    cell_files = split_core.split_pdf(
        PDF_PATH, XLSX_PATH, CELL_DIR, "サンプル", "B1",
        lambda excel, sheet, cell, pages: modes.build_filenames(modes.PER_CELL, excel, sheet, cell, pages),
    )
    check([p.name for p in cell_files] == [f"{name}.pdf" for name in make_samples.PAGE_NAMES],
          f"FR-5 セル別モードのファイル名（B1:B5）: {' / '.join(p.name for p in cell_files)}")
    inspect_outputs(cell_files, 5, "セル別")

    # --- 安全性: 元PDF・元Excelは無変更／同一フォルダー再実行は上書き中止
    check((digest(PDF_PATH), digest(XLSX_PATH)) == before, "FR-9 元のサンプルPDF・Excelは無変更（ハッシュ一致）")
    seq_digest = {p.name: digest(p) for p in seq_files}
    try:
        split_core.split_pdf(
            PDF_PATH, XLSX_PATH, SEQ_DIR, "サンプル", "A1",
            lambda excel, sheet, cell, pages: modes.build_filenames(modes.BASE_SEQ, excel, sheet, cell, pages),
        )
        check(False, "FR-10 再実行を中止していない")
    except FileExistsError as exc:
        first = str(exc).splitlines()[1] if len(str(exc).splitlines()) > 1 else ""
        check("上書きは行いません" in str(exc), f"FR-10 結果フォルダーへの再実行は中止（上書きなし）: 例 {first}")
    check(all(digest(p) == seq_digest[p.name] for p in SEQ_DIR.glob("*.pdf")),
          "FR-10 中止後も保存済みの分割結果は無変更")
    check(sorted(p.name for p in RESULT_ROOT.iterdir()) == ["セル別", "連番"],
          f"検証結果フォルダーの構成: {' / '.join(sorted(p.name for p in RESULT_ROOT.iterdir()))}")
    leftovers = [p.name for folder in (SEQ_DIR, CELL_DIR) for p in folder.iterdir() if p.name.startswith(".pdf_split_")]
    check(leftovers == [], "FR-12 一時フォルダー（.pdf_split_*）は残っていない")

    succeeded = sum(1 for ok, _ in RESULTS if ok)
    print(f"TOTAL: {succeeded}/{len(RESULTS)}")
    return 0 if succeeded == len(RESULTS) else 1


if __name__ == "__main__":
    raise SystemExit(main())
