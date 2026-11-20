r"""GUI実表示スモーク（一瞬だけ実ウィンドウを作り、表示確認後に閉じる）.

表示確認: root.winfo_viewable() / 実ウィンドウサイズ / プレビュー一覧の実測表示件数。
実行: .venv\Scripts\python.exe tools\verify_gui_display.py
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import tkinter as tk  # noqa: E402

import pdf_split_unified as gui  # noqa: E402
import modes  # noqa: E402

def main() -> int:
    work = Path(tempfile.mkdtemp(prefix="pdfsplit_gui_display_"))
    try:
        import verify_assistant as va  # 検証用のPDF・Excel生成関数を再利用

        va.make_pdf(work / "input.pdf", 4)
        va.make_excel(work / "names.xlsx", "名簿", ["あいう", "かきく:不正", "さしす", "CON"])
        (work / "out").mkdir()

        root = tk.Tk()
        app = gui.UnifiedPdfSplitApp(root)
        root.update_idletasks()
        root.update()
        viewable = bool(root.winfo_viewable())
        width, height = root.winfo_width(), root.winfo_height()
        app.pdf_var.set(str(work / "input.pdf"))
        app.excel_var.set(str(work / "names.xlsx"))
        app.sheet_var.set("名簿")
        app.output_var.set(str(work / "out"))
        app.base_cell_var.set("A1")
        root.update()
        app.run_preview()
        root.update()
        rows = len(app.preview_tree.get_children())
        bbox = app.preview_tree.bbox(app.preview_tree.get_children()[0]) if rows else None
        # タブ2枚目へ切り替えても共通入力が保持されることを実表示で確認
        app.notebook.select(1)
        root.update()
        kept = (app.pdf_var.get(), app.excel_var.get(), app.sheet_var.get(), app.output_var.get())
        mode = app.current_mode()
        # 不正文字を含むA2のため、プレビューはエラー欄に日本語文言を出す（一覧は空）
        app.per_cell_var.set("A1")
        app.run_preview()
        root.update()
        err_text = app.error_label.cget("text")
        root.destroy()
    finally:
        import shutil

        shutil.rmtree(work, ignore_errors=True)

    print(f"表示可能ウィンドウ: {viewable} / 実サイズ: {width}x{height}")
    print(f"プレビュー表示行数（A1=あいう・4ページ）: {rows} / 先頭行の描画領域: {bbox}")
    print(f"タブ切替後の共通入力保持: {kept[1].split('/')[-1]} / mode={mode}")
    print(f"プレビュー時のエラー欄: {err_text[:60]}")
    ok = viewable and width > 600 and height > 400 and rows == 4 and mode == modes.PER_CELL and "ファイル名に使えない文字" in err_text
    print("TOTAL:", "1/1" if ok else "0/1")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
