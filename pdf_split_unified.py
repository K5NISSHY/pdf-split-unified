"""統合版 PDF分割ツールのGUI・エントリポイント.

上部に共通入力欄（PDF / Excel / シート / 出力先）、下部に ttk.Notebook による
モード切替タブ2枚、その下にファイル名プレビュー一覧を置く単一ウィンドウ。
タブを切り替えても既存の入力は保持される。ロジックは split_core / modes / preview に置く。
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook

import modes
import preview
import split_core

PREVIEW_VISIBLE_ROWS = 12  # 一覧欄の表示行数（データは全件保持）
STATE_INITIAL = "プレビューを押すと、出力予定のファイル名を一覧に表示します。"
STATE_STALE = "入力変更あり: プレビューを更新してください"


class UnifiedPdfSplitApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("PDF 分割・ファイル名自動生成ツール（統合版）")
        root.resizable(False, False)
        self.pdf_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.base_cell_var = tk.StringVar(value="A1")
        self.per_cell_var = tk.StringVar(value="A1")
        self.preview_rows: list[preview.PreviewRow] = []
        self.preview_active = False

        # 注: tk.Frame/Notebookのpackで渡すpadx/padyは単一値のみ指定可（Python 3.14以降のtkinter）。
        # grid側のpady=(上,下)タプル形式は従来どおり使えるため、詰め幅はgrid側で調整している。
        top = tk.Frame(root, padx=16, pady=12)
        top.pack(fill="x")
        self._file_row(top, 0, "分割するPDF", self.pdf_var, self.choose_pdf)
        self._file_row(top, 1, "Excelファイル", self.excel_var, self.choose_excel)
        tk.Label(top, text="シート", width=20, anchor="w").grid(row=2, column=0, sticky="w", pady=5)
        self.sheet_box = ttk.Combobox(top, textvariable=self.sheet_var, state="readonly", width=56)
        self.sheet_box.grid(row=2, column=1, columnspan=2, sticky="w", padx=6, pady=5)
        self._file_row(top, 3, "出力先フォルダー", self.output_var, self.choose_output)

        buttons = tk.Frame(top)
        buttons.grid(row=4, column=0, columnspan=3, sticky="w")
        tk.Button(buttons, text="プレビュー", width=14, command=self.run_preview).pack(side="left")
        tk.Button(buttons, text="分割を実行", width=14, command=self.run).pack(side="left", padx=(10, 0))
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="x", padx=16, pady=(8, 4))
        tab_seq = tk.Frame(self.notebook, padx=12, pady=10)
        tab_cell = tk.Frame(self.notebook, padx=12, pady=10)
        self.notebook.add(tab_seq, text=modes.MODE_LABELS[modes.BASE_SEQ])
        self.notebook.add(tab_cell, text=modes.MODE_LABELS[modes.PER_CELL])
        tk.Label(
            tab_seq, text="選択したセルの値をファイル名に使用します（例: ABC123_001.pdf）。", anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        tk.Label(tab_seq, text="セル番地", width=20, anchor="w").grid(row=1, column=0, sticky="w")
        tk.Entry(tab_seq, textvariable=self.base_cell_var, width=20).grid(row=1, column=1, sticky="w")
        tk.Label(
            tab_cell,
            text="開始セルから下方向へ、PDFのページ数分のセル値をファイル名に使います（連番なし）。",
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        tk.Label(tab_cell, text="開始セル", width=20, anchor="w").grid(row=1, column=0, sticky="w")
        tk.Entry(tab_cell, textvariable=self.per_cell_var, width=20).grid(row=1, column=1, sticky="w")

        list_frame = tk.Frame(root, padx=16, pady=8)
        list_frame.pack(fill="both", expand=True)
        self.preview_tree = ttk.Treeview(
            list_frame, columns=("page", "filename", "status"),
            show="headings", height=PREVIEW_VISIBLE_ROWS,
        )
        self.preview_tree.heading("page", text="ページ番号")
        self.preview_tree.heading("filename", text="出力ファイル名")
        self.preview_tree.heading("status", text="状態")
        self.preview_tree.column("page", width=90, anchor="e", stretch=False)
        self.preview_tree.column("filename", width=480, anchor="w")
        self.preview_tree.column("status", width=130, anchor="w", stretch=False)
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=scroll.set)
        self.preview_tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        bottom = tk.Frame(root, padx=16, pady=8)
        bottom.pack(fill="x")
        self.status_label = tk.Label(bottom, text=STATE_INITIAL, anchor="w")
        self.status_label.pack(fill="x")
        self.error_label = tk.Label(bottom, text="", anchor="w", justify="left")
        self.error_label.pack(fill="x")

        # FR-15: 共通欄・セル欄・タブが変わったら一覧を消して「要更新」表示にする
        for variable in (self.pdf_var, self.excel_var, self.output_var, self.sheet_var,
                         self.base_cell_var, self.per_cell_var):
            variable.trace_add("write", lambda *_: self.invalidate_preview())
        self.notebook.bind("<<NotebookTabChanged>>", lambda *_: self.invalidate_preview())

    @staticmethod
    def _file_row(frame: tk.Frame, row: int, label: str, variable: tk.StringVar, command) -> None:
        tk.Label(frame, text=label, width=20, anchor="w").grid(row=row, column=0, sticky="w", pady=5)
        tk.Entry(frame, textvariable=variable, width=56).grid(row=row, column=1, padx=6, pady=5)
        tk.Button(frame, text="参照...", command=command).grid(row=row, column=2, pady=5)
    def choose_pdf(self) -> None:
        path = filedialog.askopenfilename(title="分割するPDFを選択", filetypes=[("PDF", "*.pdf")])
        if path:
            self.pdf_var.set(path)
            if not self.output_var.get():
                self.output_var.set(str(Path(path).parent))

    def choose_excel(self) -> None:
        path = filedialog.askopenfilename(title="Excelファイルを選択", filetypes=[("Excelブック", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            workbook = load_workbook(path, read_only=True, data_only=True)
            try:
                sheet_names = workbook.sheetnames
                active_sheet = workbook.active.title
            finally:
                workbook.close()
        except Exception as exc:
            messagebox.showerror("Excelを読み込めません", f"Excelファイルを開けませんでした。\n{exc}")
            return
        if not sheet_names:
            messagebox.showerror("シートがありません", "選択したExcelファイルにシートがありません。")
            return
        self.excel_var.set(path)
        self.sheet_box["values"] = sheet_names
        self.sheet_var.set(active_sheet if active_sheet in sheet_names else sheet_names[0])

    def choose_output(self) -> None:
        path = filedialog.askdirectory(title="出力先フォルダーを選択")
        if path:
            self.output_var.set(path)

    def current_mode(self) -> str:
        """現在アクティブなタブのモードID."""
        return modes.BASE_SEQ if self.notebook.index(self.notebook.select()) == 0 else modes.PER_CELL

    def cell_var_for(self, mode: str) -> tk.StringVar:
        return self.base_cell_var if mode == modes.BASE_SEQ else self.per_cell_var

    def check_inputs(self, mode: str) -> bool:
        """未入力欄があれば警告ダイアログを出して False を返す."""
        if not all((self.pdf_var.get(), self.excel_var.get(), self.output_var.get(),
                    self.sheet_var.get(), self.cell_var_for(mode).get().strip())):
            cell_label = "セル番地" if mode == modes.BASE_SEQ else "開始セル"
            messagebox.showwarning(
                "入力が必要です",
                f"PDF、Excel、シート、{cell_label}、出力先フォルダーをすべて指定してください。",
            )
            return False
        return True

    def inputs(self, mode: str) -> tuple[Path, Path, Path, str, str]:
        """現在の入力（その時点の値で毎回使う）."""
        return (
            Path(self.pdf_var.get()),
            Path(self.excel_var.get()),
            Path(self.output_var.get()),
            self.sheet_var.get(),
            self.cell_var_for(mode).get(),
        )

    def invalidate_preview(self) -> None:
        """FR-15: 入力変更を一覧から消して「要更新」表示にする."""
        self.preview_active = False
        self.preview_rows = []
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.status_label.configure(text=STATE_STALE)
        self.error_label.configure(text="")

    def _show_rows(self) -> None:
        shown, extra = preview.truncate_display(self.preview_rows)
        for row in shown:
            self.preview_tree.insert("", "end", values=row.display_cells())
        if extra:
            self.preview_tree.insert("", "end", values=("", "", f"ほか {extra} 件"))
        self.status_label.configure(
            text=(
                f"プレビュー: {len(self.preview_rows)} 件"
                + (f"（一覧表示 {len(shown)} 件・ほか {extra} 件、書き出しは行っていません）" if extra
                   else "（書き出しは行っていません）")
            )
        )

    def run_preview(self) -> None:
        """FR-14: 現在の入力で出力予定ファイル名を一覧表示（書き出しは一切しない）."""
        mode = self.current_mode()
        if not self.check_inputs(mode):
            return
        pdf_path, excel_path, output_dir, sheet_name, cell = self.inputs(mode)
        try:
            self.preview_rows = preview.build_preview(
                excel_path, sheet_name, cell, mode, output_dir, pdf_path,
            )
        except Exception as exc:
            self.preview_active = False
            self.preview_rows = []
            self.preview_tree.delete(*self.preview_tree.get_children())
            self.status_label.configure(text="プレビューできません")
            self.error_label.configure(text=str(exc))
            return
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_active = True
        self.error_label.configure(text="")
        self._show_rows()

    def run(self) -> None:
        """FR-15: 実行時は常にその時点の入力で再計算して分割する."""
        mode = self.current_mode()
        if not self.check_inputs(mode):
            return
        pdf_path, excel_path, output_dir, sheet_name, cell = self.inputs(mode)
        try:
            results = split_core.split_pdf(
                pdf_path, excel_path, output_dir, sheet_name, cell,
                lambda excel, sheet, target_cell, page_count: modes.build_filenames(
                    mode, excel, sheet, target_cell, page_count,
                ),
            )
        except Exception as exc:
            self.invalidate_preview()
            messagebox.showerror("分割できませんでした", str(exc))
            return
        self.invalidate_preview()
        messagebox.showinfo("完了", f"{len(results)}ページを分割しました。\n保存先: {results[0].parent}")


def main() -> None:
    root = tk.Tk()
    UnifiedPdfSplitApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()


