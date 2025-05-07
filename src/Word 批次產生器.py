# -*- coding: utf-8 -*-
"""
Created on Thu May  8 01:26:39 2025

@author: Steven, Hsin
@email: steveh8758@gmail.com
"""

from __future__ import annotations
import contextlib
import logging
import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List

from win32com.client import Dispatch


logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


# ========================================= Main Functions =========================================
# fmt: off
def load_excel(xl_app,
               xl_path: Path,
               sheet: str) -> List[Dict[str, str]]:
# fmt: on
    wb = xl_app.Workbooks.Open(str(xl_path))

    try:
        ws = xl_app.Worksheets(sheet)
        headers = []
        col = 1
        while header := ws.Cells(1, col).Value:
            headers.append(str(header).strip())
            col += 1
        if not headers:
            raise ValueError("無法讀取 excel 內書籤名稱。")

        records = []
        row = 2
        while ws.Cells(row, 1).Value not in (None, ""):
            records.append(
                {
                    hdr: "" if (val := ws.Cells(row, idx).Value) is None else str(val)
                    for idx, hdr in enumerate(headers, start=1)
                }
            )
            row += 1
        return records

    finally:
        wb.Close()

# fmt: off
def fill_docs(word_app,
              template: Path,
              out_dir: Path,
              records: List[Dict[str, str]],
              progress_cb = lambda x: None,
              f_name_prefix: str = "Output") -> None:
# fmt: on
    out_dir.mkdir(exist_ok=True)
    total = len(records)
    for idx, rec in enumerate(records, 1):
        doc = word_app.Documents.Add(str(template))
        for bm, txt in rec.items():
            if doc.Bookmarks.Exists(bm):
                doc.Bookmarks(bm).Range.Text = txt
        doc.SaveAs(str(out_dir / f"{idx:0{len(str(total))}d}_{f_name_prefix}.docx"))
        doc.Close()
        progress_cb(idx / total if total else 1)


# ============================================== Gui ===============================================
class App(tk.Tk):
    def __init__(self,
                 title: str,
                 *,
                 btn_width: int = 16,
                 debug: bool = False):
        
        messagebox.showinfo(
            "使用說明",
            """
            請依序選擇 Excel 檔案、Word 模板與輸出資料夾，設定好後點擊「執行」即可產生文件。

            1. Word 內需要事先加入書籤定位要被導入的地方。
            2. Excel 內需要在第一列輸入書籤名稱，如下：

            -----------------------------
            |ㅤ時間ㅤ|ㅤ名字ㅤ|ㅤ預算ㅤ|
            |ㅤ上午ㅤ|ㅤ小明ㅤ|ㅤ100ㅤ|
            |ㅤ中午ㅤ|ㅤ曉華ㅤ|ㅤ300ㅤ|
            |ㅤ下午ㅤ|ㅤ小美ㅤ|ㅤ600ㅤ|
            -----------------------------
            """.replace(" ", "").strip("\n")
        )
        
        super().__init__()
        self.title(title)
        self.resizable(False, False)
        self.eval('tk::PlaceWindow . center')
        
        # --------------------- Vars ---------------------
        self.debug = debug
        self.excel_path: Path | None = None
        self.template_path: Path | None = None
        self.out_dir: Path | None = None

        # -------------------- Style ---------------------
        tkfont.nametofont("TkDefaultFont").configure(size=12, family="Microsoft JhengHei")
        style = ttk.Style(self)
        style.configure("TButton", padding=4)
        style.configure("TLabelframe.Label", font=("Microsoft JhengHei", 12))


        # -------------------- Excel ---------------------
        self.f_excel = ttk.Labelframe(self, text="📊 Excel 設定")
        self.f_excel.grid(row=0, column=0, padx=12, pady=6, sticky="ew")
        ttk.Button(self.f_excel, text="選擇 Excel", width=btn_width, command=self.pick_excel).grid(
            row=0, column=0, padx=4, pady=4, sticky="w")
        self.lbl_excel = ttk.Label(self.f_excel, text="未選擇", width=30)
        self.lbl_excel.grid(row=0, column=1, padx=4, sticky="w")
        ttk.Label(self.f_excel, text="工作表名稱：", width=btn_width, anchor="e").grid(
            row=1, column=0, padx=4, pady=4, sticky="e")
        self.sheet_var = tk.StringVar(value="工作表1")
        ttk.Entry(self.f_excel, textvariable=self.sheet_var, width=18).grid(
            row=1, column=1, padx=4, sticky="w" )

        # --------------------- Word ---------------------
        self.f_word = ttk.Labelframe(self, text="📑 Word 模板")
        self.f_word.grid(row=1, column=0, padx=12, pady=6, sticky="ew")
        ttk.Button(self.f_word, text="選擇 Word 模板", width=btn_width, command=self.pick_template).grid(
            row=0, column=0, padx=4, pady=4, sticky="w")
        self.lbl_tpl = ttk.Label(self.f_word, text="未選擇", width=30)
        self.lbl_tpl.grid(row=0, column=1, padx=4, sticky="w")
        self.f_word.grid_remove()

        # -------------------- Output --------------------
        self.f_out = ttk.Labelframe(self, text="📂 輸出設定")
        self.f_out.grid(row=2, column=0, padx=12, pady=6, sticky="ew")
        ttk.Button(self.f_out, text="輸出資料夾", width=btn_width, command=self.pick_outdir).grid(
            row=0, column=0, padx=4, pady=4, sticky="w")
        self.lbl_out = ttk.Label(self.f_out, text="未選擇", width=30)
        self.lbl_out.grid(row=0, column=1, padx=4, sticky="w")
        ttk.Label(self.f_out, text="檔案前墜：", width=btn_width, anchor="e").grid(
            row=1, column=0, padx=4, pady=4, sticky="e")
        self.prefix_var = tk.StringVar(value="Output")
        ttk.Entry(self.f_out, textvariable=self.prefix_var, width=18).grid(
            row=1, column=1, padx=4, sticky="w")
        self.f_out.grid_remove()

        # ------------------- Process --------------------
        self.progress = ttk.Progressbar(self, length=420, mode="determinate")
        self.progress.grid(row=3, column=0, padx=12, pady=6)
        self.btn_run = ttk.Button(self, text="執行", width=btn_width, command=self.run, state="disabled")
        self.btn_run.grid(row=4, column=0, pady=(0, 12))
        
    # ----------------------------------- Event ------------------------------------
    def pick_excel(self):
        if path := filedialog.askopenfilename(title="選擇 Excel", filetypes=[("Excel", "*.xls*")]):
            self.excel_path = Path(path)
            self.lbl_excel.config(text=self.excel_path.name)
            self._update_visibility()

    def pick_template(self):
        if path := filedialog.askopenfilename(title="選擇 Word 模板", filetypes=[("Word", "*.dotx *.docx")]):
            self.template_path = Path(path)
            self.lbl_tpl.config(text=self.template_path.name)
            self._update_visibility()

    def pick_outdir(self):
        if path := filedialog.askdirectory(title="選擇輸出資料夾"):
            self.out_dir = Path(path)
            self.lbl_out.config(text=self.out_dir)
            self._update_visibility()

    # --------------------------------- Animation ----------------------------------
    def _update_visibility(self):
        if self.excel_path and not self.f_word.winfo_ismapped():
            self.f_word.grid()
        if self.template_path and not self.f_out.winfo_ismapped():
            self.f_out.grid()
        if self.excel_path and self.template_path and self.out_dir:
            self.btn_run.config(state="normal")

    # ------------------------------------ Run -------------------------------------
    def _update_progress(self, ratio: float):
        self.progress["value"] = ratio * 100
        self.progress.update()

    def run(self):
        if not all([self.excel_path, self.template_path, self.out_dir]):
            messagebox.showwarning("資料不完整", "請正確選擇 Excel、Word 模板與輸出資料夾！")
            return

        xl_app = Dispatch("Excel.Application")
        word_app = Dispatch("Word.Application")
        if self.debug:
            xl_app.Visible = word_app.Visible = True

        try:
            records = load_excel(xl_app, self.excel_path, self.sheet_var.get())
            if not records:
                raise ValueError("找不到資料列！")
            self.progress["value"] = 0

            fill_docs(word_app,
                      self.template_path,
                      self.out_dir,
                      records,
                      progress_cb=lambda v: self._update_progress(v),
                      f_name_prefix=self.prefix_var.get())
            messagebox.showinfo("完成", f"已成功產生 {len(records)} 份文件！")

        except Exception as exc:
            logging.exception(exc)
            messagebox.showerror("發生錯誤", str(exc))

        finally:
            with contextlib.suppress(Exception):
                xl_app.Quit()
            with contextlib.suppress(Exception):
                word_app.Quit()
            self.progress["value"] = 0


# ============================================ Entrance ============================================
if __name__ == "__main__":
    App("Word 批次產生器").mainloop()
