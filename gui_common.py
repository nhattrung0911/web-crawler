"""
gui_common.py — HA TANG CHUNG cho GUI crawler (dat o root WORX GREEN).

Muc dich: KHONG copy-paste khoi "chon Excel" vao tung crawler nua.
Crawler tim file nay bang cach di nguoc len cay thu muc:

    import sys, pathlib
    HERE = pathlib.Path(__file__).parent
    ROOT = next(p for p in (HERE, *HERE.parents) if (p / "gui_common.py").exists())
    sys.path.insert(0, str(ROOT))
    from gui_common import attach_excel_picker, ask_export_path

Quy tac (AGENTS.md 4.4):
- Sheet + Cot = DROPDOWN doc thang tu file Excel, KHONG bat user go tay.
- Thu muc anh + file ket qua dat CANH file Excel dau vao.
"""
from __future__ import annotations

import pathlib
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any

import pandas as pd

CODE_COL_RE = re.compile(r"part|mã|ma\b|code|sku|item", re.I)


def attach_excel_picker(parent: tk.Misc, xls: tk.StringVar, sheet: tk.StringVar,
                        col: tk.StringVar, outdir: tk.StringVar,
                        img_prefix: str = "anh") -> None:
    """Gan khoi: [Chon Excel] + dropdown Sheet + dropdown Cot + o Thu muc anh.

    Chon file -> tu nap sheet_names, tu nap cot cua sheet dau, tu doan cot ma,
    tu dat thu muc anh canh file Excel. Doi sheet -> nap lai cot.
    """
    fr = ttk.Frame(parent)
    fr.pack(anchor="w", pady=4)
    btn = ttk.Button(fr, text="Chon Excel")
    btn.pack(side="left")
    ttk.Label(fr, textvariable=xls, width=24).pack(side="left")

    cbs: dict[str, ttk.Combobox] = {}
    for lbl, var in (("Sheet:", sheet), ("Cot:", col)):
        f = ttk.Frame(parent)
        f.pack(anchor="w")
        ttk.Label(f, text=lbl, width=12).pack(side="left")
        cb = ttk.Combobox(f, textvariable=var, width=18, state="readonly")
        cb.pack(side="left")
        cbs[lbl] = cb

    f = ttk.Frame(parent)
    f.pack(anchor="w")
    ttk.Label(f, text="Thu muc anh:", width=12).pack(side="left")
    ttk.Entry(f, textvariable=outdir, width=20).pack(side="left")
    ttk.Label(parent, text="(sheet/cot doc tu file Excel; anh dat canh file do)",
              foreground="#777").pack(anchor="w")

    def load_cols(*_: Any) -> None:
        if not xls.get() or not sheet.get():
            return
        try:
            cols = list(pd.read_excel(xls.get(), sheet_name=sheet.get(), nrows=0).columns)
        except Exception as e:
            messagebox.showerror("Excel", f"Khong doc duoc sheet '{sheet.get()}':\n{e}")
            return
        cbs["Cot:"]["values"] = cols
        guess = next((c for c in cols if CODE_COL_RE.search(str(c))), "")
        col.set(guess or (cols[0] if cols else ""))

    def pick() -> None:
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if not p:
            return
        try:
            names = pd.ExcelFile(p).sheet_names
        except Exception as e:
            messagebox.showerror("Excel", f"Khong mo duoc file:\n{e}")
            return
        xls.set(p)
        outdir.set(str(pathlib.Path(p).parent / f"{img_prefix}_{pathlib.Path(p).stem}"))
        cbs["Sheet:"]["values"] = names
        sheet.set(names[0] if names else "")
        load_cols()

    btn.configure(command=pick)
    cbs["Sheet:"].bind("<<ComboboxSelected>>", load_cols)


def ask_export_path(excel_in: str, fallback_dir: pathlib.Path,
                    fallback_name: str) -> str:
    """Hop thoai luu ket qua: mac dinh CANH file Excel dau vao, ten <goc>_ketqua.xlsx."""
    src = pathlib.Path(excel_in) if excel_in else None
    return filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
        initialdir=str(src.parent) if src else str(fallback_dir),
        initialfile=f"{src.stem}_ketqua.xlsx" if src else fallback_name)
