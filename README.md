# Excel-CSV-Compile-Tool
This tool is designed to help professionals save hours of manual work and keep their compiled data clean, consistent, and reliable.Cross-System Ready – Packed into a .exe, runs on any Windows PC without Python.


# -*- coding: utf-8 -*-
"""
Excel/CSV Compiler Tool (English UI without Duplicate Removal)
--------------------------------------------------------------
This tool merges all Excel/CSV files (.xlsx/.xls/.csv) in a selected folder into a single file,
but only includes those files whose column headings match 100%.

- Header matching is case-insensitive and ignores extra spaces.
- The final column order follows the first valid file.
- Option: choose which row contains column headers (Excel-style row number).

Requirements:
- Python 3.9+
- pip install pandas openpyxl xlrd==1.2.0
"""
import os
import re
import glob
import traceback
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "Excel/CSV Compiler"
VALID_EXTS = (".xlsx", ".xls", ".csv")

def normalize_col(name):
    """Normalize column header: strip, collapse spaces, lower-case."""
    if name is None:
        name = ""
    name = str(name)
    name = re.sub(r"\s+", " ", name.strip()).lower()
    return name

def read_excel_headers(path, header_row=0):
    """Read only headers from the given row of an Excel/CSV file."""
    try:
        ext = os.path.splitext(path)[1].lower()
        if ext == ".xlsx":
            df_head = pd.read_excel(path, header=header_row, nrows=0, engine="openpyxl")
        elif ext == ".xls":
            df_head = pd.read_excel(path, header=header_row, nrows=0, engine="xlrd")
        elif ext == ".csv":
            df_head = pd.read_csv(path, header=header_row, nrows=0)
        else:
            raise RuntimeError("Unsupported file type")
        return list(df_head.columns)
    except Exception as e:
        raise RuntimeError(f"Header read failed: {e}")

def read_excel(path, want_cols_order, header_row=0):
    """Read full Excel/CSV file and reorder columns to match the baseline order."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        df = pd.read_excel(path, header=header_row, engine="openpyxl")
    elif ext == ".xls":
        df = pd.read_excel(path, header=header_row, engine="xlrd")
    elif ext == ".csv":
        df = pd.read_csv(path, header=header_row)
    else:
        raise RuntimeError("Unsupported file type")

    norm_to_orig = {normalize_col(c): c for c in df.columns}
    missing = [c for c in want_cols_order if c not in norm_to_orig]
    if missing:
        raise RuntimeError(f"Columns missing after load: {missing}")
    ordered_cols = [norm_to_orig[c] for c in want_cols_order]
    return df[ordered_cols]

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("780x560")
        self.resizable(True, True)

        self.folder_path = tk.StringVar()
        self.header_row = tk.IntVar(value=1)  # Excel-style default: Row 1

        self._build_ui()

    def _build_ui(self):
        pad = 8

        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", padx=pad, pady=pad)

        ttk.Label(frm_top, text="Select folder containing Excel/CSV files:", font=("Segoe UI", 10, "bold")).pack(anchor="w")

        frm_pick = ttk.Frame(frm_top)
        frm_pick.pack(fill="x", pady=4)
        ttk.Entry(frm_pick, textvariable=self.folder_path).pack(side="left", fill="x", expand=True)
        ttk.Button(frm_pick, text="Browse…", command=self.choose_folder).pack(side="left", padx=6)

        frm_opts2 = ttk.Frame(self)
        frm_opts2.pack(fill="x", padx=pad, pady=2)
        ttk.Label(frm_opts2, text="Header Row (Excel-style, e.g., 1 = first row):").pack(side="left")
        ttk.Entry(frm_opts2, textvariable=self.header_row, width=5).pack(side="left", padx=4)

        frm_btns = ttk.Frame(self)
        frm_btns.pack(fill="x", padx=pad, pady=pad)
        ttk.Button(frm_btns, text="Start Compile", command=self.compile_now).pack(side="left")
        ttk.Button(frm_btns, text="Exit", command=self.destroy).pack(side="right")

        ttk.Separator(self).pack(fill="x", padx=pad, pady=pad)

        ttk.Label(self, text="Progress / Status:", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=pad)
        self.txt = tk.Text(self, height=18, wrap="word")
        self.txt.pack(fill="both", expand=True, padx=pad, pady=(0, pad))

        self._log("Ready. Choose a folder, set header row, and click 'Start Compile'.")

    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.txt.insert("end", f"[{ts}] {msg}\n")
        self.txt.see("end")
        self.update_idletasks()

    def choose_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)
            self._log(f"Selected folder: {folder}")

    def _collect_files(self, folder):
        files = []
        for ext in VALID_EXTS:
            files.extend(glob.glob(os.path.join(folder, f"*{ext}")))
        files = [f for f in files if not os.path.basename(f).startswith("~$")]
        return sorted(files)

    def _headers_match(self, cols_a, cols_b):
        A = [normalize_col(c) for c in cols_a]
        B = [normalize_col(c) for c in cols_b]
        return set(A) == set(B)

    def compile_now(self):
        folder = self.folder_path.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror(APP_TITLE, "Please select a valid folder.")
            return

        header_row_excel = self.header_row.get()
        if header_row_excel < 1:
            messagebox.showerror(APP_TITLE, "Header Row must be >= 1 (Excel-style).")
            return
        header_row = header_row_excel - 1  # convert Excel-style to 0-based for pandas

        files = self._collect_files(folder)
        if not files:
            messagebox.showwarning(APP_TITLE, "No .xlsx/.xls/.csv files found in this folder.")
            return

        self._log(f"Total files found: {len(files)}")
        self._log("Files detected: " + ", ".join([os.path.basename(f) for f in files]))

        skipped, included = [], []
        baseline_file = baseline_cols = baseline_norm = None

        # Step 1: pick baseline
        for path in files:
            try:
                cols = read_excel_headers(path, header_row)
                baseline_file = path
                baseline_cols = cols
                baseline_norm = [normalize_col(c) for c in cols]
                self._log(f"Baseline file: {os.path.basename(path)}")
                break
            except Exception as e:
                self._log(f"SKIP (header read failed): {os.path.basename(path)} -> {e}")
                skipped.append((path, f"Header read fail: {e}"))

        if not baseline_file:
            messagebox.showerror(APP_TITLE, "No file could be read for headers.")
            return

        data_frames = []

        try:
            df0 = read_excel(baseline_file, baseline_norm, header_row)
            data_frames.append(df0)
            included.append(baseline_file)
            self._log(f"Included: {os.path.basename(baseline_file)} (rows: {len(df0)})")
        except Exception as e:
            self._log(f"SKIP (read failed): {os.path.basename(baseline_file)} -> {e}")
            skipped.append((baseline_file, f"Read fail: {e}"))

        # Step 2: process remaining
        for path in files:
            if path == baseline_file:
                continue
            try:
                cols = read_excel_headers(path, header_row)
            except Exception as e:
                self._log(f"SKIP (header read failed): {os.path.basename(path)} -> {e}")
                skipped.append((path, f"Header read fail: {e}"))
                continue

            if not self._headers_match(baseline_cols, cols):
                self._log(f"SKIP (header mismatch): {os.path.basename(path)}")
                skipped.append((path, "Header mismatch"))
                continue

            try:
                df = read_excel(path, baseline_norm, header_row)
                data_frames.append(df)
                included.append(path)
                self._log(f"Included: {os.path.basename(path)} (rows: {len(df)})")
            except Exception as e:
                self._log(f"SKIP (read failed): {os.path.basename(path)} -> {e}")
                skipped.append((path, f"Read fail: {e}"))

        if not data_frames:
            self._log("No files could be included.")
            messagebox.showwarning(APP_TITLE, "No files could be included.")
            return

        final_df = pd.concat(data_frames, ignore_index=True)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"compiled_{ts}.xlsx"
        out_path = os.path.join(folder, out_name)
        try:
            final_df.to_excel(out_path, index=False)
            self._log(f"✅ Saved: {out_path}")
            msg = (
                f"Compile finished!\n\n"
                f"Output file: {out_name}\n"
                f"Included files: {len(included)}\n"
                f"Skipped: {len(skipped)}\n"
                f"Total rows: {len(final_df)}"
            )
            messagebox.showinfo(APP_TITLE, msg)
        except Exception as e:
            self._log(f"❌ Save failed: {e}")
            messagebox.showerror(APP_TITLE, f"Save failed: {e}")

        self._log("---- Summary ----")
        self._log(f"Included ({len(included)}):")
        for p in included:
            self._log(" - " + os.path.basename(p))
        if skipped:
            self._log(f"Skipped ({len(skipped)}):")
            for p, reason in skipped:
                self._log(f" - {os.path.basename(p)} -> {reason}")

def main():
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror(APP_TITLE, f"Unexpected error: {e}")

if __name__ == "__main__":
    main()
