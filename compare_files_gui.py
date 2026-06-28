"""
File Compare Studio — modern desktop GUI
------------------------------------------
A 2026-style dark/light GUI for comparing two data files (CSV/XLSX) on a
chosen key column, with optional pre-comparison filters, producing 3 files:
  - rows only in File 1
  - rows only in File 2
  - rows matched in both files

The original source files are only READ — never modified.

Install requirements:
    pip install customtkinter pandas openpyxl

Run:
    python compare_files_gui.py
"""

import os
import sys
import traceback
import subprocess
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd

# ----------------------------------------------------------------------------
# Theme / palette  (tuples are (light_mode, dark_mode))
# ----------------------------------------------------------------------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BG            = ("#F2F3F9", "#0B0D14")
CARD_BG       = ("#FFFFFF", "#151823")
CARD_BORDER   = ("#E4E7F2", "#242838")
ROW_BG        = ("#F6F7FB", "#1C2030")
TEXT_PRIMARY  = ("#13151D", "#F5F6FA")
TEXT_SECOND   = ("#5B6175", "#9098B1")
ACCENT        = "#7C5CFF"
ACCENT_HOVER  = "#6A47F2"
TEAL          = "#16D6B5"
TEAL_HOVER    = "#10B89A"
DANGER        = "#FF5C7A"
DANGER_HOVER  = "#E84766"

FONT_FAMILY = "Segoe UI"


def F(size, weight="normal"):
    return ctk.CTkFont(family=FONT_FAMILY, size=size, weight=weight)


# ----------------------------------------------------------------------------
# Data helpers (pure logic, independent from the UI)
# ----------------------------------------------------------------------------
def read_file(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        last_err = None
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
            try:
                return pd.read_csv(path, dtype=str, keep_default_na=False, encoding=enc)
            except UnicodeDecodeError as e:
                last_err = e
                continue
        raise ValueError(f"Could not read '{path}' with common encodings.") from last_err
    elif ext in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str)
    else:
        raise ValueError(f"Unsupported file extension '{ext}'. Use .csv, .xlsx or .xls")


def guess_ip_column(columns):
    for c in columns:
        if "ip" in c.lower():
            return c
    return columns[0] if columns else ""


def run_comparison(df1, df2, key1, key2, base1, base2):
    """Returns (only_in_1, only_in_2, common) DataFrames."""
    d1 = df1.copy()
    d2 = df2.copy()
    d1["_match_key"] = d1[key1].astype(str).str.strip().str.lower()
    d2["_match_key"] = d2[key2].astype(str).str.strip().str.lower()

    set1, set2 = set(d1["_match_key"]), set(d2["_match_key"])
    only1, only2, common_keys = set1 - set2, set2 - set1, set1 & set2

    only_in_1 = d1[d1["_match_key"].isin(only1)].drop(columns=["_match_key"]).reset_index(drop=True)
    only_in_2 = d2[d2["_match_key"].isin(only2)].drop(columns=["_match_key"]).reset_index(drop=True)

    m1 = d1[d1["_match_key"].isin(common_keys)].reset_index(drop=True)
    m2 = d2[d2["_match_key"].isin(common_keys)].reset_index(drop=True)
    common = pd.merge(m1, m2, on="_match_key", suffixes=(f"_{base1}", f"_{base2}")).drop(columns=["_match_key"])

    return only_in_1, only_in_2, common


# ----------------------------------------------------------------------------
# Small reusable UI bits
# ----------------------------------------------------------------------------
class StepHeader(ctk.CTkFrame):
    def __init__(self, master, number, title, subtitle=""):
        super().__init__(master, fg_color="transparent")
        bubble = ctk.CTkLabel(
            self, text=str(number), width=28, height=28, corner_radius=14,
            fg_color=ACCENT, text_color="#FFFFFF", font=F(13, "bold"),
        )
        bubble.grid(row=0, column=0, rowspan=2 if subtitle else 1, padx=(0, 10), sticky="n")
        ctk.CTkLabel(self, text=title, font=F(16, "bold"), text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=1, sticky="w"
        )
        if subtitle:
            ctk.CTkLabel(self, text=subtitle, font=F(12), text_color=TEXT_SECOND, anchor="w").grid(
                row=1, column=1, sticky="w"
            )


class Card(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(
            master, fg_color=CARD_BG, corner_radius=18, border_width=1,
            border_color=CARD_BORDER, **kwargs
        )


class FilterRow(ctk.CTkFrame):
    """One dynamic filter line: [file] [column] [values to keep] [remove]."""

    def __init__(self, master, app, on_remove):
        super().__init__(master, fg_color=ROW_BG, corner_radius=12)
        self.app = app
        self.on_remove = on_remove
        self.grid_columnconfigure((0, 1, 2), weight=1)

        self.file_menu = ctk.CTkOptionMenu(
            self, values=app.file_labels(), command=self._on_file_change,
            width=160, fg_color=("#E9EBF6", "#262B3D"), button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, text_color=TEXT_PRIMARY, font=F(12),
        )
        self.file_menu.grid(row=0, column=0, padx=(12, 6), pady=12, sticky="ew")

        self.col_menu = ctk.CTkOptionMenu(
            self, values=["—"], width=180, fg_color=("#E9EBF6", "#262B3D"),
            button_color=ACCENT, button_hover_color=ACCENT_HOVER,
            text_color=TEXT_PRIMARY, font=F(12),
        )
        self.col_menu.grid(row=0, column=1, padx=6, pady=12, sticky="ew")

        self.values_entry = ctk.CTkEntry(
            self, placeholder_text="values to KEEP, comma-separated  e.g. BNG, Router, Switch, Controller",
            fg_color=("#FFFFFF", "#11141E"), border_color=CARD_BORDER, text_color=TEXT_PRIMARY, font=F(12),
        )
        self.values_entry.grid(row=0, column=2, padx=6, pady=12, sticky="ew")

        remove_btn = ctk.CTkButton(
            self, text="✕", width=32, height=32, corner_radius=10,
            fg_color="transparent", hover_color=DANGER, text_color=DANGER,
            font=F(14, "bold"), command=lambda: self.on_remove(self),
        )
        remove_btn.grid(row=0, column=3, padx=(6, 12), pady=12)

        self._on_file_change(self.file_menu.get())

    def _on_file_change(self, label):
        cols = self.app.columns_for_label(label)
        self.col_menu.configure(values=cols if cols else ["—"])
        self.col_menu.set(cols[0] if cols else "—")

    def get_spec(self):
        return {
            "file_label": self.file_menu.get(),
            "column": self.col_menu.get(),
            "values": [v.strip() for v in self.values_entry.get().split(",") if v.strip()],
        }


# ----------------------------------------------------------------------------
# Main application
# ----------------------------------------------------------------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("File Compare Studio")
        self.geometry("1080x800")
        self.minsize(900, 650)
        self.configure(fg_color=BG)

        self.path1, self.path2 = None, None
        self.df1, self.df2 = None, None
        self.filter_rows = []
        self.legacy_switches = {}  # filename -> CTkSwitch for "remove 1970" toggle
        self.output_dir = None

        self._build_header()
        self._build_scroll_area()
        self._build_file_card()
        self._build_filter_card()
        self._build_key_card()
        self._build_log_card()

    # ---------------- header ----------------
    def _build_header(self):
        bar = ctk.CTkFrame(self, fg_color="transparent", height=70)
        bar.pack(fill="x", padx=28, pady=(22, 6))

        left = ctk.CTkFrame(bar, fg_color="transparent")
        left.pack(side="left")
        ctk.CTkLabel(left, text="File Compare Studio", font=F(26, "bold"), text_color=TEXT_PRIMARY).pack(anchor="w")
        ctk.CTkLabel(
            left, text="Load two files → filter → compare → get 3 clean reports",
            font=F(13), text_color=TEXT_SECOND,
        ).pack(anchor="w")

        right = ctk.CTkFrame(bar, fg_color="transparent")
        right.pack(side="right")
        self.mode_switch = ctk.CTkSegmentedButton(
            right, values=["Dark", "Light"], command=self._toggle_mode,
            selected_color=ACCENT, selected_hover_color=ACCENT_HOVER, font=F(12, "bold"),
        )
        self.mode_switch.set("Dark")
        self.mode_switch.pack(anchor="e")

    def _toggle_mode(self, value):
        ctk.set_appearance_mode("dark" if value == "Dark" else "light")

    # ---------------- scroll container ----------------
    def _build_scroll_area(self):
        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=28, pady=(6, 22))
        self.scroll.grid_columnconfigure(0, weight=1)

    # ---------------- card 1: load files ----------------
    def _build_file_card(self):
        card = Card(self.scroll)
        card.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        card.grid_columnconfigure(0, weight=1)
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=22, pady=20)
        inner.grid_columnconfigure(0, weight=1)

        StepHeader(inner, 1, "Load your two files", "CSV or XLSX — they are only read, never edited").grid(
            row=0, column=0, sticky="w", pady=(0, 14)
        )

        self.file_pickers = []
        for i in range(2):
            row = ctk.CTkFrame(inner, fg_color="transparent")
            row.grid(row=1 + i, column=0, sticky="ew", pady=6)
            row.grid_columnconfigure(1, weight=1)

            ctk.CTkLabel(row, text=f"File {i + 1}", font=F(12, "bold"), text_color=TEXT_SECOND, width=55).grid(
                row=0, column=0, sticky="w"
            )
            entry = ctk.CTkEntry(
                row, placeholder_text="No file selected…", fg_color=("#FFFFFF", "#11141E"),
                border_color=CARD_BORDER, text_color=TEXT_PRIMARY, font=F(12),
            )
            entry.grid(row=0, column=1, sticky="ew", padx=10)
            browse = ctk.CTkButton(
                row, text="Browse", width=90, corner_radius=10, fg_color=ACCENT,
                hover_color=ACCENT_HOVER, font=F(12, "bold"),
                command=lambda idx=i: self._browse(idx),
            )
            browse.grid(row=0, column=2)
            self.file_pickers.append(entry)

        info_row = ctk.CTkFrame(inner, fg_color="transparent")
        info_row.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        self.load_btn = ctk.CTkButton(
            info_row, text="Load Files", width=140, height=36, corner_radius=10,
            fg_color=TEAL, hover_color=TEAL_HOVER, text_color="#06281F", font=F(13, "bold"),
            command=self._load_files,
        )
        self.load_btn.pack(side="left")
        self.load_status = ctk.CTkLabel(info_row, text="", font=F(12), text_color=TEXT_SECOND)
        self.load_status.pack(side="left", padx=14)

    def _browse(self, idx):
        path = filedialog.askopenfilename(
            title=f"Select File {idx + 1}",
            filetypes=[("CSV / Excel files", "*.csv *.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.file_pickers[idx].delete(0, "end")
            self.file_pickers[idx].insert(0, path)

    def _load_files(self):
        p1, p2 = self.file_pickers[0].get().strip(), self.file_pickers[1].get().strip()
        if not p1 or not p2:
            messagebox.showwarning("Missing files", "Please choose both File 1 and File 2.")
            return
        try:
            self.df1 = read_file(p1)
            self.df2 = read_file(p2)
            self.path1, self.path2 = p1, p2
        except Exception as e:
            messagebox.showerror("Couldn't load files", str(e))
            return

        self.load_status.configure(
            text=(
                f"✓ {os.path.basename(p1)} ({len(self.df1)} rows, {len(self.df1.columns)} cols)   •   "
                f"{os.path.basename(p2)} ({len(self.df2)} rows, {len(self.df2.columns)} cols)"
            ),
            text_color=TEAL,
        )
        self.output_dir = os.path.dirname(os.path.abspath(p1))
        self._refresh_filter_rows()
        self._refresh_legacy_switches()
        self._refresh_key_menus()
        self._log(f"Loaded '{os.path.basename(p1)}' and '{os.path.basename(p2)}'.")

    # ---------------- card 2: filters ----------------
    def _build_filter_card(self):
        self.filter_card = Card(self.scroll)
        self.filter_card.grid(row=1, column=0, sticky="ew", pady=(0, 18))
        self.filter_inner = ctk.CTkFrame(self.filter_card, fg_color="transparent")
        self.filter_inner.pack(fill="x", padx=22, pady=20)
        self.filter_inner.grid_columnconfigure(0, weight=1)

        StepHeader(
            self.filter_inner, 2, "Filters (optional)",
            "Keep only rows whose column value matches what you list — applied before comparing",
        ).grid(row=0, column=0, sticky="w", pady=(0, 14))

        self.filter_rows_frame = ctk.CTkFrame(self.filter_inner, fg_color="transparent")
        self.filter_rows_frame.grid(row=1, column=0, sticky="ew")
        self.filter_rows_frame.grid_columnconfigure(0, weight=1)

        self.add_filter_btn = ctk.CTkButton(
            self.filter_inner, text="+ Add Filter", width=130, height=32, corner_radius=10,
            fg_color="transparent", border_width=1, border_color=ACCENT, text_color=ACCENT,
            hover_color=("#EFEBFF", "#1E1A33"), font=F(12, "bold"), command=self._add_filter_row,
        )
        self.add_filter_btn.grid(row=2, column=0, sticky="w", pady=(12, 4))
        self.add_filter_btn.configure(state="disabled")

        self.legacy_frame = ctk.CTkFrame(self.filter_inner, fg_color="transparent")
        self.legacy_frame.grid(row=3, column=0, sticky="ew", pady=(14, 0))

    def file_labels(self):
        labels = []
        if self.path1:
            labels.append(os.path.basename(self.path1) + "  (File 1)")
        if self.path2:
            labels.append(os.path.basename(self.path2) + "  (File 2)")
        return labels or ["—"]

    def columns_for_label(self, label):
        if self.path1 and label.startswith(os.path.basename(self.path1)):
            return list(self.df1.columns)
        if self.path2 and label.startswith(os.path.basename(self.path2)):
            return list(self.df2.columns)
        return []

    def df_for_label(self, label):
        if self.path1 and label.startswith(os.path.basename(self.path1)):
            return self.df1, 0
        if self.path2 and label.startswith(os.path.basename(self.path2)):
            return self.df2, 1
        return None, None

    def _refresh_filter_rows(self):
        for r in self.filter_rows:
            r.destroy()
        self.filter_rows = []
        self.add_filter_btn.configure(state="normal")

    def _add_filter_row(self):
        row = FilterRow(self.filter_rows_frame, self, self._remove_filter_row)
        row.pack(fill="x", pady=6)
        self.filter_rows.append(row)

    def _remove_filter_row(self, row):
        row.destroy()
        self.filter_rows.remove(row)

    def _refresh_legacy_switches(self):
        for w in self.legacy_frame.winfo_children():
            w.destroy()
        self.legacy_switches = {}
        for df, path in ((self.df1, self.path1), (self.df2, self.path2)):
            if df is not None and "ActivationDate" in df.columns:
                name = os.path.basename(path)
                sw = ctk.CTkSwitch(
                    self.legacy_frame,
                    text=f"Remove '1970' ActivationDate rows  —  {name}",
                    progress_color=ACCENT, font=F(12), text_color=TEXT_PRIMARY,
                )
                sw.pack(anchor="w", pady=4)
                self.legacy_switches[path] = sw

    # ---------------- card 3: keys + run ----------------
    def _build_key_card(self):
        card = Card(self.scroll)
        card.grid(row=2, column=0, sticky="ew", pady=(0, 18))
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=22, pady=20)
        inner.grid_columnconfigure((0, 1), weight=1)

        StepHeader(
            inner, 3, "Comparison key",
            "Pick the column that should match between the two files (e.g. an IP address column)",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 14))

        ctk.CTkLabel(inner, text="Key column — File 1", font=F(12, "bold"), text_color=TEXT_SECOND).grid(
            row=1, column=0, sticky="w"
        )
        ctk.CTkLabel(inner, text="Key column — File 2", font=F(12, "bold"), text_color=TEXT_SECOND).grid(
            row=1, column=1, sticky="w"
        )

        self.key1_menu = ctk.CTkOptionMenu(
            inner, values=["—"], fg_color=("#E9EBF6", "#262B3D"), button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, text_color=TEXT_PRIMARY, font=F(12),
        )
        self.key1_menu.grid(row=2, column=0, sticky="ew", padx=(0, 10), pady=(4, 0))
        self.key2_menu = ctk.CTkOptionMenu(
            inner, values=["—"], fg_color=("#E9EBF6", "#262B3D"), button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, text_color=TEXT_PRIMARY, font=F(12),
        )
        self.key2_menu.grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=(4, 0))

        self.run_btn = ctk.CTkButton(
            inner, text="▶  Run Comparison", height=42, corner_radius=12, fg_color=ACCENT,
            hover_color=ACCENT_HOVER, font=F(14, "bold"), command=self._run,
        )
        self.run_btn.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(20, 0))

    def _refresh_key_menus(self):
        cols1 = list(self.df1.columns) if self.df1 is not None else ["—"]
        cols2 = list(self.df2.columns) if self.df2 is not None else ["—"]
        self.key1_menu.configure(values=cols1)
        self.key1_menu.set(guess_ip_column(cols1))
        self.key2_menu.configure(values=cols2)
        self.key2_menu.set(guess_ip_column(cols2))

    # ---------------- card 4: log / results ----------------
    def _build_log_card(self):
        card = Card(self.scroll)
        card.grid(row=3, column=0, sticky="ew")
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=22, pady=20)
        inner.grid_columnconfigure(0, weight=1)

        StepHeader(inner, 4, "Results").grid(row=0, column=0, sticky="w", pady=(0, 10))

        self.log_box = ctk.CTkTextbox(
            inner, height=160, fg_color=("#FFFFFF", "#11141E"), text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(family="Consolas", size=12), corner_radius=10,
        )
        self.log_box.grid(row=1, column=0, sticky="ew")
        self.log_box.configure(state="disabled")

        self.open_folder_btn = ctk.CTkButton(
            inner, text="Open Output Folder", width=170, height=34, corner_radius=10,
            fg_color="transparent", border_width=1, border_color=TEAL, text_color=TEAL,
            hover_color=("#E6FBF6", "#102420"), font=F(12, "bold"), command=self._open_output_folder,
        )
        self.open_folder_btn.grid(row=2, column=0, sticky="w", pady=(12, 0))
        self.open_folder_btn.configure(state="disabled")

    def _log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.configure(state="disabled")
        self.log_box.see("end")

    def _open_output_folder(self):
        if not self.output_dir:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.output_dir)
            elif sys.platform == "darwin":
                subprocess.run(["open", self.output_dir])
            else:
                subprocess.run(["xdg-open", self.output_dir])
        except Exception as e:
            messagebox.showinfo("Output folder", f"Files were saved to:\n{self.output_dir}\n\n({e})")

    # ---------------- run logic ----------------
    def _run(self):
        if self.df1 is None or self.df2 is None:
            messagebox.showwarning("No files loaded", "Please load both files first.")
            return
        try:
            d1, d2 = self.df1.copy(), self.df2.copy()

            # generic keep-filters
            for row in self.filter_rows:
                spec = row.get_spec()
                df, which = self.df_for_label(spec["file_label"])
                if df is None or spec["column"] not in df.columns or not spec["values"]:
                    continue
                target = d1 if which == 0 else d2
                before = len(target)
                vals_lower = [v.lower() for v in spec["values"]]
                mask = target[spec["column"]].astype(str).str.strip().str.lower().isin(vals_lower)
                filtered = target[mask].reset_index(drop=True)
                if which == 0:
                    d1 = filtered
                else:
                    d2 = filtered
                self._log(f"Filter: {spec['file_label']} · {spec['column']} kept {spec['values']} → {before} → {len(filtered)} rows")

            # 1970 ActivationDate switches
            for path, sw in self.legacy_switches.items():
                if sw.get() == 1:
                    target = d1 if path == self.path1 else d2
                    before = len(target)
                    mask = ~target["ActivationDate"].astype(str).str.strip().str.startswith("1970")
                    filtered = target[mask].reset_index(drop=True)
                    if path == self.path1:
                        d1 = filtered
                    else:
                        d2 = filtered
                    self._log(f"Filter: removed 1970 ActivationDate rows in {os.path.basename(path)} → {before} → {len(filtered)} rows")

            key1, key2 = self.key1_menu.get(), self.key2_menu.get()
            if key1 not in d1.columns or key2 not in d2.columns:
                messagebox.showerror("Invalid key column", "Please pick valid key columns for both files.")
                return

            base1 = os.path.splitext(os.path.basename(self.path1))[0]
            base2 = os.path.splitext(os.path.basename(self.path2))[0]

            only_in_1, only_in_2, common = run_comparison(d1, d2, key1, key2, base1, base2)

            out1 = os.path.join(self.output_dir, f"only_in_{base1}.xlsx")
            out2 = os.path.join(self.output_dir, f"only_in_{base2}.xlsx")
            out3 = os.path.join(self.output_dir, f"common_{base1}_{base2}.xlsx")
            only_in_1.to_excel(out1, index=False)
            only_in_2.to_excel(out2, index=False)
            common.to_excel(out3, index=False)

            self._log("—" * 50)
            self._log(f"Only in {base1}: {len(only_in_1)} rows → {out1}")
            self._log(f"Only in {base2}: {len(only_in_2)} rows → {out2}")
            self._log(f"Matched in both: {len(common)} rows → {out3}")
            self._log("Done. Original files were not modified.")
            self.open_folder_btn.configure(state="normal")

        except Exception:
            messagebox.showerror("Error while running comparison", traceback.format_exc())


if __name__ == "__main__":
    app = App()
    app.mainloop()
