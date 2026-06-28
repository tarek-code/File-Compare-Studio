"""
Dynamic File Comparison Tool
-----------------------------
Compares two data files (CSV or XLSX) on a chosen key column from each file,
optionally applies filters before comparing, and produces three output files:
  1. Rows that exist only in File 1 (not in File 2)
  2. Rows that exist only in File 2 (not in File 1)
  3. Rows that exist in BOTH files (matched)

The original source files are only READ - they are never modified.

Requirements:
    pip install pandas openpyxl

Usage:
    python compare_files.py
    (then follow the interactive prompts)

Example for this project:
    File #1 -> noms.xlsx        (key column: IpAddress)
    File #2 -> NA_190.csv       (key column: Device IP)
    Filter  -> noms.xlsx, column "HWType", keep values: BNG, Router, Switch, Controller
    Special filter -> remove ActivationDate rows starting with "1970" (asked automatically)
"""

import os
import sys
import pandas as pd


def read_file(path: str) -> pd.DataFrame:
    """Read a CSV or Excel file into a DataFrame, treating every cell as text."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
            try:
                return pd.read_csv(path, dtype=str, keep_default_na=False, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise ValueError(f"Could not read '{path}' with common encodings.")
    elif ext in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str)
    else:
        raise ValueError(f"Unsupported file extension '{ext}'. Only .csv, .xlsx, .xls are supported.")


def ask_yes_no(prompt: str) -> bool:
    while True:
        ans = input(prompt + " (y/n): ").strip().lower()
        if ans in ("y", "yes"):
            return True
        if ans in ("n", "no"):
            return False
        print("Please answer 'y' or 'n'.")


def ask_int(prompt: str) -> int:
    while True:
        try:
            return int(input(prompt).strip())
        except ValueError:
            print("Please enter a valid whole number.")


def choose_file_index(files):
    for i, f in enumerate(files):
        print(f"  {i + 1}. {f}")
    while True:
        choice = ask_int("Enter the file number: ")
        if 1 <= choice <= len(files):
            return choice - 1
        print(f"Please enter a number between 1 and {len(files)}.")


def main():
    print("=" * 60)
    print("Dynamic File Comparison Tool")
    print("=" * 60)
    print("\nThis tool reads two files (CSV or XLSX), lets you filter")
    print("them, and outputs:")
    print("  1) Rows only in File 1")
    print("  2) Rows only in File 2")
    print("  3) Rows matched in both files")
    print("\nThe original files are only READ, never modified.\n")

    # ---------- Step 1: how many / which files ----------
    while True:
        n = ask_int("How many files do you want to load? (this tool compares exactly 2): ")
        if n == 2:
            break
        print("This comparison tool needs exactly 2 files. Please enter 2.")

    files = []
    for i in range(n):
        while True:
            fname = input(f"Enter the name (with extension) of file #{i + 1}: ").strip().strip('"')
            if os.path.exists(fname):
                files.append(fname)
                break
            print(f"  -> File '{fname}' was not found. Check the name/path and try again.")

    dataframes = []
    for f in files:
        df = read_file(f)
        dataframes.append(df)
        print(f"\nLoaded '{f}': {len(df)} rows, {len(df.columns)} columns")
        print(f"Columns: {list(df.columns)}")

    # ---------- Step 2: optional generic filters ----------
    if ask_yes_no("\nDo you want to apply a filter on any of the loaded files?"):
        keep_filtering = True
        while keep_filtering:
            print("\nWhich file do you want to filter?")
            f_idx = choose_file_index(files)
            df = dataframes[f_idx]
            print(f"Columns in '{files[f_idx]}': {list(df.columns)}")
            col = input("Which column do you want to filter on?: ").strip()
            if col not in df.columns:
                print(f"  -> Column '{col}' not found in '{files[f_idx]}'. Filter skipped.")
            else:
                raw_values = input(
                    f"Enter the value(s) to KEEP in '{col}', comma-separated "
                    f"(e.g. BNG,Router,Switch,Controller): "
                ).strip()
                values = [v.strip().lower() for v in raw_values.split(",") if v.strip()]
                before = len(df)
                mask = df[col].astype(str).str.strip().str.lower().isin(values)
                dataframes[f_idx] = df[mask].reset_index(drop=True)
                print(f"  -> Filtered '{files[f_idx]}' on '{col}': {before} -> {len(dataframes[f_idx])} rows kept")
            keep_filtering = ask_yes_no("Do you want to apply another filter?")

    # ---------- Step 3: special ActivationDate / 1970 filter (always asked if column exists) ----------
    for idx, df in enumerate(dataframes):
        if "ActivationDate" in df.columns:
            if ask_yes_no(
                f"\nFile '{files[idx]}' has an 'ActivationDate' column. "
                f"Do you want to remove rows where ActivationDate starts with '1970'?"
            ):
                before = len(df)
                mask = ~df["ActivationDate"].astype(str).str.strip().str.startswith("1970")
                dataframes[idx] = df[mask].reset_index(drop=True)
                print(f"  -> Removed 1970 entries from '{files[idx]}': {before} -> {len(dataframes[idx])} rows kept")

    # ---------- Step 4: choose the comparison key column for each file ----------
    print("\nNow choose the column to compare on (the matching key) for each file.")
    key_columns = []
    for idx, f in enumerate(files):
        df = dataframes[idx]
        print(f"\nColumns in '{f}': {list(df.columns)}")
        while True:
            col = input(f"Comparison key column for '{f}': ").strip()
            if col in df.columns:
                key_columns.append(col)
                break
            print(f"  -> Column '{col}' not found. Please pick one from the list above.")

    df1, df2 = dataframes[0].copy(), dataframes[1].copy()
    key1, key2 = key_columns[0], key_columns[1]

    # Normalized helper columns for robust matching (whitespace/case-insensitive)
    df1["_match_key"] = df1[key1].astype(str).str.strip().str.lower()
    df2["_match_key"] = df2[key2].astype(str).str.strip().str.lower()

    set1 = set(df1["_match_key"])
    set2 = set(df2["_match_key"])

    only_keys_1 = set1 - set2
    only_keys_2 = set2 - set1
    common_keys = set1 & set2

    only_in_1 = df1[df1["_match_key"].isin(only_keys_1)].drop(columns=["_match_key"]).reset_index(drop=True)
    only_in_2 = df2[df2["_match_key"].isin(only_keys_2)].drop(columns=["_match_key"]).reset_index(drop=True)

    base1 = os.path.splitext(os.path.basename(files[0]))[0]
    base2 = os.path.splitext(os.path.basename(files[1]))[0]

    matched_1 = df1[df1["_match_key"].isin(common_keys)].reset_index(drop=True)
    matched_2 = df2[df2["_match_key"].isin(common_keys)].reset_index(drop=True)

    common = pd.merge(
        matched_1, matched_2,
        on="_match_key",
        suffixes=(f"_{base1}", f"_{base2}"),
    ).drop(columns=["_match_key"])

    # ---------- Step 5: write outputs (never touching the original files) ----------
    out_dir = os.path.dirname(os.path.abspath(files[0])) or "."
    out1 = os.path.join(out_dir, f"only_in_{base1}.xlsx")
    out2 = os.path.join(out_dir, f"only_in_{base2}.xlsx")
    out3 = os.path.join(out_dir, f"common_{base1}_{base2}.xlsx")

    only_in_1.to_excel(out1, index=False)
    only_in_2.to_excel(out2, index=False)
    common.to_excel(out3, index=False)

    print("\n" + "=" * 60)
    print("Done! Original files were not modified. Output files created:")
    print(f"  1) Only in {files[0]}: {out1}  ({len(only_in_1)} rows)")
    print(f"  2) Only in {files[1]}: {out2}  ({len(only_in_2)} rows)")
    print(f"  3) Matched in both:    {out3}  ({len(common)} rows)")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelled by user.")
        sys.exit(1)
