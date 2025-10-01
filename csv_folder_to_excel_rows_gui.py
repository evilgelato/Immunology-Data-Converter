import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import re
import numpy as np

# ---------- Configurable reorder scheme ----------
TARGET_ROW_ORDER = [
    "Albumin",
    "Alpha 1-Antitrypsin",
    "Apolipoprotein A1",
    "Apolipoprotein B",
    "Ceruloplasmin",
    "Complement C3",
    "Complement C4",
    "IgA",
    "IgM",
    "Transferrin",
    "AAG",
    "AMG",
    "Haptoglobin",
    "Prealbumin",
    "IgE",
    "Anti-CCP",
    "Ferritin",
]

# Normalize: lowercase, drop punctuation, collapse spaces
_norm_non_alnum = re.compile(r"[^a-z0-9]+")
def normalize_label(s: str) -> str:
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = s.strip().lower()
    s = _norm_non_alnum.sub(" ", s)
    s = " ".join(s.split())
    return s

def col_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    num = 0
    for ch in letter:
        if not ('A' <= ch <= 'Z'):
            raise ValueError(f"Invalid column letter: {letter}")
        num = num * 26 + (ord(ch) - ord('A') + 1)
    return num - 1  # 0-based

def read_column_by_index(csv_path: Path, col_idx: int, has_header: bool,
                         encoding="utf-8-sig", delimiter=","):
    """Read one column as a list of strings (header row skipped if has_header=True)."""
    try:
        if has_header:
            df = pd.read_csv(csv_path, dtype=str, sep=delimiter, encoding=encoding,
                             low_memory=False, header=0, usecols=[col_idx])
        else:
            df = pd.read_csv(csv_path, dtype=str, sep=delimiter, encoding=encoding,
                             low_memory=False, header=None, usecols=[col_idx])
    except Exception as e:
        return None, f"{csv_path.name}: read_csv failed: {e}"
    if df.shape[1] != 1:
        return None, f"{csv_path.name}: column index {col_idx} not present"
    series = df.iloc[:, 0]
    return series.fillna("").astype(str).str.strip().tolist(), None

def read_p_and_q_filtered(csv_path: Path, has_header: bool,
                          encoding="utf-8-sig", delimiter=","):
    """Return Q values where corresponding P does NOT contain 'comment'."""
    p_idx = col_letter_to_index("P")  # 15
    q_idx = col_letter_to_index("Q")  # 16
    try:
        if has_header:
            df = pd.read_csv(csv_path, dtype=str, sep=delimiter, encoding=encoding,
                             low_memory=False, header=0, usecols=[p_idx, q_idx])
        else:
            df = pd.read_csv(csv_path, dtype=str, sep=delimiter, encoding=encoding,
                             low_memory=False, header=None, usecols=[p_idx, q_idx])
    except Exception as e:
        return None, f"{csv_path.name}: read_csv failed: {e}"
    if df.shape[1] != 2:
        return None, f"{csv_path.name}: columns P and/or Q not present"
    df.columns = ["P", "Q"]
    df["P"] = df["P"].fillna("").astype(str).str.strip()
    df["Q"] = df["Q"].fillna("").astype(str).str.strip()
    mask = ~df["P"].str.contains("comment", case=False, na=False)
    kept_q = df.loc[mask, "Q"].tolist()
    return kept_q, None

def every_third(values, offset=2):
    """Return every 3rd item starting with index=offset (default 2 → 3rd, 6th, 9th…)."""
    return values[offset::3] if values else []

number_pat = re.compile(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?")
def clean_to_float(val):
    """Strip text, keep only numeric part, convert to float; else NaN."""
    if not isinstance(val, str):
        return val
    m = number_pat.search(val)
    return float(m.group()) if m else np.nan

def process_folder(folder: Path, output_xlsx: Path, has_header: bool,
                   encoding="utf-8-sig", delimiter=",") -> bool:
    csv_files = sorted(folder.glob("*.csv"))
    if not csv_files:
        messagebox.showerror("No CSVs", f"No .csv files found in:\n{folder}")
        return False

    columns_data = []
    skipped = []
    kept_counts = []

    # --- First column: Column P from the FIRST CSV ---
    first_csv = csv_files[0]
    p_idx = col_letter_to_index("P")
    p_vals, err = read_column_by_index(first_csv, p_idx, has_header, encoding, delimiter)
    if err:
        skipped.append(err)
        p_col = [first_csv.stem]
    else:
        p_vals_third = every_third(p_vals)  # 3rd, 6th, 9th...
        p_col = [first_csv.stem] + p_vals_third
        kept_counts.append((f"{first_csv.name} (P)", len(p_vals_third)))
    columns_data.append(p_col)

    # --- Each CSV contributes Q (filtered by P not containing 'comment'), every 3rd row ---
    for csv_path in csv_files:
        q_vals, err2 = read_p_and_q_filtered(csv_path, has_header, encoding, delimiter)
        if err2:
            skipped.append(err2)
            continue
        q_vals_third = every_third(q_vals)
        col_list = [csv_path.stem] + q_vals_third
        columns_data.append(col_list)
        kept_counts.append((f"{csv_path.name} (Q filtered)", len(q_vals_third)))

    if not columns_data:
        messagebox.showerror("Nothing to write", "No data extracted.")
        return False

    # Pad columns to same length
    max_len = max(len(col) for col in columns_data)
    for col in columns_data:
        if len(col) < max_len:
            col.extend([""] * (max_len - len(col)))

    # Build DataFrame where each list is a column
    df = pd.DataFrame({i: col for i, col in enumerate(columns_data)})

    # --- Clean all columns except the first (entire column, including row 0) ---
    if df.shape[1] > 1:
        for col in df.columns[1:]:
            df[col] = df[col].map(clean_to_float)

    # --- REORDER ROWS using PARTIAL, case-insensitive matches on first column ---
    # Row 0 is filenames -> keep at top. Rows 1.. are labels in first column.
    # Precompute normalized first-column values
    row_norm = {idx: normalize_label(val) for idx, val in df.iloc[1:, 0].items()}

    # Find matching row for each target, preferring first unused row that matches
    used_rows = set()
    ordered_rows = [0]  # keep filenames row at top
    missing = []

    for label in TARGET_ROW_ORDER:
        tnorm = normalize_label(label)
        found_idx = None
        # First pass: row contains target OR target contains row
        for idx in range(1, len(df)):
            if idx in used_rows:
                continue
            rnorm = row_norm.get(idx, "")
            if not rnorm:
                continue
            if (tnorm in rnorm) or (rnorm in tnorm and len(rnorm) >= 2):
                found_idx = idx
                break
        if found_idx is not None:
            ordered_rows.append(found_idx)
            used_rows.add(found_idx)
        else:
            missing.append(label)

    # Append any remaining rows (not matched) in original order
    for idx in range(1, len(df)):
        if idx not in used_rows:
            ordered_rows.append(idx)

    df = df.iloc[ordered_rows].reset_index(drop=True)

    # Write Excel
    try:
        df.to_excel(output_xlsx, index=False, header=False, engine="xlsxwriter")
    except Exception as e:
        messagebox.showerror("Write error", f"Failed to write Excel:\n{e}")
        return False

    # Summary
    summary = [
        f"Processed files: {len(csv_files)}",
        f"Output: {output_xlsx}",
        "",
        "Kept counts (after filtering & every 3rd row):"
    ]
    summary.extend([f" - {name}: {count}" for name, count in kept_counts[:12]])
    if len(kept_counts) > 12:
        summary.append(f"...and {len(kept_counts) - 12} more")

    if missing:
        summary.append("\nNot found (no partial match in first column):")
        summary.extend(f" - {m}" for m in missing)

    if skipped:
        summary.append("\nSkipped:")
        summary.extend(" - " + s for s in skipped[:10])
        if len(skipped) > 10:
            summary.append(f"...and {len(skipped) - 10} more")

    messagebox.showinfo("Done", "\n".join(summary))
    return True

def main():
    root = tk.Tk()
    root.withdraw()

    has_header = messagebox.askyesno(
        "Header Row",
        "Do your CSV files have a header row?\n\n"
        "Yes = first row is header (not data)\n"
        "No  = all rows are data"
    )

    folder_str = filedialog.askdirectory(title="Select folder containing CSV files")
    if not folder_str:
        return

    folder = Path(folder_str)
    output_xlsx = folder / "combined_output_reordered.xlsx"
    process_folder(folder, output_xlsx, has_header)

if __name__ == "__main__":
    main()
