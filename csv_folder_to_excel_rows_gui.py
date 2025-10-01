import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

#Convertion of columns to simple index - JJ
def col_letter_to_index(letter: str) -> int:
    """Convert Excel-style column letter (e.g., 'A', 'Q', 'BL') to 0-based index."""
    letter = letter.strip().upper()
    if not letter or any(not ('A' <= ch <= 'Z') for ch in letter):
        raise ValueError(f"Invalid column letter: {letter}")
    num = 0
    for ch in letter:
        num = num * 26 + (ord(ch) - ord('A') + 1)
    return num - 1  # 0-based indexing -JJ

#Column Q is the current data entry point within the LIMS Summary - JJ
#Column can be different and be updated here if different index - JJ
def read_column_q(csv_path: Path, has_header: bool, encoding="utf-8-sig", delimiter=","):
    """
    Read Column Q (17th column, 0-based index 16) from a CSV as a list of strings.
    If has_header=True, pandas uses the first row as header and returns only data rows.
    If has_header=False, all rows are treated as data.
    """
    q_idx = col_letter_to_index("Q")  # Q maps to index 16 -JJ 
    try:
        if has_header:
            df = pd.read_csv(csv_path, dtype=str, sep=delimiter, encoding=encoding,
                             low_memory=False, header=0, usecols=[q_idx])
        else:
            df = pd.read_csv(csv_path, dtype=str, sep=delimiter, encoding=encoding,
                             low_memory=False, header=None, usecols=[q_idx])
    except Exception as e: #failure exception -JJ
        return None, f"read_csv failed for {csv_path.name}: {e}"

    if df.shape[1] != 1:
        return None, f"{csv_path.name}: Column Q not present"

    series = df.iloc[:, 0]
    values = series.fillna("").astype(str).str.strip().tolist()
    return values, None

#CSV delimiter chosen of ","
def process_folder_to_excel(folder: Path, output_xlsx: Path, has_header: bool,
                            encoding="utf-8-sig", delimiter=",") -> bool:
    csv_files = sorted(folder.glob("*.csv"))
    if not csv_files:
        messagebox.showerror("No CSVs", f"No .csv files found in:\n{folder}")
        return False

    columns_data = []   # list of lists; each list is one output column: [filename, Q1, Q2, ...]
    skipped = []

    #Error addins - JJ
    for csv_path in csv_files:
        values, err = read_column_q(csv_path, has_header=has_header, encoding=encoding, delimiter=delimiter)
        if err is not None:
            skipped.append(err)
            continue

        col_list = [csv_path.stem] + values  # top cell: filename (no extension), then column Q values
        columns_data.append(col_list)

    #Error notification addins - JJ
    if not columns_data:
        messagebox.showerror("Nothing to write", "All files were skipped.\n\n" + "\n".join(skipped[:10]))
        return False

    # Normalize lengths (pad shorter columns so we can write a rectangle)
    max_len = max(len(col) for col in columns_data)
    for col in columns_data:
        if len(col) < max_len:
            col.extend([""] * (max_len - len(col)))

    # Build a DataFrame where each list is a COLUMN, and write with NO headers
    df = pd.DataFrame({i: col for i, col in enumerate(columns_data)})

    try:
        df.to_excel(output_xlsx, index=False, header=False, engine="xlsxwriter")
    except Exception as e:
        messagebox.showerror("Write error", f"Failed to write Excel:\n{e}")
        return False

    # Summary
    summary = [f"Processed: {len(columns_data)} files",
               f"Output: {output_xlsx}"]
    if skipped:
        summary.append("\nSkipped:")
        summary.extend(" - " + s for s in skipped[:10])
        if len(skipped) > 10:
            summary.append(f" ...and {len(skipped)-10} more")

    messagebox.showinfo("Done", "\n".join(summary))
    return True

def main():
    root = tk.Tk()
    root.withdraw()
#GUI assistance from Chatgpt - Checked for content JJ
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
    output_xlsx = folder / "Combined_Output_Data.xlsx"

    process_folder_to_excel(folder, output_xlsx, has_header=has_header)

if __name__ == "__main__":
    main()
