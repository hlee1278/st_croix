"""
KESTREL SPREADSHEET COMBINER
==============================
PURPOSE:
    Reads all Kestrel CSV files from a single folder, adds a "kestrel_name"
    column to each (pulled from the filename, e.g. "0618_TEST2_KESTRELFIRE_1.csv"
    → "FIRE"), and combines everything into one Excel file saved in the same folder.

    Only files with "KESTREL" in the name are used. All others are skipped.

HOW TO USE THIS SCRIPT:
    1. Fill in Section 1 below (the only section you need to edit)
    2. Run the script:   python combine_kestrel.py
    3. The combined output file will be saved inside the same folder as your CSVs,
       automatically named:  COMBINED_KESTREL_<folder name>.xlsx

DEPENDENCIES — install once if not already installed:
    pip install pandas openpyxl
"""

import os
import re
import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# =============================================================================
# SECTION 1: CONFIGURATION — THIS IS THE ONLY SECTION YOU NEED TO EDIT
# =============================================================================

INPUT_FOLDER = r"C:\Users\dresd\Downloads\0618_R1000"
                            # Path to the folder containing the Kestrel CSV files
                            # Only files with "KESTREL" in the filename will be read
                            # All other CSV files in the folder will be skipped

SHEET_NAME = "Combined_Kestrel"
                            # Name of the sheet in the output Excel file
                            # Keep it short — Excel allows a maximum of 31 characters


# =============================================================================
# END OF USER INPUT — DO NOT EDIT BELOW THIS LINE UNLESS CUSTOMIZING
# =============================================================================


# =============================================================================
# SECTION 2: VALIDATE INPUTS AND BUILD OUTPUT FILE PATH
# =============================================================================

if not os.path.isdir(INPUT_FOLDER):
    print(f"ERROR: Input folder not found: {INPUT_FOLDER}")
    sys.exit(1)

if len(SHEET_NAME) > 31:
    print(f"ERROR: SHEET_NAME is too long ({len(SHEET_NAME)} characters). Excel allows a maximum of 31.")
    sys.exit(1)

# Output file is saved inside the same folder, named after the folder itself
# Example: folder "0618_R1000" → output file "COMBINED_KESTREL_0618_R1000.xlsx"
folder_name = os.path.basename(os.path.normpath(INPUT_FOLDER))
OUTPUT_FILE = os.path.join(INPUT_FOLDER, f"COMBINED_KESTREL_{folder_name}.xlsx")


# =============================================================================
# SECTION 3: FIND ALL CSV FILES, KEEPING ONLY KESTREL FILES
# =============================================================================

all_files_in_folder = sorted([
    f for f in os.listdir(INPUT_FOLDER)
    if f.lower().endswith(".csv")
])

kestrel_files = []
skipped       = []

for f in all_files_in_folder:
    if re.search(r"kestrel", f, re.IGNORECASE):
        kestrel_files.append(os.path.join(INPUT_FOLDER, f))
    else:
        skipped.append(f)

if skipped:
    print(f"=== Skipping {len(skipped)} non-Kestrel file(s) ===")
    for f in skipped:
        print(f"  SKIPPED: {f}")
    print()

if not kestrel_files:
    print(f"ERROR: No Kestrel CSV files found in folder: {INPUT_FOLDER}")
    print("       Make sure filenames contain 'KESTREL' (e.g. 0618_TEST2_KESTRELFIRE_1.csv)")
    sys.exit(1)

print(f"=== Found {len(kestrel_files)} Kestrel CSV file(s) to combine ===")
for f in kestrel_files:
    print(f"  {os.path.basename(f)}")
print()


# =============================================================================
# SECTION 4: EXTRACT KESTREL NAME FROM FILENAME
# =============================================================================
# The kestrel name is the word that comes immediately after "KESTREL" in the
# filename (letters only, up to the next underscore or number).
#
# Examples:
#   "0618_TEST2_KESTRELFIRE_1.csv"  → kestrel_name = "FIRE"
#   "0618_TEST2_KESTRELWIND_2.csv"  → kestrel_name = "WIND"

def extract_kestrel_name(filepath):
    """
    Returns the name that follows 'KESTREL' in the filename (uppercase).
    Returns None if no such name can be found.
    """
    filename = os.path.basename(filepath)
    # Match "KESTREL" followed immediately by one or more letters
    match = re.search(r"KESTREL([A-Za-z]+)", filename, re.IGNORECASE)
    if match:
        return match.group(1).upper()
    return None


# =============================================================================
# SECTION 5: READ ALL KESTREL FILES, ADD kestrel_name COLUMN, VERIFY COLUMNS MATCH
# =============================================================================

print("Reading Kestrel files...")

frames        = []
expected_cols = None   # Set from the first file; all others must match

for filepath in kestrel_files:
    filename = os.path.basename(filepath)

    # --- Extract Kestrel name ---
    kestrel_name = extract_kestrel_name(filepath)
    if kestrel_name is None:
        print(f"  WARNING: Could not find a name after 'KESTREL' in '{filename}' — skipping.")
        print("           Expected format: ...KESTREL<NAME>... (e.g. KESTRELFIRE)")
        continue

    # --- Read the CSV ---
    try:
        df = pd.read_csv(filepath, dtype=str)
    except Exception as e:
        print(f"  ERROR reading '{filename}': {e}")
        print("  Skipping this file.")
        continue

    if df.empty:
        print(f"  WARNING: '{filename}' is empty — skipping.")
        continue

    # --- Check that columns match the first file ---
    if expected_cols is None:
        expected_cols = list(df.columns)
        print(f"  Columns detected (from first file): {expected_cols}")
        print()

    if list(df.columns) != expected_cols:
        print(f"  ERROR: '{filename}' has different column headers than the first file.")
        print(f"    Expected : {expected_cols}")
        print(f"    Found    : {list(df.columns)}")
        print("  Skipping this file — fix the column headers and re-run.")
        continue

    # --- Add kestrel_name as the first column ---
    df.insert(0, "kestrel_name", kestrel_name)

    print(f"  {filename}: Kestrel '{kestrel_name}', {len(df)} row(s)")
    frames.append(df)

print()

if not frames:
    print("ERROR: No Kestrel files could be read successfully. Nothing to combine.")
    sys.exit(1)


# =============================================================================
# SECTION 6: COMBINE ALL DATA INTO ONE TABLE
# =============================================================================

combined = pd.concat(frames, ignore_index=True)

print(f"=== Combined total: {len(combined)} rows across {len(frames)} file(s) ===")
print()


# =============================================================================
# SECTION 7: SAVE TO EXCEL WITH FORMATTING
# =============================================================================

print(f"Saving combined file to: {OUTPUT_FILE}")

combined.to_excel(OUTPUT_FILE, sheet_name=SHEET_NAME, index=False)

# --- Apply formatting with openpyxl ---
wb = load_workbook(OUTPUT_FILE)
ws = wb[SHEET_NAME]

header_font   = Font(name="Arial", bold=True)
header_fill   = PatternFill("solid", fgColor="BDD7EE")   # Light blue
header_align  = Alignment(horizontal="center")
header_border = Border(bottom=Side(style="thin"))

for cell in ws[1]:   # Row 1 = headers
    cell.font      = header_font
    cell.fill      = header_fill
    cell.alignment = header_align
    cell.border    = header_border

# Apply Arial font to all data rows
data_font = Font(name="Arial")
for row in ws.iter_rows(min_row=2):
    for cell in row:
        cell.font = data_font

ws.freeze_panes = "A2"   # Freeze header row so it stays visible when scrolling

# Auto-size columns to fit content
for col in ws.columns:
    max_len = max(
        len(str(cell.value)) if cell.value is not None else 0
        for cell in col
    )
    ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

wb.save(OUTPUT_FILE)

print()
print(f"=== Done! Combined Kestrel file saved to: {OUTPUT_FILE} ===")
print(f"    Sheet : {SHEET_NAME}")
print(f"    Rows  : {len(combined)}")
print(f"    Cols  : {', '.join(combined.columns)}")