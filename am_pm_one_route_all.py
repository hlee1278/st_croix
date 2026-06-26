"""
HEAT MAPPING DATA COMBINATION SCRIPT
======================================
PURPOSE:
    Reads all cleaned Excel session files from a single folder (e.g., all_R27),
    automatically detects:
      - The route number and date from each filename
        (e.g., "combined_R27_06212026_PL1_AM_noGPS.xlsx" → route 27, date 6/21/2026)
      - The PocketLab number from each filename
        (e.g., "PL1" → pocketlab = 1; if absent → pocketlab = "Unknown")
      - Whether each file is morning or afternoon (AM/PM from filename)

    Produces one combined Excel file containing all rows from all input files:
      e.g.  full_R27_06212026.xlsx

    ERRORS ARE RAISED IF:
      - Two files share the same PocketLab number AND the same time block
      - Files in the folder belong to different routes or different dates

HOW TO USE THIS SCRIPT:
    1. Fill in Section 1 below (the only section you need to edit)
    2. Run the script:   python heat_mapping_combine.py
    3. The combined Excel file will be saved to the folder you specify in OUTPUT_DIR

DEPENDENCIES — install once if not already installed:
    pip install pandas openpyxl
"""

import os
import re
import sys
import pandas as pd


# =============================================================================
# SECTION 1: CONFIGURATION — THIS IS THE ONLY SECTION YOU NEED TO EDIT
# =============================================================================

INPUT_DIR  = r"C:\Users\dresd\USVI_2026\combined_data\0621_R27\all_R27"
                            # Path to the folder containing all your cleaned Excel files
                            # All files must belong to the same route and date

OUTPUT_DIR = r"C:\Users\dresd\USVI_2026\combined_data\0621_R27"
                            # Folder where the combined Excel file will be saved
                            # This folder must already exist before running the script


# =============================================================================
# END OF USER INPUT — DO NOT EDIT BELOW THIS LINE UNLESS CUSTOMIZING
# =============================================================================


# =============================================================================
# SECTION 2: FIND ALL EXCEL FILES IN THE INPUT FOLDER
# =============================================================================

EXPECTED_COLUMNS = [
    "index", "route_id", "route_name", "route_number", "time_block",
    "date", "time", "latitude", "longitude", "temp_probe",
    "humidity", "dew_point", "heat_index", "ambient_light"
]

OUTPUT_COLUMNS = [
    "pocketlab", "route_id", "route_name", "route_number", "time_block",
    "date", "time", "latitude", "longitude", "temp_probe",
    "humidity", "dew_point", "heat_index", "ambient_light"
]

if not os.path.isdir(INPUT_DIR):
    print(f"ERROR: Input folder not found: {INPUT_DIR}")
    sys.exit(1)

if not os.path.isdir(OUTPUT_DIR):
    print(f"ERROR: Output folder does not exist: {OUTPUT_DIR}")
    print("       Please create this folder before running the script.")
    sys.exit(1)

all_xlsx = sorted([
    f for f in os.listdir(INPUT_DIR)
    if f.lower().endswith(".xlsx")
])

if not all_xlsx:
    print(f"ERROR: No Excel (.xlsx) files found in folder: {INPUT_DIR}")
    sys.exit(1)

print(f"=== Found {len(all_xlsx)} Excel file(s) in folder ===")
for f in all_xlsx:
    print(f"  {f}")
print()


# =============================================================================
# SECTION 3: PARSE METADATA FROM EACH FILENAME
# =============================================================================
# Expected filename format:
#   combined_R{route}_{MMDDYYYY}_{PL#}_{AM|PM}_noGPS.xlsx
#
# Examples:
#   combined_R27_06212026_PL1_AM_noGPS.xlsx  → route=27, date=06212026, pocketlab=PL1, block=AM
#   combined_R27_06212026_PL17_PM_noGPS.xlsx → route=27, date=06212026, pocketlab=PL17, block=PM
#   combined_R27_06212026_BH_AM_noGPS.xlsx   → route=27, date=06212026, pocketlab=Unknown, block=AM

FILE_PATTERN = re.compile(
    r"^combined_R(\d+)_(\d{8})_(.+?)_(AM|PM)_noGPS\.xlsx$",
    re.IGNORECASE
)

def parse_filename(filename):
    """
    Parse route number, date string, pocketlab label, and time block from a filename.

    Returns a dict with keys:
        route_number  : int
        date_str      : str  (e.g. "06212026")
        pocketlab     : str  (e.g. "PL1", "PL17", or "Unknown")
        time_block    : str  ("morning" or "afternoon")
        time_label    : str  ("AM" or "PM")

    Returns None if the filename does not match the expected pattern.
    """
    m = FILE_PATTERN.match(filename)
    if not m:
        return None

    route_number = int(m.group(1))
    date_str     = m.group(2)          # e.g. "06212026"
    middle_part  = m.group(3)          # e.g. "PL1", "PL17", "BH"
    time_label   = m.group(4).upper()  # "AM" or "PM"

    # Determine pocketlab: any part matching PL followed by digits is a PocketLab ID
    if re.fullmatch(r"PL\d+", middle_part, re.IGNORECASE):
        pocketlab = middle_part.upper()
    else:
        pocketlab = "Unknown"

    time_block = "morning" if time_label == "AM" else "afternoon"

    return {
        "route_number": route_number,
        "date_str":     date_str,
        "pocketlab":    pocketlab,
        "time_block":   time_block,
        "time_label":   time_label,
    }


parsed_files = []   # List of (filename, metadata_dict)
unparseable  = []   # Filenames that didn't match the pattern

for filename in all_xlsx:
    meta = parse_filename(filename)
    if meta is None:
        unparseable.append(filename)
    else:
        parsed_files.append((filename, meta))

if unparseable:
    print("ERROR: The following file(s) do not match the expected naming format")
    print("       (expected: combined_R{route}_{MMDDYYYY}_{label}_{AM|PM}_noGPS.xlsx):")
    for f in unparseable:
        print(f"  {f}")
    sys.exit(1)


# =============================================================================
# SECTION 4: VALIDATE — ALL FILES MUST SHARE THE SAME ROUTE AND DATE
# =============================================================================

routes_found = set(meta["route_number"] for _, meta in parsed_files)
dates_found  = set(meta["date_str"]     for _, meta in parsed_files)

if len(routes_found) > 1:
    print("ERROR: Files in this folder belong to more than one route.")
    print(f"       Routes found: {sorted(routes_found)}")
    print("       All files in the input folder must share the same route number.")
    sys.exit(1)

if len(dates_found) > 1:
    print("ERROR: Files in this folder belong to more than one date.")
    print(f"       Dates found: {sorted(dates_found)}")
    print("       All files in the input folder must share the same date.")
    sys.exit(1)

ROUTE_NUMBER  = next(iter(routes_found))   # e.g. 27
DATE_STR      = next(iter(dates_found))    # e.g. "06212026"

print(f"=== Validated: all files share route R{ROUTE_NUMBER}, date {DATE_STR} ===")
print()


# =============================================================================
# SECTION 5: VALIDATE — NO DUPLICATE POCKETLAB + TIME BLOCK COMBINATIONS
# =============================================================================

seen_combos = {}   # { (pocketlab, time_block): filename }

for filename, meta in parsed_files:
    key = (meta["pocketlab"], meta["time_block"])
    if key in seen_combos:
        print("ERROR: Duplicate PocketLab + time block combination detected.")
        print(f"       PocketLab : {meta['pocketlab']}")
        print(f"       Time block: {meta['time_block']}")
        print(f"       File 1: {seen_combos[key]}")
        print(f"       File 2: {filename}")
        print("       Each PocketLab must have at most one AM file and one PM file.")
        sys.exit(1)
    seen_combos[key] = filename

print("=== Validated: no duplicate PocketLab + time block combinations ===")
print()


# =============================================================================
# SECTION 6: READ AND COMBINE ALL FILES
# =============================================================================

print("=== Reading files ===")

all_frames = []

for filename, meta in sorted(parsed_files, key=lambda x: (x[1]["pocketlab"], x[1]["time_label"])):
    filepath = os.path.join(INPUT_DIR, filename)
    print(f"  Reading: {filename}")
    print(f"    PocketLab  : {meta['pocketlab']}")
    print(f"    Time block : {meta['time_label']}")

    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        print(f"  ERROR: Could not read file '{filename}': {e}")
        sys.exit(1)

    # --- Check that all expected columns are present ---
    missing = [c for c in EXPECTED_COLUMNS if c not in df.columns]
    if missing:
        print(f"  ERROR: The following expected columns are missing from '{filename}':")
        for c in missing:
            print(f"    • '{c}'")
        print("  Columns actually found in this file:")
        for c in df.columns:
            print(f"    • '{c}'")
        sys.exit(1)

    # --- Drop the index column ---
    df = df.drop(columns=["index"])

    # --- Coerce numeric columns to proper number types ---
    numeric_cols = ["route_number", "latitude", "longitude", "temp_probe",
                    "humidity", "dew_point", "heat_index", "ambient_light"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # --- Ensure text columns are stored as strings, not mixed types ---
    text_cols = ["route_id", "route_name", "time_block", "date", "time"]
    for col in text_cols:
        df[col] = df[col].astype(str)

    # --- Add the pocketlab column at the front ---
    df.insert(0, "pocketlab", meta["pocketlab"])

    # --- Enforce output column order ---
    df = df[OUTPUT_COLUMNS]

    print(f"    Rows read  : {len(df)}")
    all_frames.append(df)
    print()

combined = pd.concat(all_frames, ignore_index=True)

print(f"=== Combined total rows: {len(combined)} ===")
print()


# =============================================================================
# SECTION 7: SAVE COMBINED OUTPUT TO EXCEL
# =============================================================================

output_filename_base = f"full_R{ROUTE_NUMBER}_{DATE_STR}"
output_filepath      = os.path.join(OUTPUT_DIR, output_filename_base + ".xlsx")

# Excel sheet names must be 31 characters or fewer
sheet_name = output_filename_base[:31]

print(f"=== Saving output ===")
print(f"  File  : {output_filepath}")
print(f"  Sheet : {sheet_name}")

try:
    combined.to_excel(output_filepath, sheet_name=sheet_name, index=False)
except Exception as e:
    print(f"ERROR: Could not save output file: {e}")
    sys.exit(1)

print(f"  Rows  : {len(combined)}")
print(f"  Cols  : {', '.join(OUTPUT_COLUMNS)}")
print()
print(f"=== Done! Combined file saved to: {output_filepath} ===")