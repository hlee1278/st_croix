import os
import re
import csv
import glob
import pandas as pd


# =============================================================================
# Kestrel averaging + combining
# -----------------------------------------------------------------------------
# Point it at a FOLDER of Kestrel CSVs (sub-folders are searched too).
# For each Kestrel CSV in that folder, this builds ONE summary row:
#   - Kestrel Name  = the file's "Device Name"
#   - Route Number  = the number after R in the filename (0614_R34_... -> 34)
#   - Time          = the file's built-in "Start time"
#   - every measurement column -> the AVERAGE of that column (converted to °F/imperial)
#   - every text column        -> the value from the first row (kept, not averaged)
# All rows are stacked into one new Excel sheet.
# =============================================================================

# EDIT THIS: the folder that holds your Kestrel CSVs.

INPUT_FOLDER = "/Users/timjun/Desktop/USVI_2026/kestrel_data_cleaned"

# Where the combined sheet is saved (default location = alongside this script)
OUTPUT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kestrel_averages.csv")

# Kestrel files have 3 device-info rows, then the header row, then a units row.
METADATA_ROWS = 3

# A column is "quantitative" (gets averaged) when its units-row cell is one of these.
MEASUREMENT_UNITS = {"°C", "°F", "%", "mb", "inHg", "m", "ft", "km/h", "mph", "Deg", "w/m2"}

# What each source unit becomes after converting to imperial
IMPERIAL_OF = {
    "°C": "°F", "°F": "°F",
    "mb": "inHg", "inHg": "inHg",
    "m": "ft", "ft": "ft",
    "km/h": "mph", "mph": "mph",
    "%": "%", "Deg": "Deg", "w/m2": "w/m2",
}


def to_imperial(value, unit):
    """Convert a numeric value from its source unit to imperial."""
    if unit == "°C":
        return value * 9 / 5 + 32
    if unit == "mb":
        return value * 0.0295299830714
    if unit == "m":
        return value * 3.28083989501
    if unit == "km/h":
        return value * 0.621371192
    return value  # °F, inHg, ft, mph, %, Deg, w/m2 are already imperial/unitless


def smart_round(value):
    """Whole numbers for normal readings, but keep decimals for very small values
    (e.g. Vapor Pressure Deficit ~0.4) so they don't collapse to 0."""
    if abs(value) >= 1:
        return round(value)      # whole number
    return round(value, 2)       # keep 2 decimals for values under 1


def read_device_name(path):
    """Pull the Kestrel name out of the first line: 'Device Name,<name>' (ignoring trailing commas).
    utf-8-sig strips the BOM that cleaned exports often add at the start of the file."""
    with open(path, encoding="utf-8-sig", errors="replace", newline="") as f:
        first_line = f.readline()
    fields = next(csv.reader([first_line]))
    return fields[1].strip() if len(fields) > 1 else ""


def is_kestrel_file(path):
    """A Kestrel export starts with a 'Device Name,...' line (BOM-safe)."""
    with open(path, encoding="utf-8-sig", errors="replace") as f:
        first = f.readline().strip().lower()
    return first.startswith("device name")


def route_number(path):
    """Pull the number after R in the filename, dropping any leading zeros (0610_R09_... -> '9')."""
    match = re.search(r"R(\d+)", os.path.basename(path))
    return str(int(match.group(1))) if match else ""


def clean_start_time(raw):
    """The 'Start time' field looks like '2026-06-07 10:48:39 AM 10:48:39 AM'; keep the first datetime."""
    match = re.search(r"\d{4}-\d{2}-\d{2} \d{1,2}:\d{2}:\d{2}\s*[AP]M", str(raw))
    return match.group(0) if match else str(raw)


def is_blank(v):
    return not isinstance(v, str) or v.strip() == "" or v.strip() == "--"


def first_nonblank(series):
    """First value in a column that isn't empty/'--' (skips glitchy first rows)."""
    for v in series:
        if not is_blank(v):
            return v
    return ""


def process_file(path):
    """Return (summary_row_dict, units_dict, ordered_columns, warnings) for one Kestrel CSV."""
    # header = row 4, so the first "data" row read here is actually the units row
    # utf-8-sig handles the BOM that cleaned CSVs often start with
    df = pd.read_csv(path, skiprows=METADATA_ROWS, dtype=str, keep_default_na=False,
                     encoding="utf-8-sig")
    units = df.iloc[0]
    data = df.iloc[1:].reset_index(drop=True)

    device = read_device_name(path)
    route = route_number(path)

    # prefer the built-in Start time field; cleaned exports often only have FORMATTED DATE_TIME
    if "Start time" in data.columns:
        time_val = clean_start_time(first_nonblank(data["Start time"]))
    elif "FORMATTED DATE_TIME" in data.columns:
        time_val = first_nonblank(data["FORMATTED DATE_TIME"])
    else:
        time_val = ""

    row = {"Kestrel Name": device, "Route Number": route, "Time": time_val}
    unit_row = {"Kestrel Name": "", "Route Number": "", "Time": ""}
    ordered = ["Kestrel Name", "Route Number", "Time"]
    warnings = []

    for col in data.columns:
        # the per-reading timestamp column is replaced by our Time column
        if col.strip().upper() == "FORMATTED DATE_TIME":
            continue

        ordered.append(col)
        unit = str(units[col]).strip()

        if unit in MEASUREMENT_UNITS:
            # known measurement -> average, then convert to imperial
            nums = pd.to_numeric(data[col].replace("--", ""), errors="coerce")
            if nums.notna().any():
                row[col] = smart_round(to_imperial(nums.mean(), unit))
            else:
                row[col] = ""  # column was entirely blank / "--"
            unit_row[col] = IMPERIAL_OF.get(unit, unit)
        elif unit != "":
            # a unit I haven't seen before: still average it, but DON'T convert,
            # and flag it so it can be checked/added to the converter
            nums = pd.to_numeric(data[col].replace("--", ""), errors="coerce")
            row[col] = smart_round(nums.mean()) if nums.notna().any() else ""
            unit_row[col] = unit
            warnings.append(
                f"  ! {os.path.basename(path)}: column '{col}' has an unrecognized unit "
                f"'{unit}' - averaged but NOT converted to imperial"
            )
        else:
            # no unit -> non-quantitative -> keep the first non-blank value
            row[col] = first_nonblank(data[col])
            unit_row[col] = ""

            # flag if this "kept" column isn't actually constant within the file
            distinct = {v.strip() for v in data[col] if not is_blank(v)}
            if len(distinct) > 1:
                warnings.append(
                    f"  ! {os.path.basename(path)}: column '{col}' changes within the file "
                    f"({len(distinct)} different values) - kept the first one"
                )

    return row, unit_row, ordered, warnings


def main():
    # use the folder set at the top of the file, or ask for it if that's left blank
    folder = INPUT_FOLDER.strip()
    if not folder:
        folder = input("Paste the path to the folder that holds your Kestrel CSVs: ").strip()
    folder = folder.strip().strip('"').strip("'")

    if not os.path.isdir(folder):
        print(f"Not a folder: {folder}")
        return

    # every .csv in the folder (and any sub-folders), skipping our own output
    paths = sorted(
        p for p in glob.glob(os.path.join(folder, "**", "*.csv"), recursive=True)
        if not os.path.basename(p).startswith("~$")
    )

    if not paths:
        print(f"No .csv files found in: {folder}")
        return

    print(f"\nFound {len(paths)} CSV file(s) in {folder}")

    rows = []
    unit_row = None
    ordered = None
    all_warnings = []

    for path in paths:
        if not is_kestrel_file(path):
            print(f"\nSkipping (not a Kestrel file): {os.path.basename(path)}")
            continue
        print(f"\nProcessing: {os.path.basename(path)}")
        row, units, cols, warnings = process_file(path)
        rows.append(row)
        all_warnings.extend(warnings)
        if ordered is None:
            ordered, unit_row = cols, units
        else:
            # keep any columns a later file introduces (and carry their units)
            for c in cols:
                if c not in ordered:
                    ordered.append(c)
                    unit_row[c] = units.get(c, "")

    if not rows:
        print("\nNo Kestrel files to process.")
        return

    # build the sheet: units row first (Kestrel style), then one row per file
    out = pd.DataFrame([unit_row] + rows)
    out = out.reindex(columns=ordered)
    out.to_csv(OUTPUT_FILE, index=False)

    print("\n" + "=" * 60)
    if all_warnings:
        print("FLAGS (please double-check these files):")
        for w in all_warnings:
            print(w)
    else:
        print("No consistency problems found.")
    print("=" * 60)
    print(f"\nSaved {len(rows)} row(s) to:\n  {OUTPUT_FILE}")
    print("All done!")


if __name__ == "__main__":
    main()
