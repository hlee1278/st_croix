"""
apply_adjustments.py
---------------------
Reads the offsets spreadsheet produced by compare_sensors.py and applies
those corrections to the original PocketLab data, saving a new Excel file.

HOW THE ADJUSTMENT WORKS:
  The offset for each PocketLab represents how much that device reads HIGH
  or LOW compared to the Kestrel reference average, calculated as:

      offset = mean(PocketLab_value − Kestrel_average)

  To correct the PocketLab readings, we simply subtract that offset:

      adjusted_value = raw_value − offset

  So if PL 3 reads +1.2°C too high on average, every Temperature reading
  from PL 3 will be reduced by 1.2°C in the output file.
  If PL 7 reads −0.8% too low on Humidity, every Humidity reading from
  PL 7 will be increased by 0.8% in the output file.

  All other columns (Date, Time, Elapsed Seconds, Ambient Light, etc.)
  are left completely unchanged.
"""

# ════════════════════════════════════════════════════════════════
#  INPUT YOUR FILE PATHS HERE
# ════════════════════════════════════════════════════════════════
POCKETLAB_FILE  = r"C:\Users\dresd\Downloads\0618_R1000\COMBINED_0618_R1000.xlsx"    # ← original PocketLab Excel file
OFFSETS_FILE    = r"C:\Users\dresd\Downloads\offsets.xlsx"           # ← offsets.xlsx from compare_sensors.py
OUTPUT_FILE     = r"C:\Users\dresd\Downloads\pocketlab_adjusted.xlsx"  # ← name for the corrected output file
# ════════════════════════════════════════════════════════════════

import pandas as pd
import numpy as np
from pathlib import Path

# Maps offset spreadsheet column names → PocketLab column names
METRIC_MAP = {
    "Heat Index Avg Offset (°C)":   "Heat Index",
    "Dew Point Avg Offset (°C)":    "Dew Point",
    "Temperature Avg Offset (°C)":  "Temperature Probe",
    "Humidity Avg Offset (%)":      "Humidity",
}

def main():
    # ── Load original PocketLab data ──────────────────────────────────────
    print(f"Loading PocketLab data from: {POCKETLAB_FILE}")
    pl_df = pd.read_excel(POCKETLAB_FILE, engine="openpyxl")
    pl_df.columns = pl_df.columns.str.strip()

    # ── Load offsets from the Summary sheet ──────────────────────────────
    print(f"Loading offsets from:        {OFFSETS_FILE}")
    offsets_df = pd.read_excel(OFFSETS_FILE, sheet_name="Summary", engine="openpyxl")
    offsets_df.columns = offsets_df.columns.str.strip()

    # Parse PocketLab number out of "PL 1", "PL 2", etc.
    offsets_df["pocketlab_number"] = (
        offsets_df["PocketLab #"]
        .str.replace("PL ", "", regex=False)
        .astype(int)
    )
    offsets_df = offsets_df.set_index("pocketlab_number")

    # ── Apply adjustments ─────────────────────────────────────────────────
    adjusted_df = pl_df.copy()

    devices_adjusted = []
    for pl_num in offsets_df.index:
        mask = adjusted_df["pocketlab_number"] == pl_num
        if not mask.any():
            print(f"  ⚠ PL {pl_num} not found in PocketLab data — skipping")
            continue

        for offset_col, pl_col in METRIC_MAP.items():
            if pl_col not in adjusted_df.columns:
                print(f"  ⚠ Column '{pl_col}' not found in PocketLab file — skipping")
                continue
            if offset_col not in offsets_df.columns:
                print(f"  ⚠ Offset column '{offset_col}' not found in offsets file — skipping")
                continue

            offset_val = offsets_df.loc[pl_num, offset_col]
            if pd.isna(offset_val) or offset_val == "":
                continue  # no offset available for this device/metric

            adjusted_df.loc[mask, pl_col] = (
                pd.to_numeric(adjusted_df.loc[mask, pl_col], errors="coerce") - offset_val
            )

        devices_adjusted.append(pl_num)

    print(f"\nAdjustments applied to {len(devices_adjusted)} PocketLab(s): {devices_adjusted}")

    # ── Save output ───────────────────────────────────────────────────────
    out_path = Path(POCKETLAB_FILE).parent / OUTPUT_FILE
    adjusted_df.to_excel(out_path, index=False, engine="openpyxl")
    print(f"\n✅ Adjusted data saved to: {out_path}")

if __name__ == "__main__":
    main()