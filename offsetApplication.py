import pandas as pd
import glob
import os


#THIS IS THE ONLY LINE YOU NEED TO EDIT
#Put the path to the folder that holds your Excel files below. This folder is the one that contains the "somewhat clean" data (it's the folder that has one excel file and then a subfolder)
ROOT_DIR = "/Users/timjun/Desktop/USVI_2026/combined_data/0607_R17_wNewCode"

#Nothing to touch below this line
#Folder created INSIDE ROOT_DIR to hold the offset files. It keeps the
#original folder's name and just adds "offset output".
OUTPUT_DIR = os.path.basename(os.path.normpath(ROOT_DIR)) + " offset output"

#These are the offset numbers based on Dresden's tests
offsets = {
    1:  {"humidity": -8.722, "temp":  0.051, "dew": -1.276, "heat": -0.587},
    2:  {"humidity": -7.045, "temp":  0.532, "dew": -0.757, "heat": -0.029},
    3:  {"humidity": -8.939, "temp": -0.437, "dew": -1.037, "heat":  0.057},
    4:  {"humidity": -7.555, "temp": -0.175, "dew": -1.052, "heat": -0.407},
    5:  {"humidity": -7.841, "temp":  0.030, "dew": -1.371, "heat": -1.147},
    6:  {"humidity": -8.211, "temp": -1.020, "dew": -1.163, "heat": -0.444},
    7:  {"humidity": -7.306, "temp":  0.482, "dew": -0.745, "heat":  0.132},
    8:  {"humidity": -7.198, "temp":  0.447, "dew": -0.717, "heat":  0.149},
    9:  {"humidity": -5.992, "temp":  0.353, "dew": -0.373, "heat":  0.487},
    # PL 10 intentionally omitted (no offset in the tables)
    11: {"humidity": -5.218, "temp": -0.237, "dew": -0.317, "heat":  0.324},
    12: {"humidity": -10.771, "temp": -0.242, "dew": -1.988, "heat": -1.281},
    13: {"humidity": -8.191, "temp":  0.241, "dew": -0.972, "heat": -0.026},
    14: {"humidity": -8.847, "temp": -0.163, "dew": -1.204, "heat": -0.303},
    15: {"humidity": -7.588, "temp":  0.474, "dew": -0.969, "heat": -0.218},
    16: {"humidity": -7.298, "temp":  0.197, "dew": -1.409, "heat": -1.236},
    17: {"humidity": -4.938, "temp":  0.050, "dew": -0.495, "heat": -0.200},
    18: {"humidity": -6.316, "temp": -0.495, "dew": -0.859, "heat": -0.464},
    19: {"humidity": -8.709, "temp":  0.489, "dew": -1.289, "heat": -0.476},
}

# maps each offset category to the keyword used to find its column in the sheet
CATEGORY_KEYWORDS = {
    "temp":     ["temp"],
    "humidity": ["humid"],
    "dew":      ["dew"],
    "heat":     ["heat"],
}


#find a column whose name contains any of the given keywords
def find_col(df, keys):
    for c in df.columns:
        if any(k in str(c).lower() for k in keys):
            return c
    return None


#add the per-pocketlab offset to each category column; returns (df, rows_changed)
def apply_offsets(df):
    pl_col = find_col(df, ["pocketlab"])
    if pl_col is None:
        print("  ! no 'pocketlab' column found - leaving file unchanged")
        return df, 0

    #pull the number out of values like 'PL 17', 'PL17', or '17' (blank -> NaN)
    pl_num = pd.to_numeric(
        df[pl_col].astype(str).str.extract(r"(\d+)")[0],
        errors="coerce",
    )

    #a row is offset only if its pocketlab has an entry in the offsets table
    applied = pl_num.map(lambda p: pd.notna(p) and int(p) in offsets)

    for category, keywords in CATEGORY_KEYWORDS.items():
        col = find_col(df, keywords)
        if col is None:
            print(f"  ! no column found for '{category}' - skipping it")
            continue

        #per-row offset for this category (0 when the pocketlab has no offset)
        off = pl_num.map(
            lambda p: offsets.get(int(p), {}).get(category, 0.0) if pd.notna(p) else 0.0
        ).astype(float)

        df[col] = (pd.to_numeric(df[col], errors="coerce") - off).round(3)

    return df, int(applied.sum())


def main():
    #every .xlsx under ROOT_DIR, skipping Excel temp files and prior outputs
    paths = [
        p for p in glob.glob(os.path.join(ROOT_DIR, "**", "*.xlsx"), recursive=True)
        if not os.path.basename(p).startswith("~$")
        and not p.endswith("_offset.xlsx")
    ]

    if not paths:
        print(f"No .xlsx files found under: {ROOT_DIR}")
        return

    print(f"Found {len(paths)} Excel file(s) under {ROOT_DIR}\n")

    for path in paths:
        print(f"Processing: {path}")
        df = pd.read_excel(path)
        df, n = apply_offsets(df)

        #put the offset folder INSIDE ROOT_DIR, mirroring its sub-folder layout
        rel = os.path.relpath(path, ROOT_DIR)
        base, ext = os.path.splitext(rel)
        out_path = os.path.join(ROOT_DIR, OUTPUT_DIR, base + "_offset" + ext)
        os.makedirs(os.path.dirname(out_path), exist_ok=True)

        df.to_excel(out_path, index=False)
        print(f"  offset applied to {n} rows -> {out_path}\n")

    print("All done!")


if __name__ == "__main__":
    main()
