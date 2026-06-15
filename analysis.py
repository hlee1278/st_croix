import pandas as pd
import glob
import os
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np

# ---------------------------------------------------
# 1. SELECT ROUTE + LOAD FOLDER
# ---------------------------------------------------
route = input("Enter route (e.g. R31): ").upper()

matching_folders = glob.glob(f"raw_data/*_{route}")

if len(matching_folders) == 0:
    raise ValueError(f"No folder found for route {route}")

folder = matching_folders[0]
print("Using folder:", folder)

files = glob.glob(os.path.join(folder, "*.csv"))

dfs = []

# ---------------------------------------------------
# 2. LOAD FILES (IGNORE NON-TRIAL FILES)
# ---------------------------------------------------
for f in files:

    name = os.path.basename(f)

    if "_T" not in name:
        continue

    df = pd.read_csv(f, engine="python", on_bad_lines="skip")

    parts = name.replace(".csv", "").split("_")
    if len(parts) < 4:
        continue

    df["person"] = parts[2]

    dfs.append(df)

data = pd.concat(dfs, ignore_index=True)
data.columns = [c.strip() for c in data.columns]

# ---------------------------------------------------
# 3. FIND COLUMNS ROBUSTLY
# ---------------------------------------------------
def find_col(df, keywords):
    best, best_score = None, 0
    for col in df.columns:
        low = col.lower()
        score = sum(k in low for k in keywords)
        if score > best_score:
            best, best_score = col, score
    return best

time_col = find_col(data, ["elapsed"]) or find_col(data, ["time"])

temp_col = find_col(data, ["temp"])
light_col = find_col(data, ["light"])
heat_col = find_col(data, ["heat"])
humidity_col = find_col(data, ["humid"])
dew_col = find_col(data, ["dew"])

print("\nDetected columns:")
print(time_col, temp_col, light_col, heat_col, humidity_col, dew_col)

# ---------------------------------------------------
# 4. CLEAN NUMERIC DATA
# ---------------------------------------------------
def extract_numeric(series):
    return (
        series.astype(str)
        .str.extract(r"([-+]?\d*\.?\d+)")[0]
        .astype(float)
    )

for col in [temp_col, light_col, heat_col, humidity_col, dew_col]:
    if col:
        data[col] = extract_numeric(data[col])

# ---------------------------------------------------
# 5. CREATE REAL DATETIME
# ---------------------------------------------------
if "Date" in data.columns and "Time" in data.columns:

    data["datetime"] = pd.to_datetime(
        data["Date"].astype(str) + " " + data["Time"].astype(str),
        errors="coerce"
    )

else:
    data["datetime"] = pd.to_datetime(data[time_col], errors="coerce")

data = data.dropna(subset=["datetime", "person"])
data = data.sort_values("datetime")

# ---------------------------------------------------
# 6. AM / PM SPLIT (FIXED AT 1:30 PM)
# ---------------------------------------------------
data["hour_decimal"] = (
    data["datetime"].dt.hour +
    data["datetime"].dt.minute / 60
)

LUNCH_SPLIT = 13.5  # 1:30 PM

data["period"] = np.where(
    data["hour_decimal"] < LUNCH_SPLIT,
    "AM",
    "PM"
)

# ---------------------------------------------------
# 7. OUTLIER REMOVAL (±3 STD PER PERSON)
# ---------------------------------------------------
variables = [temp_col, light_col, heat_col, humidity_col, dew_col]

for person in data["person"].unique():

    mask = data["person"] == person

    for col in variables:

        if col is None:
            continue

        mean = data.loc[mask, col].mean()
        sd = data.loc[mask, col].std()

        upper = mean + 3 * sd
        lower = mean - 3 * sd

        outliers = mask & (
            (data[col] > upper) | (data[col] < lower)
        )

        data.loc[outliers, col] = np.nan

# ---------------------------------------------------
# 8. PEOPLE + COLORS
# ---------------------------------------------------
people = sorted(data["person"].unique())

colors = {
    "KH": "#1f77b4",  # blue
    "PU": "#d62728",  # red
    "CC": "#2ca02c"   # green
}

default_color = "black"

# ---------------------------------------------------
# 9. PLOT SETUP (5 x 2 = 10 PLOTS)
# ---------------------------------------------------
fig, axes = plt.subplots(
    5,
    2,
    figsize=(18, 18),
    sharex=False
)

plots = [
    (temp_col, "Temperature Probe"),
    (light_col, "Ambient Light"),
    (heat_col, "Heat Index"),
    (humidity_col, "Humidity"),
    (dew_col, "Dew Point")
]

# ---------------------------------------------------
# 10. PLOTTING
# ---------------------------------------------------
for row, (col, title) in enumerate(plots):

    if col is None:
        continue

    for col_idx, period in enumerate(["AM", "PM"]):

        ax = axes[row, col_idx]

        subset_period = data[data["period"] == period]

        for person in people:

            subset = subset_period[
                subset_period["person"] == person
            ]

            ax.plot(
                subset["datetime"],
                subset[col],
                label=person,
                linewidth=1.5,
                color=colors.get(person, default_color)
            )

        ax.set_title(f"{title} ({period})")
        ax.grid(True, alpha=0.3)

        ax.xaxis.set_major_formatter(
            mdates.DateFormatter("%H:%M")
        )

        ax.tick_params(axis="x", rotation=45)

# ---------------------------------------------------
# 11. LEGEND (ONCE)
# ---------------------------------------------------
handles, labels = axes[0, 0].get_legend_handles_labels()

fig.legend(
    handles,
    labels,
    title="Person",
    loc="upper right"
)

# ---------------------------------------------------
# 12. SAVE OUTPUT
# ---------------------------------------------------
os.makedirs("outputs", exist_ok=True)

plt.tight_layout()

plt.savefig(
    f"outputs/{route}_AM_PM_multiplot.png",
    dpi=300,
    bbox_inches="tight"
)

plt.show()