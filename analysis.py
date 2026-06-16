print("started script")

import pandas as pd
import glob
import os
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np

# ---------------------------------------------------
# 1. ROUTE INPUT
# ---------------------------------------------------
route = input("Enter route (e.g. R31): ").upper()

folder = glob.glob(f"raw_data/*_{route}")[0]
files = glob.glob(os.path.join(folder, "*.csv"))

# ---------------------------------------------------
# 2. LOAD DATA
# ---------------------------------------------------
dfs = []

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
# 3. DATETIME
# ---------------------------------------------------
data["datetime"] = pd.to_datetime(
    data["Date"].astype(str) + " " + data["Time"].astype(str),
    errors="coerce"
)
data = data.dropna(subset=["datetime", "person"])
data = data.sort_values("datetime")

# ---------------------------------------------------
# 4. FIND VARIABLES
# ---------------------------------------------------
def find_col(df, keys):
    for c in df.columns:
        if any(k in c.lower() for k in keys):
            return c
    return None

temp_col  = find_col(data, ["temp"])
light_col = find_col(data, ["light"])
heat_col  = find_col(data, ["heat"])
humid_col = find_col(data, ["humid"])
dew_col   = find_col(data, ["dew"])

vars_all = [temp_col, light_col, heat_col, humid_col, dew_col]

def to_num(s):
    return pd.to_numeric(
        s.astype(str).str.extract(r"([-+]?\d*\.?\d+)")[0],
        errors="coerce"
    )

for c in vars_all:
    if c:
        data[c] = to_num(data[c])



# ---------------------------------------------------
# 5. OUTLIERS → NaN
# ---------------------------------------------------
for p in data["person"].unique():
    mask = data["person"] == p
    for c in vars_all:
        if c:
            m = data.loc[mask, c].mean()
            s = data.loc[mask, c].std()
            data.loc[mask & (data[c] > m + 3*s), c] = np.nan
            data.loc[mask & (data[c] < m - 3*s), c] = np.nan

# ---------------------------------------------------
# 6. BUILD COMPLETE SERIES WITH REAL GAP BREAKS
# ---------------------------------------------------
def build_complete_series(df, person, value_col):
    sub = df[df["person"] == person].copy()
    sub = sub.sort_values("datetime").set_index("datetime")

    if len(sub) < 2:
        return sub

    diffs = sub.index.to_series().diff().dropna()
    diffs = diffs[diffs > pd.Timedelta(0)]
    dt = diffs.mode()[0]

    gap_threshold = dt * 1.5

    rows = []
    for i in range(len(sub)):
        rows.append(sub.iloc[[i]])
        if i < len(sub) - 1:
            gap = sub.index[i + 1] - sub.index[i]
            if gap > gap_threshold:
                mid = sub.index[i] + gap / 2
                nan_row = pd.DataFrame(
                    {c: [np.nan] for c in sub.columns},
                    index=[mid]
                )
                nan_row["person"] = person
                rows.append(nan_row)

    return pd.concat(rows)

# ---------------------------------------------------
# 7. COLORS
# ---------------------------------------------------
colors = {
    "KH": "#4C72B0",
    "PU": "#DD8452",
    "CC": "#55A868"
}

people = sorted(data["person"].unique())

# ---------------------------------------------------
# 8. PLOT
# ---------------------------------------------------
fig, axes = plt.subplots(5, 2, figsize=(18, 18))

plots = [
    (temp_col,  "Temperature (°C)"),
    (light_col, "Light (lux)"),
    (heat_col,  "Heat Index (°C)"),
    (humid_col, "Humidity (%)"),
    (dew_col,   "Dew Point (°C)")
]

for r, (col, title) in enumerate(plots):
    if col is None:
        continue

    for j, period in enumerate(["AM", "PM"]):
        ax = axes[r, j]

        for p in people:
            series = build_complete_series(data, p, col)

            hour = series.index.hour + series.index.minute / 60
            if period == "AM":
                series = series[hour < 13.5]
            else:
                series = series[hour >= 13.5]

            ax.plot(
                series.index,
                series[col],
                color=colors.get(p),
                linewidth=1.5
            )

        ax.set_title(f"{title} ({period})", fontsize=9)
        ax.set_ylabel(title, fontsize=8)
        ax.grid(True, alpha=0.3)
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))
        ax.tick_params(axis="x", rotation=45, labelsize=7)
        ax.tick_params(axis="y", labelsize=7)

# ---------------------------------------------------
# 9. LEGEND
# ---------------------------------------------------
handles = [plt.Line2D([0], [0], color=c, lw=2) for c in colors.values()]
labels  = list(colors.keys())
fig.legend(handles, labels, loc="upper right", title="Person")

# ---------------------------------------------------
# 10. SAVE
# ---------------------------------------------------
os.makedirs("outputs", exist_ok=True)
plt.tight_layout()
plt.savefig(
    f"outputs/{route}_FINAL_NO_CONTINUITY.png",
    dpi=300,
    bbox_inches="tight"
)
plt.show()