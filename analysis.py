#imports
import pandas as pd
import glob
import os
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np

#input route number
route = input("Enter route (e.g. R31): ").upper()

folder = glob.glob(f"raw_data/*_{route}")[0]
files = glob.glob(os.path.join(folder, "*.csv"))

#load data
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

#date and time
data["datetime"] = pd.to_datetime(
    data["Date"].astype(str) + " " + data["Time"].astype(str),
    errors="coerce"
)
data = data.dropna(subset=["datetime", "person"])
data = data.sort_values("datetime")

#find variables
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

#eliminate outliers outside of +-3 standard deviations
for p in data["person"].unique():
    mask = data["person"] == p
    for c in vars_all:
        if c:
            m = data.loc[mask, c].mean()
            s = data.loc[mask, c].std()
            data.loc[mask & (data[c] > m + 3*s), c] = np.nan
            data.loc[mask & (data[c] < m - 3*s), c] = np.nan

#build series with gaps
def build_complete_series(df, person, value_col):
    sub = df[df["person"] == person].copy()
    sub = sub.sort_values("datetime").set_index("datetime")

    if len(sub) < 2:
        return sub

    diffs = sub.index.to_series().diff().dropna()
    diffs_pos = diffs[diffs > pd.Timedelta(0)]
    dt = diffs_pos.mode()[0]
    gap_threshold = dt * 1.5

    gap_mask = diffs > gap_threshold
    gap_starts = diffs[gap_mask].index

    if len(gap_starts) == 0:
        return sub

    nan_rows = pd.DataFrame(
        {c: np.nan for c in sub.columns},
        index=gap_starts - pd.Timedelta(milliseconds=1)
    )
    nan_rows["person"] = person

    result = pd.concat([sub, nan_rows]).sort_index()
    return result

#colors
people = sorted(data["person"].unique())
default_colors = plt.rcParams["axes.prop_cycle"].by_key()["color"]
colors = {p: default_colors[i % len(default_colors)] for i, p in enumerate(people)}

#plot
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
                color=colors[p],
                linewidth=1.5
            )

        ax.set_title(f"{title} ({period})", fontsize=9)
        ax.set_ylabel(title, fontsize=8)
        ax.grid(True, alpha=0.3)
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))
        ax.tick_params(axis="x", rotation=45, labelsize=7)
        ax.tick_params(axis="y", labelsize=7)

#legend
handles = [plt.Line2D([0], [0], color=colors[p], lw=2) for p in people]
labels  = list(people)
fig.legend(handles, labels, loc="upper right", title="Person")

#save
os.makedirs("outputs", exist_ok=True)
plt.tight_layout()
plt.savefig(
    f"outputs/{route}_AM_PM_multiplot.png",
    dpi=300,
    bbox_inches="tight"
)
plt.show()