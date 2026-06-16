#imports
import pandas as pd
import glob
import os
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np

#Load files
files = glob.glob("raw_data/Test/*.csv")
files = sorted(files, key=lambda f: int(os.path.splitext(os.path.basename(f))[0]))

dfs = []

for f in files:
    recording = os.path.splitext(os.path.basename(f))[0]

    df = pd.read_csv(f, engine="python", on_bad_lines="skip")

    df["recording"] = recording

    dfs.append(df)

data = pd.concat(dfs, ignore_index=True)
data.columns = [c.strip() for c in data.columns]

#combine date and time into a single datetime column
data["datetime"] = pd.to_datetime(
    data["Date"].astype(str) + " " + data["Time"].astype(str),
    errors="coerce"
)

#Remove N/A's
data = data.dropna(subset=["datetime"])

data = data.sort_values("datetime")

#searches a column name by keyword
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


recordings = sorted(data["recording"].unique(), key=int)
cmap = plt.get_cmap("tab20")
colors = {r: cmap(i % 20) for i, r in enumerate(recordings)}

#Plotting
fig, axes = plt.subplots(5, 1, figsize=(14, 18), sharex=True)

plots = [
    (temp_col,  "Temperature (°C)"),
    (light_col, "Light (lux)"),
    (heat_col,  "Heat Index (°C)"),
    (humid_col, "Humidity (%)"),
    (dew_col,   "Dew Point (°C)")
]

for ax, (col, title) in zip(axes, plots):

    if col is None:
        ax.set_title(f"{title} (NOT FOUND)")
        continue

    for r in recordings:
        subset = data[data["recording"] == r].sort_values("datetime")

        ax.plot(
            subset["datetime"],
            subset[col],
            color=colors[r],
            linewidth=1.2,
            label=r
        )

    ax.set_title(title, fontsize=10)
    ax.set_ylabel(title, fontsize=8)
    ax.grid(True, alpha=0.3)

#Formating time
axes[-1].set_xlabel("Time of Day")
axes[-1].xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))
fig.autofmt_xdate()

#Legend
handles, labels = axes[0].get_legend_handles_labels()
fig.legend(handles, labels, title="Recording", loc="upper right", ncol=2)

os.makedirs("outputs", exist_ok=True)

plt.tight_layout()
plt.savefig("outputs/Test_all_recordings_overlay.png", dpi=300, bbox_inches="tight")

plt.show()
