import pandas as pd
import glob
import os
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# ---------------------------------------------------
# 1. LOAD FILES
# ---------------------------------------------------
files = glob.glob("raw_data/0608_R31/*.csv")

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
    df["trial"] = parts[3]

    dfs.append(df)

data = pd.concat(dfs, ignore_index=True)
data.columns = [c.strip() for c in data.columns]

# ---------------------------------------------------
# 2. FIND COLUMNS
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

# ---------------------------------------------------
# 3. CLEAN NUMBERS
# ---------------------------------------------------
def extract_numeric(series):
    return series.astype(str).str.extract(r"([-+]?\d*\.?\d+)")[0].astype(float)

for col in [temp_col, light_col, heat_col, humidity_col, dew_col]:
    if col:
        data[col] = extract_numeric(data[col])

# ---------------------------------------------------
# 4. BUILD REAL TIME (CRITICAL FIX)
# ---------------------------------------------------
# PocketLab usually has Date + Time columns
if "Date" in data.columns and "Time" in data.columns:

    data["datetime"] = pd.to_datetime(
        data["Date"].astype(str) + " " + data["Time"].astype(str),
        errors="coerce"
    )

else:
    # fallback if missing
    data["datetime"] = pd.to_datetime(data[time_col], errors="coerce")

# remove bad timestamps only
data = data.dropna(subset=["datetime", "person"])

# IMPORTANT: sort time
data = data.sort_values("datetime")

# ---------------------------------------------------
# 5. PEOPLE LIST
# ---------------------------------------------------
people = sorted(data["person"].unique())

# ---------------------------------------------------
# 6. MULTI-PANEL FIGURE (REAL TIME AXIS)
# ---------------------------------------------------
fig, axes = plt.subplots(5, 1, figsize=(12, 18), sharex=True)

plots = [
    (temp_col, "Temperature Probe"),
    (light_col, "Ambient Light"),
    (heat_col, "Heat Index"),
    (humidity_col, "Humidity"),
    (dew_col, "Dew Point")
]

for ax, (col, title) in zip(axes, plots):

    if col is None:
        ax.set_title(f"{title} (NOT FOUND)")
        continue

    for p in people:
        subset = data[data["person"] == p]

        ax.plot(
            subset["datetime"],
            subset[col],
            label=p,
            linewidth=1
        )

    ax.set_title(title)
    ax.set_ylabel(title)
    ax.grid(True)

# ---------------------------------------------------
# 7. TIME AXIS FORMATTING (SHOW BREAKS + REAL TIME)
# ---------------------------------------------------
axes[-1].set_xlabel("Time of Day")

locator = mdates.AutoDateLocator()
formatter = mdates.DateFormatter("%H:%M")

axes[-1].xaxis.set_major_locator(locator)
axes[-1].xaxis.set_major_formatter(formatter)

fig.autofmt_xdate()

# ---------------------------------------------------
# 8. SINGLE LEGEND (NO DUPLICATION)
# ---------------------------------------------------
handles, labels = axes[0].get_legend_handles_labels()
fig.legend(handles, labels, title="Person", loc="upper right")

# ---------------------------------------------------
# 9. SAVE
# ---------------------------------------------------
os.makedirs("outputs", exist_ok=True)

plt.tight_layout()
plt.savefig("outputs/0608_R31_multiplot_time_correct.png",
            dpi=300, bbox_inches="tight")

plt.show()