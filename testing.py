#imports
import pandas as pd
import glob
import os
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import textwrap

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

#SIMILARITY CODE

elapsed_col = find_col(data, ["elapsed"])

#Rounding to the nearest second (not a floor func)
data["sec"] = data[elapsed_col].round().astype("Int64")

groups = {
    "1-4":   ["1", "2", "3", "4"],
    "5-8":   ["5", "6", "7", "8"],
    "9-12":  ["9", "10", "11", "12"],
    "13-16": ["13", "14", "15", "16"],
}

var_labels = {
    temp_col:  "Temp",
    light_col: "Light",
    heat_col:  "Heat",
    humid_col: "Humid",
    dew_col:   "Dew",
}

os.makedirs("outputs", exist_ok=True)

tables = {}

#For loop to repeat for each group of pocketlabs (4 pocketlabs per group)
for group_name, labs in groups.items():

    g = data[data["recording"].isin(labs)]

    result = pd.DataFrame(index=labs)

    for col in vars_all:
        if col is None:
            continue

        #wide table: rows = elapsed second, columns = each lab's reading
        wide = g.pivot_table(index="sec", columns="recording", values=col)

        #Getting the average each second
        averageSec = wide.mean(axis=1)

        #Getting the extend of the difference for each Pocketlab through Root Mean Square Error.
        gap = wide.sub(averageSec, axis=0)
        rmse = np.sqrt((gap ** 2).mean())

        result[var_labels[col]] = rmse

    result = result.reindex(labs).round(2)

    print(f"PocketLabs {group_name}: RMSE deviation from group average")
    print(result)
    print("Biggest deviator per variable:")
    print(result.idxmax().to_string())

    tables[group_name] = result

#Designing the tables shown
tfig, taxes = plt.subplots(2, 2, figsize=(14, 9))

for ax, (group_name, result) in zip(taxes.flat, tables.items()):

    ax.set_axis_off()
    #pad keeps the title close to the table instead of floating above it
    ax.set_title(f"PocketLabs {group_name}", fontsize=11, pad=4)

    #Rescaling the values from 0 to 1 so that the number can be corresponded to the intensity of the color in the table
    norm = (result - result.min()) / (result.max() - result.min())
    norm = norm.fillna(0)

    #Building the table
    tbl = ax.table(
        cellText=result.values,
        rowLabels=[f"{i}" for i in result.index],
        colLabels=result.columns,
        cellLoc="center",
        loc="upper center",
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(10)
    tbl.scale(1, 1.6)

    #Using red to shade the table and setting up intensity as well
    cmap = plt.get_cmap("Reds")
    for (r, c), cell in tbl.get_celld().items():
        if r == 0 or c == -1:
            cell.set_text_props(fontweight="bold")
            continue
        shade = norm.iloc[r - 1, c]
        cell.set_facecolor(cmap(0.15 + 0.7 * shade))
        if shade > 0.6:
            cell.set_text_props(color="white")

tfig.suptitle("PocketLab agreement by group (darker = bigger deviation)", fontsize=14, y=0.99)

#Big explanation
main_note = ("First, we group the Pocketlabs into 4 groups of 4 based on which pocketlabs we tested together. Then, we calculated the group's average for every second. For each second, we calculated the difference between the group average and the indiviudal pocketlab's reading. Then, we calcualted the Root Mean Square Error (RMSE) for each pocketlab by squaring the difference, averaging them, and then square rooting the result. \n"
             "Lower = tracks the group; darker red = bigger deviation (the outlier).")

#wrap the lines so we can see all of them
main_note = "\n".join(textwrap.fill(line, width=110) for line in main_note.split("\n"))

#Latex to show the formula
formula = r"$\mathrm{RMSE}=\sqrt{\dfrac{1}{T}\sum_{t=1}^{T}\left(x_t-\bar{x}_t\right)^2}$   (T = total seconds in group)"

#Stops overlapping (to an extent tbh)
tfig.text(0.5, 0.965, main_note + "\n" + formula,
          ha="center", va="top", fontsize=9.5, style="italic", color="#444444")

#Just combatting spacing issues
tfig.tight_layout(rect=[0, 0, 1, 0.80])
tfig.savefig("outputs/Test_similarity_tables.png", dpi=300, bbox_inches="tight")

plt.show()
