#import libraries
import pandas as pd #working with data tables
import glob #search for files
import os #work with file paths
import matplotlib.pyplot as plt #plot graphs
import matplotlib.dates as mdates #formatting dates and times
import numpy as np #math library (for NaN)

#input route number
route = input("Enter route (e.g. R31): ").upper()

#find subfolder in raw data for folder name that ends in the route number
folder = glob.glob(f"raw_data/*_{route}")[0]

#lists every .csv file inside the route folder
files = glob.glob(os.path.join(folder, "*.csv"))

dfs = []

#iterate over all files in the route folder
for f in files:

    #get filename
    name = os.path.basename(f)

    #skip files that don't have trial numbers (aren't pocketlab files)
    if "_T" not in name:
        continue

    #read csv into a DataFrame and ignores bad lines
    df = pd.read_csv(f, engine="python", on_bad_lines="skip")
    
    #formats file name into a list
    parts = name.replace(".csv", "").split("_")
    
    #skips files that aren't in the right format of DATE_ROUTE_INITALS_TRIAL
    if len(parts) < 4:
        continue

    #extracts person's name
    df["person"] = parts[2]

    #add DataFrame to list
    dfs.append(df)

#stack DataFrames into table and resets row numbers
data = pd.concat(dfs, ignore_index=True)

#remove accidental leading/trailing spaces from column names
data.columns = [c.strip() for c in data.columns]

#combine data and time into one string and parses into a datetime
data["datetime"] = pd.to_datetime(
    data["Date"].astype(str) + " " + data["Time"].astype(str),
    errors="coerce"
)

#remove rows where datetime or person is missing
data = data.dropna(subset=["datetime", "person"])

#sorts table chronologically
data = data.sort_values("datetime")

#function that searches a column by keyword
def find_col(df, keys):
    for c in df.columns:
        if any(k in c.lower() for k in keys):
            return c
    return None

#runs the search for all five variables
temp_col  = find_col(data, ["temp"])
light_col = find_col(data, ["light"])
heat_col  = find_col(data, ["heat"])
humid_col = find_col(data, ["humid"])
dew_col   = find_col(data, ["dew"])
vars_all = [temp_col, light_col, heat_col, humid_col, dew_col]

#function to extract number from string column
def to_num(s):
    return pd.to_numeric(
        s.astype(str).str.extract(r"([-+]?\d*\.?\d+)")[0],
        errors="coerce"
    )

#run string->number conversion on every variable column
for c in vars_all:
    if c:
        data[c] = to_num(data[c])

#eliminate outliers outside of +-3 standard deviations
for p in data["person"].unique():
    mask = data["person"] == p
    for c in vars_all:
        if c:

            #calculate mean and standard deviation for person's values in the column
            m = data.loc[mask, c].mean()
            s = data.loc[mask, c].std()

            #values that are more than 3 standard deviations above or below the mean get replaced with NaN
            data.loc[mask & (data[c] > m + 3*s), c] = np.nan
            data.loc[mask & (data[c] < m - 3*s), c] = np.nan

#function that inserts NaN breaks when there is a gap in data
def build_complete_series(df, person, value_col):

    #filters to a copy of a person's rows
    sub = df[df["person"] == person].copy()

    #sorts rows by time and replaces row numbers with datetime values as index
    sub = sub.sort_values("datetime").set_index("datetime")

    if len(sub) < 2:
        return sub

    #calculate time difference between every row
    diffs = sub.index.to_series().diff().dropna()
    
    #keep gaps that are longer than 1.5 seconds
    gap_threshold = pd.Timedelta(seconds=1.5)  # hardcoded since we know rate is 1/sec

    #marks if gaps are outside of threshold
    gap_mask = diffs > gap_threshold

    #keeps rows outside of the threshold
    gap_starts = diffs[gap_mask].index

    #no gaps
    if len(gap_starts) == 0:
        return sub

    #create a NaN row for each gap
    nan_rows = pd.DataFrame(
        {c: np.nan for c in sub.columns},
        index=gap_starts - pd.Timedelta(milliseconds=1)
    )
    nan_rows["person"] = person

    #combine real data and NaN rows into a table and sorts by timestamp
    result = pd.concat([sub, nan_rows]).sort_index()
    return result

#sort people names in alphabetical order
people = sorted(data["person"].unique())

#cycle matplotlib default colors and assign a color to each person
default_colors = plt.rcParams["axes.prop_cycle"].by_key()["color"]
colors = {p: default_colors[i % len(default_colors)] for i, p in enumerate(people)}

#creates a 5 by 2 grid of subplots
fig, axes = plt.subplots(5, 2, figsize=(18, 18))

#create axis titles
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

            #build time series with gaps
            series = build_complete_series(data, p, col)

            #convert datetime index to decimal number for x axis
            hour = series.index.hour + series.index.minute / 60
            
            #separate AM data and PM data (cutoff at 13.5 for 1:30PM)
            if period == "AM":
                series = series[hour < 13.5]
            else:
                series = series[hour >= 13.5]

            #draw line on subplot
            ax.plot(
                series.index,
                series[col],
                color=colors[p],
                linewidth=1.5
            )

        #set subplot title
        ax.set_title(f"{title} ({period})", fontsize=9)

        #label y axis 
        ax.set_ylabel(title, fontsize=8)

        #add grid in background
        ax.grid(True, alpha=0.3)

        #formats x axis as hour:minute
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))
        
        #changes x and y axis rotation and font size
        ax.tick_params(axis="x", rotation=45, labelsize=7)
        ax.tick_params(axis="y", labelsize=7)

#creates one legend for respresenting people for whole figure
handles = [plt.Line2D([0], [0], color=colors[p], lw=2) for p in people]
labels  = list(people)
fig.legend(handles, labels, loc="upper right", title="Person")

#creates outputs folder if it doesn't exist
os.makedirs("outputs", exist_ok=True)

#adjust spacing between subplots
plt.tight_layout()

#save figure as a png
plt.savefig(
    f"outputs/{route}_AM_PM_multiplot.png",
    dpi=300,
    bbox_inches="tight"
)

#display figure as a window
plt.show()