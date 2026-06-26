"""
compare_sensors.py
------------------
Compares PocketLab (1-19) vs Kestrel (FIRE, HEAT) across four metrics.

Changes vs prior version:
  - Kestrel lines only drawn during time windows when pocketlabs are active
  - Offset table below each graph: one column per PocketLab, avg difference vs Kestrel mean

HOW TO USE:
  1.  pip install pandas matplotlib openpyxl
  2.  Edit the two file paths below
  3.  python compare_sensors.py
"""

# ════════════════════════════════════════════════════════════════
#  INPUT YOUR FILE PATHS HERE
# ════════════════════════════════════════════════════════════════
POCKETLAB_FILE = r"C:\Users\dresd\Downloads\0618_R1000\COMBINED_0618_R1000.xlsx"   # ← path to your PocketLab Excel file
KESTREL_FILE   = r"C:\Users\dresd\Downloads\0618_R1000\COMBINED_KESTREL_0618_R1000.xlsx"     # ← path to your Kestrel Excel file
# ════════════════════════════════════════════════════════════════

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.gridspec as gridspec
from pathlib import Path

# ── Config ────────────────────────────────────────────────────────────────────
TIME_START        = "14:18:00"
TIME_END          = "17:18:00"
GAP_THRESHOLD_MIN = 5      # minutes — gaps larger than this break the kestrel line

KESTREL_LW   = 2.5
POCKETLAB_LW = 1.2
FIG_WIDTH    = 20
PLOT_H       = 5
TABLE_H      = 1.8

KESTREL_STYLES = {
    "FIRE": dict(color="black", lw=KESTREL_LW, ls="-",  zorder=10),
    "HEAT": dict(color="black", lw=KESTREL_LW, ls="--", zorder=10),
}

METRICS = [
    ("heat_index",  "Heat Index (°C)"),
    ("dew_point",   "Dew Point (°C)"),
    ("temperature", "Temperature (°C)"),
    ("humidity",    "Humidity / Relative Humidity (%)"),
]

# ── Helpers ───────────────────────────────────────────────────────────────────

def f_to_c(s):
    return (s - 32) * 5 / 9

def read_file(path):
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(p, engine="openpyxl")
    return pd.read_csv(p)

# ── Loaders ───────────────────────────────────────────────────────────────────

def load_pocketlab(path):
    df = read_file(path)
    df.columns = df.columns.str.strip()
    df["datetime"] = pd.to_datetime(
        df["Date"].astype(str) + " " + df["Time"].astype(str)
    )
    df = df.rename(columns={
        "pocketlab_number": "device",
        "Heat Index":        "heat_index",
        "Dew Point":         "dew_point",
        "Temperature Probe": "temperature",
        "Humidity":          "humidity",
    })[["device", "datetime", "heat_index", "dew_point", "temperature", "humidity"]]
    df["device"] = df["device"].astype(str)
    for col in ("heat_index", "dew_point", "temperature", "humidity"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df

def load_kestrel(path):
    df = read_file(path)
    df.columns = df.columns.str.strip()
    df["datetime"] = pd.to_datetime(
        df["Date"].astype(str) + " " + df["Time"].astype(str)
    )
    df = df.rename(columns={
        "kestrel_name":      "device",
        "Heat Index":        "heat_index",
        "Dew Point":         "dew_point",
        "Temperature":       "temperature",
        "Relative Humidity": "humidity",
    })[["device", "datetime", "heat_index", "dew_point", "temperature", "humidity"]]
    df["device"] = df["device"].astype(str)
    for col in ("heat_index", "dew_point", "temperature"):
        df[col] = f_to_c(pd.to_numeric(df[col], errors="coerce"))
    df["humidity"] = pd.to_numeric(df["humidity"], errors="coerce")
    return df

def filter_window(df, ref_date):
    s = pd.Timestamp(f"{ref_date} {TIME_START}")
    e = pd.Timestamp(f"{ref_date} {TIME_END}")
    return df[(df["datetime"] >= s) & (df["datetime"] <= e)].copy()

# ── Active window detection ───────────────────────────────────────────────────

def get_active_windows(pl_df, gap_min=GAP_THRESHOLD_MIN):
    """
    Find contiguous time windows where any pocketlab has data.
    Returns list of (start, end) Timestamps.
    """
    all_times = np.sort(pl_df["datetime"].unique())
    if len(all_times) == 0:
        return []
    windows, win_start, prev = [], all_times[0], all_times[0]
    for t in all_times[1:]:
        if (t - prev) / np.timedelta64(1, "m") > gap_min:
            windows.append((pd.Timestamp(win_start), pd.Timestamp(prev)))
            win_start = t
        prev = t
    windows.append((pd.Timestamp(win_start), pd.Timestamp(prev)))
    return windows

# ── Offset computation ────────────────────────────────────────────────────────

def build_kestrel_avg(k_df, metric_col):
    """
    Average FIRE and HEAT into one reference time series.
    Uses merge_asof so the two kestrels don't need identical timestamps.
    """
    fire = (k_df[k_df["device"] == "FIRE"][["datetime", metric_col]]
            .sort_values("datetime")
            .rename(columns={metric_col: "fire"}))
    heat = (k_df[k_df["device"] == "HEAT"][["datetime", metric_col]]
            .sort_values("datetime")
            .rename(columns={metric_col: "heat"}))

    if fire.empty and heat.empty:
        return pd.DataFrame(columns=["datetime", "kestrel_avg"])
    if fire.empty:
        return heat.rename(columns={"heat": "kestrel_avg"})
    if heat.empty:
        return fire.rename(columns={"fire": "kestrel_avg"})

    merged = pd.merge_asof(fire, heat, on="datetime",
                           tolerance=pd.Timedelta("2min"), direction="nearest")
    merged["kestrel_avg"] = merged[["fire", "heat"]].mean(axis=1)
    return merged[["datetime", "kestrel_avg"]]

def compute_offsets(pl_df, k_avg_df, metric_col, pl_devices):
    """
    HOW THE OFFSET IS CALCULATED
    ─────────────────────────────
    1. Kestrel average: FIRE and HEAT are merged on their timestamps
       (nearest match within 2 minutes) and their values are averaged
       into a single reference series — see build_kestrel_avg().

    2. For each PocketLab reading, we find the closest kestrel-average
       timestamp (within a 5-minute tolerance) using pd.merge_asof with
       direction="nearest".

    3. The per-reading difference is:
           diff = PocketLab_value − Kestrel_average

       A positive value means the PocketLab reads HIGHER than the Kestrels.
       A negative value means the PocketLab reads LOWER than the Kestrels.

    4. All per-reading differences for a given PocketLab are averaged
       (mean) to produce a single summary offset for that device.

    Units: °C for Heat Index / Dew Point / Temperature; % for Humidity.
    """
    offsets = {}
    k_sorted = k_avg_df.sort_values("datetime")
    for dev in pl_devices:
        sub = (pl_df[pl_df["device"] == dev][["datetime", metric_col]]
               .sort_values("datetime")
               .dropna(subset=[metric_col]))
        if sub.empty or k_sorted.empty:
            offsets[dev] = np.nan
            continue
        merged = pd.merge_asof(sub, k_sorted, on="datetime",
                               tolerance=pd.Timedelta("5min"), direction="nearest")
        merged = merged.dropna(subset=[metric_col, "kestrel_avg"])
        offsets[dev] = (merged[metric_col] - merged["kestrel_avg"]).mean() if not merged.empty else np.nan
    return offsets

# ── Plot helpers ──────────────────────────────────────────────────────────────

def plot_metric_ax(ax, metric_col, ylabel,
                   pl_df, k_df, pl_devices, pl_colors,
                   time_start, time_end, windows):
    """Draw the time-series panel."""

    # Kestrel lines drawn FIRST so they sit beneath all PocketLab lines
    for kname, style in KESTREL_STYLES.items():
        sub = k_df[k_df["device"] == kname].sort_values("datetime")
        first_seg = True
        for (ws, we) in windows:
            seg = sub[(sub["datetime"] >= ws) & (sub["datetime"] <= we)]
            if seg.empty:
                continue
            lbl = f"Kestrel {kname}" if first_seg else "_nolegend_"
            ax.plot(seg["datetime"], seg[metric_col], **{**style, "zorder": 2}, label=lbl)
            first_seg = False

    # PocketLab lines drawn ON TOP of kestrels
    for dev, color in zip(pl_devices, pl_colors):
        sub = pl_df[pl_df["device"] == dev].sort_values("datetime")
        if sub.empty:
            continue
        ax.plot(sub["datetime"], sub[metric_col],
                color=color, lw=POCKETLAB_LW, alpha=0.85, zorder=3, label=f"PL {dev}")

    ax.set_xlim(time_start, time_end)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%I:%M %p"))
    ax.xaxis.set_major_locator(mdates.MinuteLocator(byminute=[0, 15, 30, 45]))
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=30, ha="right", fontsize=8)
    ax.set_ylabel(ylabel, fontsize=10)
    ax.set_xlabel("Time", fontsize=9)
    ax.grid(True, alpha=0.3, linestyle=":")

    handles, labels = ax.get_legend_handles_labels()
    k_idx = [i for i, l in enumerate(labels) if l.startswith("Kestrel")]
    p_idx = [i for i, l in enumerate(labels) if not l.startswith("Kestrel")]
    ax.legend(
        [handles[i] for i in k_idx + p_idx],
        [labels[i]  for i in k_idx + p_idx],
        loc="upper right", fontsize=7, ncol=5, framealpha=0.85,
    )


def draw_offset_table(ax, offsets, pl_devices, metric_col):
    """
    Draw a horizontal table: columns = PL numbers, one row = avg offset.
    Cells are tinted red (PL reads high) or blue (PL reads low).
    """
    ax.axis("off")
    unit = "%" if metric_col == "humidity" else "°C"

    col_labels = [f"PL {d}" for d in pl_devices]
    values     = [offsets.get(d, np.nan) for d in pl_devices]

    # Build cell text and colours
    cell_text   = [[]]
    cell_colors = [[]]

    for val in values:
        if np.isnan(val):
            cell_text[0].append("N/A")
            cell_colors[0].append("#f0f0f0")
        else:
            cell_text[0].append(f"{val:+.2f}")
            intensity = min(abs(val) / 5.0, 0.45)
            if val > 0:
                r, g, b = 1.0, 1.0 - intensity, 1.0 - intensity   # red tint
            else:
                r, g, b = 1.0 - intensity, 1.0 - intensity, 1.0   # blue tint
            cell_colors[0].append((r, g, b))

    row_labels = [f"Avg Offset ({unit})\nvs Kestrel Mean"]

    tbl = ax.table(
        cellText=cell_text,
        cellColours=cell_colors,
        colLabels=col_labels,
        rowLabels=row_labels,
        loc="center",
        cellLoc="center",
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8)
    tbl.scale(1, 1.8)

    # Bold the column headers
    for (row, col), cell in tbl.get_celld().items():
        if row == 0:
            cell.set_text_props(fontweight="bold")

# ── Main ──────────────────────────────────────────────────────────────────────


def export_offsets_xlsx(metric_offsets, pl_devices, ref_date, outdir):
    """
    Write a spreadsheet (offsets.xlsx) with one sheet per metric.
    Each sheet has:
      - Column A: PocketLab number
      - Column B: Average offset vs Kestrel mean
    """
    units = {
        "heat_index":  "°C",
        "dew_point":   "°C",
        "temperature": "°C",
        "humidity":    "%",
    }
    metric_labels = {
        "heat_index":  "Heat Index",
        "dew_point":   "Dew Point",
        "temperature": "Temperature",
        "humidity":    "Humidity",
    }

    xlsx_path = outdir / "offsets.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:

        # ── Summary sheet: all metrics side by side ───────────────────
        rows = []
        for dev in pl_devices:
            row = {"PocketLab #": f"PL {dev}"}
            for metric_col, label in metric_labels.items():
                unit = units[metric_col]
                val  = metric_offsets[metric_col].get(dev, np.nan)
                row[f"{label} Avg Offset ({unit})"] = round(val, 3) if not np.isnan(val) else ""
            rows.append(row)

        summary_df = pd.DataFrame(rows)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        # ── One sheet per metric ──────────────────────────────────────
        for metric_col, label in metric_labels.items():
            unit = units[metric_col]
            offsets = metric_offsets[metric_col]
            sheet_rows = []
            for dev in pl_devices:
                val = offsets.get(dev, np.nan)
                sheet_rows.append({
                    "PocketLab #": f"PL {dev}",
                    f"Avg Offset ({unit}) vs Kestrel Mean": round(val, 3) if not np.isnan(val) else "",
                })
            pd.DataFrame(sheet_rows).to_excel(writer, sheet_name=label, index=False)

    return xlsx_path

def main():
    outdir = Path(__file__).parent

    print(f"Loading PocketLab: {POCKETLAB_FILE}")
    pl_df = load_pocketlab(POCKETLAB_FILE)

    print(f"Loading Kestrel:   {KESTREL_FILE}")
    k_df  = load_kestrel(KESTREL_FILE)

    ref_date   = pl_df["datetime"].dt.date.iloc[0].strftime("%Y-%m-%d")
    time_start = pd.Timestamp(f"{ref_date} {TIME_START}")
    time_end   = pd.Timestamp(f"{ref_date} {TIME_END}")
    print(f"Reference date: {ref_date}  |  Window: 2:18 PM – 5:18 PM")

    pl_df = filter_window(pl_df, ref_date)
    k_df  = filter_window(k_df,  ref_date)

    pl_devices = sorted(pl_df["device"].unique(),
                        key=lambda x: int(x) if x.isdigit() else x)
    pl_colors  = [plt.colormaps["tab20"](i / max(len(pl_devices), 1))
                  for i in range(len(pl_devices))]

    windows = get_active_windows(pl_df)
    print(f"\nActive pocketlab windows detected ({len(windows)}):")
    for ws, we in windows:
        print(f"  {ws.strftime('%I:%M %p')} – {we.strftime('%I:%M %p')}")

    print(f"\nPocketLabs : {pl_devices}")
    print(f"Kestrels   : {sorted(k_df['device'].unique().tolist())}")

    # Pre-compute kestrel averages and per-PL offsets for all metrics
    metric_offsets = {}
    for metric_col, _ in METRICS:
        k_avg = build_kestrel_avg(k_df, metric_col)
        metric_offsets[metric_col] = compute_offsets(pl_df, k_avg, metric_col, pl_devices)

    # ── Combined figure (plot + table for each of 4 metrics) ─────────────
    print("\nSaving combined_metrics.png …")
    n   = len(METRICS)
    fig = plt.figure(figsize=(FIG_WIDTH, (PLOT_H + TABLE_H) * n))
    gs  = gridspec.GridSpec(n * 2, 1,
                            height_ratios=[PLOT_H, TABLE_H] * n,
                            hspace=0.55)
    fig.suptitle(
        f"PocketLab vs Kestrel  —  {ref_date}  |  2:18 PM – 5:18 PM\n"
        f"Kestrel lines shown only while PocketLabs are recording",
        fontsize=13, fontweight="bold",
    )
    for i, (metric_col, ylabel) in enumerate(METRICS):
        ax_plot  = fig.add_subplot(gs[i * 2])
        ax_table = fig.add_subplot(gs[i * 2 + 1])
        plot_metric_ax(ax_plot, metric_col, ylabel,
                       pl_df, k_df, pl_devices, pl_colors,
                       time_start, time_end, windows)
        draw_offset_table(ax_table, metric_offsets[metric_col], pl_devices, metric_col)

    fig.savefig(outdir / "combined_metrics.png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    print("  → combined_metrics.png saved")

    # ── Individual figures ────────────────────────────────────────────────
    fnames = ["heat_index.png", "dew_point.png", "temperature.png", "humidity.png"]
    for (metric_col, ylabel), fname in zip(METRICS, fnames):
        print(f"Saving {fname} …")
        fig = plt.figure(figsize=(FIG_WIDTH, PLOT_H + TABLE_H))
        gs  = gridspec.GridSpec(2, 1,
                                height_ratios=[PLOT_H, TABLE_H],
                                hspace=0.45)
        fig.suptitle(
            f"{ylabel.split(' (')[0]}  —  PocketLab vs Kestrel\n"
            f"{ref_date}  |  2:18 PM – 5:18 PM",
            fontsize=13, fontweight="bold",
        )
        ax_plot  = fig.add_subplot(gs[0])
        ax_table = fig.add_subplot(gs[1])
        plot_metric_ax(ax_plot, metric_col, ylabel,
                       pl_df, k_df, pl_devices, pl_colors,
                       time_start, time_end, windows)
        draw_offset_table(ax_table, metric_offsets[metric_col], pl_devices, metric_col)
        fig.savefig(outdir / fname, dpi=150, bbox_inches="tight")
        plt.close(fig)
        print(f"  → {fname} saved")

    # ── Export offsets spreadsheet ───────────────────────────────────────
    print("Saving offsets.xlsx …")
    xlsx_path = export_offsets_xlsx(metric_offsets, pl_devices, ref_date, outdir)
    print(f"  → {xlsx_path.name} saved")

    print(f"\n✅ Done! All outputs saved to: {outdir.resolve()}")


if __name__ == "__main__":
    main()