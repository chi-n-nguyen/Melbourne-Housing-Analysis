"""
Melbourne Housing Visualisation Generator

Generates the five analysis charts saved to images/.
Run after feature_engineering.py.

Charts produced:
    images/quarterly_trends.png          - Quarterly median price trend (line)
    images/median_price_by_suburb.png    - Median price per suburb (horizontal bar)
    images/value_suburbs.png             - Undervalued suburb price gaps (grouped bar)
    images/property_type_distribution.png - House / Unit / Townhouse breakdown (bar)
    images/price_vs_rooms.png            - Price by bedroom count (box plot)

Usage:
    python3 src/generate_visualisations.py
"""

import matplotlib
matplotlib.use("Agg")  # headless rendering -no display required

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import pandas as pd
from pathlib import Path


# ── Style constants ──────────────────────────────────────────────────────────
PRIMARY   = "#1F4E79"
SECONDARY = "#2E75B6"
ACCENT    = "#70AD47"
WARNING   = "#ED7D31"
GRID_COL  = "#E0E0E0"
DPI       = 150

plt.rcParams.update({
    "font.family":       "sans-serif",
    "axes.spines.top":   False,
    "axes.spines.right": False,
    "axes.grid":         True,
    "grid.color":        GRID_COL,
    "grid.linewidth":    0.6,
    "figure.dpi":        DPI,
})


def _dollar_fmt(x, _):
    """Format axis tick as $NNNk or $N.NM."""
    if x >= 1_000_000:
        return f"${x/1_000_000:.1f}M"
    return f"${x/1_000:.0f}k"


# ── Chart 1: Quarterly median price trend ───────────────────────────────────

def plot_quarterly_trends(df: pd.DataFrame, out_path: Path) -> None:
    quarterly = (
        df.groupby("Quarter")["Price"]
        .median()
        .reset_index()
        .sort_values("Quarter")
    )

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.plot(
        quarterly["Quarter"], quarterly["Price"],
        color=PRIMARY, linewidth=2.5, marker="o", markersize=6, zorder=3,
    )
    ax.fill_between(quarterly["Quarter"], quarterly["Price"],
                    alpha=0.12, color=PRIMARY)

    # Annotate first and last quarters
    for i in [0, -1]:
        row = quarterly.iloc[i]
        ax.annotate(
            f"${row['Price']/1000:.0f}k",
            xy=(row["Quarter"], row["Price"]),
            xytext=(0, 12), textcoords="offset points",
            ha="center", fontsize=9, color=PRIMARY, fontweight="bold",
        )

    ax.set_title("Quarterly Median Sale Price -Melbourne Top 20 Suburbs",
                 fontsize=13, fontweight="bold", color=PRIMARY, pad=12)
    ax.set_xlabel("Quarter", fontsize=10)
    ax.set_ylabel("Median Price", fontsize=10)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_dollar_fmt))
    ax.tick_params(axis="x", rotation=45)

    fig.tight_layout()
    fig.savefig(out_path)
    plt.close(fig)


# ── Chart 2: Median price by suburb ─────────────────────────────────────────

def plot_median_price_by_suburb(df: pd.DataFrame, out_path: Path) -> None:
    suburb_median = (
        df.groupby("Suburb")["Price"]
        .median()
        .sort_values(ascending=True)
    )

    fig, ax = plt.subplots(figsize=(10, 7))
    bars = ax.barh(suburb_median.index, suburb_median.values, color=SECONDARY, height=0.7)

    # Highlight the three value suburbs
    VALUE_SUBURBS = {"Reservoir", "Glenroy", "Coburg"}
    for bar, suburb in zip(bars, suburb_median.index):
        if suburb in VALUE_SUBURBS:
            bar.set_color(ACCENT)

    market_median = df["Price"].median()
    ax.axvline(market_median, color=WARNING, linewidth=1.5, linestyle="--",
               label=f"Market median  ${market_median/1000:.0f}k")

    ax.set_title("Median Sale Price by Suburb (Top 20)",
                 fontsize=13, fontweight="bold", color=PRIMARY, pad=12)
    ax.set_xlabel("Median Price", fontsize=10)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(_dollar_fmt))
    ax.legend(fontsize=9)

    # Small note for value suburbs
    ax.annotate("Green = identified value suburb",
                xy=(0.02, 0.02), xycoords="axes fraction",
                fontsize=8, color=ACCENT)

    fig.tight_layout()
    fig.savefig(out_path)
    plt.close(fig)


# ── Chart 3: Undervalued suburb price gaps ───────────────────────────────────

def plot_value_suburbs(df: pd.DataFrame, out_path: Path) -> None:
    pairs = [
        ("Reservoir",  "Northcote"),
        ("Glenroy",    "Moonee Ponds"),
        ("Coburg",     "Brunswick"),
    ]
    medians = df.groupby("Suburb")["Price"].median()

    labels, value_prices, premium_prices, discounts = [], [], [], []
    for value_sub, premium_sub in pairs:
        v = medians.get(value_sub, 0)
        p = medians.get(premium_sub, 0)
        labels.append(f"{value_sub}\nvs {premium_sub}")
        value_prices.append(v)
        premium_prices.append(p)
        discounts.append(round((p - v) / p * 100) if p > 0 else 0)

    x = range(len(labels))
    width = 0.35

    fig, ax = plt.subplots(figsize=(9, 5))
    b1 = ax.bar([i - width/2 for i in x], value_prices,  width, label="Value suburb",   color=ACCENT)
    b2 = ax.bar([i + width/2 for i in x], premium_prices, width, label="Premium suburb", color=PRIMARY)

    # Annotate discount percentages
    for i, disc in enumerate(discounts):
        ax.annotate(
            f"{disc}% discount",
            xy=(i, max(value_prices[i], premium_prices[i]) + 15_000),
            ha="center", fontsize=9, color=WARNING, fontweight="bold",
        )

    ax.set_title("Undervalued Suburbs vs Adjacent Premium Suburbs",
                 fontsize=13, fontweight="bold", color=PRIMARY, pad=12)
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, fontsize=10)
    ax.set_ylabel("Median Price", fontsize=10)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_dollar_fmt))
    ax.legend(fontsize=9)

    fig.tight_layout()
    fig.savefig(out_path)
    plt.close(fig)


# ── Chart 4: Property type distribution ─────────────────────────────────────

def plot_property_type_distribution(df: pd.DataFrame, out_path: Path) -> None:
    type_stats = (
        df.groupby("PropertyType")["Price"]
        .agg(["median", "count"])
        .reset_index()
        .sort_values("median", ascending=False)
    )

    fig, (ax_bar, ax_count) = plt.subplots(1, 2, figsize=(11, 5))

    # Left: median price by type
    colours = [PRIMARY, SECONDARY, ACCENT]
    ax_bar.bar(type_stats["PropertyType"], type_stats["median"],
               color=colours[:len(type_stats)], width=0.5)
    ax_bar.set_title("Median Price by Property Type", fontsize=11,
                     fontweight="bold", color=PRIMARY)
    ax_bar.set_ylabel("Median Price", fontsize=10)
    ax_bar.yaxis.set_major_formatter(mticker.FuncFormatter(_dollar_fmt))

    for i, row in type_stats.iterrows():
        ax_bar.text(
            list(type_stats["PropertyType"]).index(row["PropertyType"]),
            row["median"] + 10_000,
            f"${row['median']/1000:.0f}k",
            ha="center", fontsize=9, fontweight="bold",
        )

    # Right: transaction count by type
    ax_count.bar(type_stats["PropertyType"], type_stats["count"],
                 color=colours[:len(type_stats)], width=0.5)
    ax_count.set_title("Transaction Count by Property Type", fontsize=11,
                       fontweight="bold", color=PRIMARY)
    ax_count.set_ylabel("Transactions", fontsize=10)

    for i, row in type_stats.iterrows():
        ax_count.text(
            list(type_stats["PropertyType"]).index(row["PropertyType"]),
            row["count"] + 5,
            f"{int(row['count']):,}",
            ha="center", fontsize=9, fontweight="bold",
        )

    fig.suptitle("Property Type Analysis -Top 20 Suburbs",
                 fontsize=13, fontweight="bold", color=PRIMARY, y=1.01)
    fig.tight_layout()
    fig.savefig(out_path, bbox_inches="tight")
    plt.close(fig)


# ── Chart 5: Price vs bedrooms (box plot) ────────────────────────────────────

def plot_price_vs_rooms(df: pd.DataFrame, out_path: Path) -> None:
    houses = df[df["PropertyType"] == "House"]
    room_counts = sorted(houses["Rooms"].unique())
    # Keep only bedroom counts with enough data to box-plot meaningfully
    room_counts = [r for r in room_counts if (houses["Rooms"] == r).sum() >= 10]

    data_by_room = [houses.loc[houses["Rooms"] == r, "Price"].values for r in room_counts]

    fig, ax = plt.subplots(figsize=(10, 5))
    bp = ax.boxplot(
        data_by_room,
        tick_labels=[f"{int(r)} BR" for r in room_counts],
        patch_artist=True,
        medianprops=dict(color=WARNING, linewidth=2),
        whiskerprops=dict(color=PRIMARY),
        capprops=dict(color=PRIMARY),
        flierprops=dict(marker="o", markerfacecolor=GRID_COL,
                        markersize=3, linestyle="none"),
    )
    for patch in bp["boxes"]:
        patch.set_facecolor(SECONDARY)
        patch.set_alpha(0.6)

    ax.set_title("House Price Distribution by Bedroom Count -Top 20 Suburbs",
                 fontsize=13, fontweight="bold", color=PRIMARY, pad=12)
    ax.set_xlabel("Bedrooms", fontsize=10)
    ax.set_ylabel("Sale Price", fontsize=10)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_dollar_fmt))

    fig.tight_layout()
    fig.savefig(out_path)
    plt.close(fig)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    project_root = Path(__file__).parent.parent
    data_dir     = project_root / "data"
    images_dir   = project_root / "images"
    images_dir.mkdir(exist_ok=True)

    print("Loading analysis data...")
    df = pd.read_csv(data_dir / "processed" / "melb_data_analysis.csv")
    df["Date"] = pd.to_datetime(df["Date"])
    print(f"  Loaded {len(df):,} records across {df['Suburb'].nunique()} suburbs")

    charts = [
        ("Quarterly trends",         plot_quarterly_trends,           "quarterly_trends.png"),
        ("Median price by suburb",   plot_median_price_by_suburb,     "median_price_by_suburb.png"),
        ("Value suburb gaps",        plot_value_suburbs,               "value_suburbs.png"),
        ("Property type distribution", plot_property_type_distribution, "property_type_distribution.png"),
        ("Price vs bedrooms",        plot_price_vs_rooms,             "price_vs_rooms.png"),
    ]

    print("\nGenerating charts...")
    for name, fn, filename in charts:
        out = images_dir / filename
        fn(df, out)
        print(f"  {filename:<40} saved")

    print(f"\n5 charts written to: {images_dir}")


if __name__ == "__main__":
    main()
