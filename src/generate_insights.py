"""
Melbourne Housing Insights Generator

Generates investment insights and summary statistics from the analysis dataset.
"""

import pandas as pd
import numpy as np
from pathlib import Path


def suburb_comparison(df: pd.DataFrame) -> pd.DataFrame:
    """Generate suburb summary statistics."""
    stats = df.groupby("Suburb").agg({
        "Price": ["median", "mean", "count"],
        "Price_per_sqm": "median",
        "Rooms": "median",
        "Distance": "first"
    }).round(0)

    stats.columns = [
        "Median_Price", "Mean_Price", "Transaction_Count",
        "Median_Price_per_sqm", "Median_Rooms", "Distance_to_CBD"
    ]
    return stats.reset_index().sort_values("Median_Price", ascending=False)


def quarterly_trends(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate quarterly price trends."""
    stats = df.groupby("Quarter").agg({
        "Price": ["median", "mean", "count"],
        "Price_per_sqm": "median"
    }).round(0)

    stats.columns = ["Median_Price", "Mean_Price", "Transaction_Count", "Median_Price_per_sqm"]
    return stats.reset_index()


def property_type_premium(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate house vs unit premium by suburb."""
    pivot = df.groupby(["Suburb", "PropertyType"])["Price"].median().unstack()
    pivot["House_vs_Unit_Premium"] = (
        (pivot["House"] - pivot["Unit"]) / pivot["Unit"] * 100
    ).round(0)
    return pivot.dropna(subset=["House", "Unit"]).sort_values(
        "House_vs_Unit_Premium", ascending=False
    )


def bedroom_premium(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate 4BR vs 3BR premium for houses."""
    houses = df[df["PropertyType"] == "House"]
    pivot = houses.groupby(["Suburb", "Rooms"])["Price"].median().unstack()

    if 3 in pivot.columns and 4 in pivot.columns:
        pivot["4BR_vs_3BR_Premium"] = (
            (pivot[4] - pivot[3]) / pivot[3] * 100
        ).round(0)
        return pivot.dropna(subset=[3, 4]).sort_values(
            "4BR_vs_3BR_Premium", ascending=False
        )
    return pd.DataFrame()


def value_suburb_gaps(df: pd.DataFrame) -> dict:
    """Identify undervalued suburbs compared to neighbors."""
    suburb_stats = df.groupby("Suburb")["Price"].median()

    comparisons = {
        "Reservoir_vs_Northcote": {
            "value_suburb": "Reservoir",
            "premium_suburb": "Northcote",
            "value_price": suburb_stats.get("Reservoir", 0),
            "premium_price": suburb_stats.get("Northcote", 0),
        },
        "Glenroy_vs_MooneePonds": {
            "value_suburb": "Glenroy",
            "premium_suburb": "Moonee Ponds",
            "value_price": suburb_stats.get("Glenroy", 0),
            "premium_price": suburb_stats.get("Moonee Ponds", 0),
        },
        "Coburg_vs_Brunswick": {
            "value_suburb": "Coburg",
            "premium_suburb": "Brunswick",
            "value_price": suburb_stats.get("Coburg", 0),
            "premium_price": suburb_stats.get("Brunswick", 0),
        }
    }

    for key, comp in comparisons.items():
        if comp["premium_price"] > 0:
            comp["discount_pct"] = round(
                (comp["premium_price"] - comp["value_price"]) / comp["premium_price"] * 100
            )

    return comparisons


def main():
    """Main execution function."""
    project_root = Path(__file__).parent.parent
    data_dir = project_root / "data"

    # Load analysis data
    print("Loading analysis data...")
    df = pd.read_csv(data_dir / "processed" / "melb_data_analysis.csv")
    print(f"  Loaded {len(df):,} records")

    # Generate insights
    print("\n" + "=" * 60)
    print("INVESTMENT INSIGHTS")
    print("=" * 60)

    # Value suburbs
    print("\n1. UNDERVALUED SUBURBS")
    gaps = value_suburb_gaps(df)
    for comp in gaps.values():
        print(f"   {comp['value_suburb']}: ${comp['value_price']:,.0f} "
              f"({comp.get('discount_pct', 0)}% below {comp['premium_suburb']})")

    # Property type premiums
    print("\n2. HOUSE VS UNIT PREMIUMS (Top 5)")
    type_premium = property_type_premium(df)
    for suburb in type_premium.head(5).index:
        premium = type_premium.loc[suburb, "House_vs_Unit_Premium"]
        print(f"   {suburb}: {premium:.0f}% premium")

    # Bedroom premiums
    print("\n3. 4BR VS 3BR PREMIUMS (Top 5)")
    br_premium = bedroom_premium(df)
    for suburb in br_premium.head(5).index:
        premium = br_premium.loc[suburb, "4BR_vs_3BR_Premium"]
        print(f"   {suburb}: {premium:.0f}% premium")

    # Quarterly trends
    print("\n4. MARKET TREND")
    quarterly = quarterly_trends(df)
    q2_2016 = quarterly[quarterly["Quarter"] == "2016Q2"]["Median_Price"].values[0]
    q3_2017 = quarterly[quarterly["Quarter"] == "2017Q3"]["Median_Price"].values[0]
    growth = (q3_2017 - q2_2016) / q2_2016 * 100
    print(f"   Q2 2016 -> Q3 2017: +{growth:.0f}% market growth")
    print(f"   ${q2_2016:,.0f} -> ${q3_2017:,.0f}")


if __name__ == "__main__":
    main()
