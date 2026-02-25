"""
Melbourne Housing Power BI Export Pipeline

Exports data in a star-schema format optimised for Power BI ingestion.
Run this after feature_engineering.py.

Output directory: data/powerbi/

  Fact table:
    fact_transactions.csv         -one row per sale

  Dimension tables:
    dim_suburb.csv                -suburb metadata + investment ratings
    dim_date.csv                  -full date dimension for time intelligence
    dim_property_type.csv         -property type lookup

  Pre-aggregated summary tables:
    agg_quarterly_trends.csv      -quarterly medians, volumes, QoQ growth
    agg_suburb_summary.csv        -suburb-level KPIs vs market median
    agg_property_type_premium.csv -house-vs-unit premium by suburb
    agg_value_gaps.csv            -undervalued suburb comparisons

  DAX measures:
    powerbi_measures.dax          -copy-paste measures for Power BI

Usage:
    python src/create_powerbi_export.py

Power BI setup:
    1. Open Power BI Desktop → Get Data → Text/CSV
    2. Import all CSVs from data/powerbi/
    3. Set relationships:
         fact_transactions[Suburb]       → dim_suburb[Suburb]
         fact_transactions[Date]         → dim_date[Date]
         fact_transactions[PropertyType] → dim_property_type[PropertyType]
    4. Copy DAX measures from powerbi_measures.dax into new measures
    5. Build visuals (suggested pages below each table)
"""

import pandas as pd
import numpy as np
from pathlib import Path


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _investment_rating(pct_vs_median: float) -> str:
    """Map percentage deviation from market median to an investment rating."""
    if pct_vs_median < -0.30:
        return "VALUE BUY"
    elif pct_vs_median < 0:
        return "BELOW MARKET"
    elif pct_vs_median < 0.30:
        return "AT MARKET"
    else:
        return "PREMIUM"


def _region(distance_km: float) -> str:
    """Classify suburb into Melbourne distance zones."""
    if distance_km <= 10:
        return "Inner (0–10 km)"
    elif distance_km <= 20:
        return "Middle (10–20 km)"
    else:
        return "Outer (20+ km)"


# ---------------------------------------------------------------------------
# Fact table
# ---------------------------------------------------------------------------

def build_fact_transactions(df: pd.DataFrame) -> pd.DataFrame:
    """
    One row per property sale.

    Keeps the grain of the raw analysis dataset while exposing only the
    columns needed for Power BI measures and slicing.
    """
    fact_cols = [
        "Suburb",
        "Date",
        "PropertyType",
        "Price",
        "Rooms",
        "Bathroom",
        "Car",
        "Landsize",
        "Price_per_sqm",
        "Suburb_Median",
        "Price_Deviation",
    ]

    fact = df[fact_cols].copy()

    # Stable surrogate key for Power BI row-level operations
    fact.insert(0, "TransactionID", range(1, len(fact) + 1))

    # Ensure Date is plain date string (Power BI parses YYYY-MM-DD cleanly)
    fact["Date"] = pd.to_datetime(fact["Date"]).dt.strftime("%Y-%m-%d")

    fact = fact.rename(columns={
        "Rooms":    "Bedrooms",
        "Bathroom": "Bathrooms",
        "Car":      "Car_Spaces",
    })

    return fact


# ---------------------------------------------------------------------------
# Dimension tables
# ---------------------------------------------------------------------------

def build_dim_suburb(df: pd.DataFrame) -> pd.DataFrame:
    """
    One row per suburb.

    Includes geographic metadata, investment classification, and an
    Is_Value_Suburb flag for the three undervalued northern suburbs
    identified in the analysis (Reservoir, Glenroy, Coburg).
    """
    market_median = df["Price"].median()

    VALUE_SUBURBS = {"Reservoir", "Glenroy", "Coburg"}

    suburb_stats = (
        df.groupby("Suburb")
        .agg(
            Median_Price=("Price", "median"),
            Transaction_Count=("Price", "count"),
            Distance_to_CBD=("Distance", "first"),
            Postcode=("Postcode", "first"),
        )
        .round(0)
        .reset_index()
    )

    suburb_stats["Pct_vs_Market"] = (
        (suburb_stats["Median_Price"] - market_median) / market_median
    )

    suburb_stats["Investment_Rating"] = suburb_stats["Pct_vs_Market"].apply(
        _investment_rating
    )

    suburb_stats["Region"] = suburb_stats["Distance_to_CBD"].apply(_region)

    suburb_stats["Is_Value_Suburb"] = suburb_stats["Suburb"].isin(VALUE_SUBURBS)

    suburb_stats = suburb_stats.drop(columns=["Pct_vs_Market"])

    suburb_stats["Median_Price"] = suburb_stats["Median_Price"].astype(int)
    suburb_stats["Postcode"] = suburb_stats["Postcode"].astype(int)

    return suburb_stats.sort_values("Suburb").reset_index(drop=True)


def build_dim_date(df: pd.DataFrame) -> pd.DataFrame:
    """
    One row per unique sale date.

    Provides all time-intelligence fields Power BI needs to build
    quarter slicers, YoY comparisons, and trend lines without DAX
    calendar functions.
    """
    dates = pd.to_datetime(df["Date"]).drop_duplicates().sort_values()

    dim = pd.DataFrame({"Date": dates})
    dim["Date_Key"]      = dim["Date"].dt.strftime("%Y-%m-%d")
    dim["Year"]          = dim["Date"].dt.year
    dim["Month_Num"]     = dim["Date"].dt.month
    dim["Month_Name"]    = dim["Date"].dt.strftime("%B")
    dim["Month_Short"]   = dim["Date"].dt.strftime("%b")
    dim["Quarter_Num"]   = dim["Date"].dt.quarter
    dim["Quarter"]       = dim["Date"].dt.to_period("Q").astype(str)
    dim["Quarter_Label"] = "Q" + dim["Quarter_Num"].astype(str) + " " + dim["Year"].astype(str)
    dim["Year_Quarter"]  = dim["Year"].astype(str) + "-Q" + dim["Quarter_Num"].astype(str)

    # Flag the first and last quarters for Power BI time annotations
    min_q = dim["Quarter"].min()
    max_q = dim["Quarter"].max()
    dim["Is_First_Quarter"] = dim["Quarter"] == min_q
    dim["Is_Last_Quarter"]  = dim["Quarter"] == max_q

    dim["Date"] = dim["Date_Key"]
    dim = dim.drop(columns=["Date_Key"])

    return dim.reset_index(drop=True)


def build_dim_property_type() -> pd.DataFrame:
    """
    Static property type lookup table.
    """
    return pd.DataFrame([
        {"PropertyType": "House",     "Type_Code": "h", "Description": "Detached house",     "Is_House": True},
        {"PropertyType": "Unit",      "Type_Code": "u", "Description": "Apartment or unit",  "Is_House": False},
        {"PropertyType": "Townhouse", "Type_Code": "t", "Description": "Attached townhouse", "Is_House": False},
    ])


# ---------------------------------------------------------------------------
# Pre-aggregated summary tables
# ---------------------------------------------------------------------------

def build_agg_quarterly_trends(df: pd.DataFrame) -> pd.DataFrame:
    """
    Quarterly median price, mean price, volume, and QoQ growth.

    Power BI suggestion: Line chart -Quarter on X, Median_Price on Y.
    Add Transaction_Count as a secondary bar to show volume.
    """
    agg = (
        df.groupby("Quarter")
        .agg(
            Median_Price=("Price", "median"),
            Mean_Price=("Price", "mean"),
            Transaction_Count=("Price", "count"),
            Median_Price_per_sqm=("Price_per_sqm", "median"),
        )
        .round(0)
        .reset_index()
        .sort_values("Quarter")
    )

    agg["QoQ_Growth_Pct"] = (
        agg["Median_Price"].pct_change() * 100
    ).round(1)

    # Overall growth anchor for the resume KPI (Q2 2016 → Q3 2017)
    base = agg.loc[agg["Quarter"] == "2016Q2", "Median_Price"]
    peak = agg.loc[agg["Quarter"] == "2017Q3", "Median_Price"]
    if not base.empty and not peak.empty:
        agg["Growth_from_Base_Pct"] = (
            (agg["Median_Price"] - float(base.values[0])) / float(base.values[0]) * 100
        ).round(1)
    else:
        agg["Growth_from_Base_Pct"] = np.nan

    agg["Median_Price"]    = agg["Median_Price"].astype(int)
    agg["Mean_Price"]      = agg["Mean_Price"].astype(int)

    return agg


def build_agg_suburb_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    Suburb-level KPIs sorted by median price descending.

    Power BI suggestion: Bar chart -Suburb on Y, Median_Price on X.
    Use Investment_Rating as a slicer.
    """
    market_median = df["Price"].median()

    agg = (
        df.groupby("Suburb")
        .agg(
            Median_Price=("Price", "median"),
            Mean_Price=("Price", "mean"),
            Transaction_Count=("Price", "count"),
            Distance_to_CBD=("Distance", "first"),
            Median_Price_per_sqm=("Price_per_sqm", "median"),
            Median_Bedrooms=("Rooms", "median"),
        )
        .round(0)
        .reset_index()
        .sort_values("Median_Price", ascending=False)
    )

    agg["Vs_Market_Median_Pct"] = (
        (agg["Median_Price"] - market_median) / market_median * 100
    ).round(1)

    agg["Investment_Rating"] = (agg["Vs_Market_Median_Pct"] / 100).apply(
        _investment_rating
    )

    agg["Median_Price"] = agg["Median_Price"].astype(int)
    agg["Mean_Price"]   = agg["Mean_Price"].astype(int)

    return agg.reset_index(drop=True)


def build_agg_property_type_premium(df: pd.DataFrame) -> pd.DataFrame:
    """
    House-vs-unit median price premium by suburb.

    Power BI suggestion: Clustered bar -House, Unit, Townhouse medians
    side by side per suburb. Add House_vs_Unit_Premium_Pct as a tooltip.
    """
    pivot = (
        df.groupby(["Suburb", "PropertyType"])["Price"]
        .median()
        .unstack()
        .reset_index()
    )

    for col in ["House", "Unit", "Townhouse"]:
        if col not in pivot.columns:
            pivot[col] = np.nan

    pivot = pivot[["Suburb", "House", "Unit", "Townhouse"]]

    pivot["House_vs_Unit_Premium_Pct"] = (
        (pivot["House"] - pivot["Unit"]) / pivot["Unit"] * 100
    ).round(1)

    pivot = pivot.dropna(subset=["House", "Unit"]).sort_values(
        "House_vs_Unit_Premium_Pct", ascending=False
    )

    for col in ["House", "Unit", "Townhouse"]:
        pivot[col] = pivot[col].where(pivot[col].notna(), other=np.nan)
        pivot[col] = pivot[col].apply(lambda x: int(x) if pd.notna(x) else "")

    return pivot.reset_index(drop=True)


def build_agg_value_gaps(df: pd.DataFrame) -> pd.DataFrame:
    """
    Undervalued suburb comparisons vs adjacent premium suburbs.

    The three pairs mirror the resume finding: Reservoir/Glenroy/Coburg
    trading at 45–46% discounts to neighbouring premium suburbs.

    Power BI suggestion: Grouped bar -Value_Price vs Premium_Price,
    with Discount_Pct shown as a data label.
    """
    suburb_medians = df.groupby("Suburb")["Price"].median()

    comparisons = [
        ("Reservoir",  "Northcote"),
        ("Glenroy",    "Moonee Ponds"),
        ("Coburg",     "Brunswick"),
    ]

    rows = []
    for value_sub, premium_sub in comparisons:
        v_price = suburb_medians.get(value_sub, np.nan)
        p_price = suburb_medians.get(premium_sub, np.nan)

        if pd.notna(v_price) and pd.notna(p_price) and p_price > 0:
            discount = round((p_price - v_price) / p_price * 100, 1)
        else:
            discount = np.nan

        rows.append({
            "Value_Suburb":   value_sub,
            "Premium_Suburb": premium_sub,
            "Value_Price":    int(v_price) if pd.notna(v_price) else "",
            "Premium_Price":  int(p_price) if pd.notna(p_price) else "",
            "Discount_Pct":   discount,
        })

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    project_root = Path(__file__).parent.parent
    data_dir     = project_root / "data"
    output_dir   = data_dir / "powerbi"
    output_dir.mkdir(parents=True, exist_ok=True)

    # Load feature-engineered analysis data
    print("Loading analysis data...")
    df = pd.read_csv(data_dir / "processed" / "melb_data_analysis.csv")
    df["Date"] = pd.to_datetime(df["Date"])
    print(f"  Loaded {len(df):,} records across {df['Suburb'].nunique()} suburbs")

    print("\nBuilding star-schema tables...")

    # --- Fact table ---
    fact = build_fact_transactions(df)
    fact.to_csv(output_dir / "fact_transactions.csv", index=False)
    print(f"  fact_transactions.csv        {len(fact):>6,} rows  {len(fact.columns)} cols")

    # --- Dimension tables ---
    dim_suburb = build_dim_suburb(df)
    dim_suburb.to_csv(output_dir / "dim_suburb.csv", index=False)
    print(f"  dim_suburb.csv               {len(dim_suburb):>6,} rows  {len(dim_suburb.columns)} cols")

    dim_date = build_dim_date(df)
    dim_date.to_csv(output_dir / "dim_date.csv", index=False)
    print(f"  dim_date.csv                 {len(dim_date):>6,} rows  {len(dim_date.columns)} cols")

    dim_property_type = build_dim_property_type()
    dim_property_type.to_csv(output_dir / "dim_property_type.csv", index=False)
    print(f"  dim_property_type.csv        {len(dim_property_type):>6,} rows  {len(dim_property_type.columns)} cols")

    # --- Pre-aggregated tables ---
    agg_quarterly = build_agg_quarterly_trends(df)
    agg_quarterly.to_csv(output_dir / "agg_quarterly_trends.csv", index=False)
    print(f"  agg_quarterly_trends.csv     {len(agg_quarterly):>6,} rows  {len(agg_quarterly.columns)} cols")

    agg_suburb = build_agg_suburb_summary(df)
    agg_suburb.to_csv(output_dir / "agg_suburb_summary.csv", index=False)
    print(f"  agg_suburb_summary.csv       {len(agg_suburb):>6,} rows  {len(agg_suburb.columns)} cols")

    agg_premium = build_agg_property_type_premium(df)
    agg_premium.to_csv(output_dir / "agg_property_type_premium.csv", index=False)
    print(f"  agg_property_type_premium.csv{len(agg_premium):>6,} rows  {len(agg_premium.columns)} cols")

    agg_gaps = build_agg_value_gaps(df)
    agg_gaps.to_csv(output_dir / "agg_value_gaps.csv", index=False)
    print(f"  agg_value_gaps.csv           {len(agg_gaps):>6,} rows  {len(agg_gaps.columns)} cols")

    # --- Key stats summary ---
    print("\n" + "=" * 60)
    print("EXPORT SUMMARY")
    print("=" * 60)

    market_median = df["Price"].median()
    q2_2016 = df[df["Quarter"] == "2016Q2"]["Price"].median()
    q3_2017 = df[df["Quarter"] == "2017Q3"]["Price"].median()
    growth  = (q3_2017 - q2_2016) / q2_2016 * 100

    print(f"  Market median price:   ${market_median:,.0f}")
    print(f"  Market growth:         +{growth:.0f}%  (Q2 2016 → Q3 2017)")
    print(f"  Value suburbs (3):     Reservoir, Glenroy, Coburg")

    gaps = build_agg_value_gaps(df)
    for _, row in gaps.iterrows():
        print(f"    {row['Value_Suburb']:<12} {row['Discount_Pct']:.0f}% below {row['Premium_Suburb']}")

    total_files = 8
    print(f"\n  {total_files} files written to: {output_dir}")
    print("\nNext step: open Power BI Desktop and import all CSVs from data/powerbi/")
    print("           then copy DAX measures from data/powerbi/powerbi_measures.dax")

    return output_dir


if __name__ == "__main__":
    main()
