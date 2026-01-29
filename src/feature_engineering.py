"""
Melbourne Housing Feature Engineering Script

Creates analysis features and filters to top 20 suburbs by transaction volume.
"""

import pandas as pd
import numpy as np
from pathlib import Path


def add_features(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add analysis features to the dataset.

    Features added:
    - Price_per_sqm: Price divided by Landsize
    - Date parsed to datetime
    - Year, Month, Quarter extracted
    - PropertyType: Human-readable property type labels
    """
    df = df.copy()

    # Price per sqm
    df["Price_per_sqm"] = df["Price"] / df["Landsize"]

    # Parse dates
    df["Date"] = pd.to_datetime(df["Date"], format="%d/%m/%Y")
    df["Year"] = df["Date"].dt.year
    df["Month"] = df["Date"].dt.month
    df["Quarter"] = df["Date"].dt.to_period("Q").astype(str)

    # Property type labels
    type_map = {"h": "House", "u": "Unit", "t": "Townhouse"}
    df["PropertyType"] = df["Type"].map(type_map)

    return df


def get_top_suburbs(df: pd.DataFrame, n: int = 20) -> list:
    """Get top N suburbs by transaction volume."""
    return df["Suburb"].value_counts().head(n).index.tolist()


def add_outlier_detection(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add outlier detection columns based on suburb median.

    Adds:
    - Suburb_Median: Median price for the suburb
    - Price_Deviation: Percentage deviation from suburb median
    """
    df = df.copy()
    suburb_medians = df.groupby("Suburb")["Price"].median()
    df["Suburb_Median"] = df["Suburb"].map(suburb_medians)
    df["Price_Deviation"] = (df["Price"] - df["Suburb_Median"]) / df["Suburb_Median"]
    return df


def main():
    """Main execution function."""
    project_root = Path(__file__).parent.parent
    data_dir = project_root / "data"

    # Load cleaned data
    print("Loading cleaned data...")
    df = pd.read_csv(data_dir / "processed" / "melb_data_cleaned.csv")
    print(f"  Loaded {len(df):,} records")

    # Add features
    print("\nEngineering features...")
    df = add_features(df)
    print(f"  Added: Price_per_sqm, Quarter, PropertyType")
    print(f"  Date range: {df['Date'].min().date()} to {df['Date'].max().date()}")

    # Get top 20 suburbs
    top_suburbs = get_top_suburbs(df, n=20)
    print(f"\nTop 20 suburbs by volume:")
    for i, suburb in enumerate(top_suburbs[:5], 1):
        count = (df["Suburb"] == suburb).sum()
        print(f"  {i}. {suburb}: {count} transactions")
    print(f"  ... and 15 more")

    # Filter to analysis subset
    df_analysis = df[df["Suburb"].isin(top_suburbs)].copy()
    df_analysis = add_outlier_detection(df_analysis)

    print(f"\nAnalysis dataset: {len(df_analysis):,} records ({len(df_analysis)/len(df)*100:.1f}% of data)")

    # Count outliers
    underpriced = (df_analysis["Price_Deviation"] < -0.3).sum()
    overpriced = (df_analysis["Price_Deviation"] > 0.5).sum()
    print(f"  Underpriced properties (<30% below median): {underpriced}")
    print(f"  Overpriced properties (>50% above median): {overpriced}")

    # Save analysis dataset
    output_path = data_dir / "processed" / "melb_data_analysis.csv"
    df_analysis.to_csv(output_path, index=False)
    print(f"\nAnalysis data saved to: {output_path}")


if __name__ == "__main__":
    main()
