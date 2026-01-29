"""
Melbourne Housing Data Cleaning Script

Cleans raw housing data from Kaggle Melbourne Housing dataset.
Removes invalid records, standardizes suburb names, and exports cleaned data.
"""

import pandas as pd
import numpy as np
from pathlib import Path


def load_raw_data(filepath: str) -> pd.DataFrame:
    """Load raw Melbourne housing data."""
    return pd.read_csv(filepath)


def audit_data(df: pd.DataFrame) -> dict:
    """Generate audit statistics for the dataset."""
    return {
        "total_rows": len(df),
        "total_columns": len(df.columns),
        "completeness_pct": df.notna().mean().mean() * 100,
        "missing_by_column": df.isnull().sum().to_dict()
    }


def clean_data(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Clean the housing dataset.

    Cleaning steps:
    1. Standardize suburb names (strip whitespace, title case)
    2. Remove records with invalid Landsize (<=0 or >50,000 sqm)
    3. Fill missing Car values with 0

    Returns:
        Tuple of (cleaned DataFrame, cleaning statistics dict)
    """
    initial_rows = len(df)
    initial_completeness = df.notna().mean().mean() * 100

    df_clean = df.copy()

    # Standardize suburb names
    suburb_fixes = (df_clean["Suburb"].str.strip().str.title() != df_clean["Suburb"]).sum()
    df_clean["Suburb"] = df_clean["Suburb"].str.strip().str.title()

    # Remove invalid Landsize records
    invalid_landsize = df_clean["Landsize"] <= 0
    extreme_landsize = df_clean["Landsize"] > 50000
    removed_landsize = invalid_landsize.sum() + extreme_landsize.sum()
    df_clean = df_clean[~(invalid_landsize | extreme_landsize)]

    # Fill missing Car values
    df_clean["Car"] = df_clean["Car"].fillna(0)

    final_completeness = df_clean.notna().mean().mean() * 100

    stats = {
        "initial_rows": initial_rows,
        "final_rows": len(df_clean),
        "removed_rows": initial_rows - len(df_clean),
        "suburb_name_fixes": suburb_fixes,
        "removed_invalid_landsize": removed_landsize,
        "initial_completeness_pct": round(initial_completeness, 1),
        "final_completeness_pct": round(final_completeness, 1)
    }

    return df_clean, stats


def main():
    """Main execution function."""
    project_root = Path(__file__).parent.parent
    data_dir = project_root / "data"

    # Load data
    print("Loading raw data...")
    df = load_raw_data(data_dir / "raw" / "melb_data.csv")

    # Audit before cleaning
    print("\nInitial audit:")
    audit = audit_data(df)
    print(f"  Rows: {audit['total_rows']:,}")
    print(f"  Completeness: {audit['completeness_pct']:.1f}%")

    # Clean data
    print("\nCleaning data...")
    df_clean, stats = clean_data(df)

    # Print cleaning summary
    print("\nCleaning Summary:")
    print(f"  Initial rows: {stats['initial_rows']:,}")
    print(f"  Final rows: {stats['final_rows']:,}")
    print(f"  Removed: {stats['removed_rows']:,}")
    print(f"  Completeness: {stats['initial_completeness_pct']}% -> {stats['final_completeness_pct']}%")

    # Save cleaned data
    output_path = data_dir / "processed" / "melb_data_cleaned.csv"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df_clean.to_csv(output_path, index=False)
    print(f"\nCleaned data saved to: {output_path}")


if __name__ == "__main__":
    main()
