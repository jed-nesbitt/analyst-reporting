from __future__ import annotations

import pandas as pd


def _add_month(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "date" in df.columns:
        df["month"] = df["date"].dt.to_period("M").dt.to_timestamp()
    else:
        df["month"] = pd.NaT
    return df


def build_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Returns 4 tables:
      - Summary
      - Trends (by month)
      - Variance (MoM)
      - Drilldowns (region/product pivots combined on one sheet later)
    """
    df = _add_month(df)

    # Ensure columns exist even if missing in source
    for c in ["cost", "units", "region", "product"]:
        if c not in df.columns:
            df[c] = pd.NA

    df["cost"] = pd.to_numeric(df["cost"], errors="coerce")
    df["units"] = pd.to_numeric(df["units"], errors="coerce")
    df["revenue"] = pd.to_numeric(df["revenue"], errors="coerce")

    df["gross_profit"] = df["revenue"] - df["cost"]
    df["margin"] = df["gross_profit"] / df["revenue"]

    # --- Summary (overall) ---
    summary = pd.DataFrame(
        {
            "metric": ["revenue", "cost", "gross_profit", "margin", "units", "rows_loaded"],
            "value": [
                df["revenue"].sum(skipna=True),
                df["cost"].sum(skipna=True),
                df["gross_profit"].sum(skipna=True),
                (df["gross_profit"].sum(skipna=True) / df["revenue"].sum(skipna=True)) if df["revenue"].sum(skipna=True) else pd.NA,
                df["units"].sum(skipna=True),
                len(df),
            ],
        }
    )

    # --- Trends (monthly) ---
    trends = (
        df.groupby("month", dropna=False, as_index=False)
        .agg(
            revenue=("revenue", "sum"),
            cost=("cost", "sum"),
            gross_profit=("gross_profit", "sum"),
            units=("units", "sum"),
        )
        .sort_values("month")
    )
    trends["margin"] = trends["gross_profit"] / trends["revenue"]

    # --- Variance (MoM) ---
    variance = trends.copy()
    for col in ["revenue", "cost", "gross_profit", "units", "margin"]:
        variance[f"{col}_mom_abs"] = variance[col].diff()
        variance[f"{col}_mom_pct"] = variance[col].pct_change()

    # --- Drilldowns (for later writing onto one sheet) ---
    by_region = (
        df.groupby(["month", "region"], dropna=False, as_index=False)
        .agg(revenue=("revenue", "sum"), gross_profit=("gross_profit", "sum"))
        .sort_values(["month", "region"])
    )

    by_product = (
        df.groupby(["month", "product"], dropna=False, as_index=False)
        .agg(revenue=("revenue", "sum"), gross_profit=("gross_profit", "sum"))
        .sort_values(["month", "product"])
    )

    return {
        "Summary": summary,
        "Trends": trends,
        "Variance": variance,
        "Drilldown_Region": by_region,
        "Drilldown_Product": by_product,
    }
