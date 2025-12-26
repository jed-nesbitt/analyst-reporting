from __future__ import annotations

import argparse
from pathlib import Path
import pandas as pd


def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Split Superstore into monthly 'messy' dumps for the reporting suite.")
    ap.add_argument("--input", type=str, required=True, help="Path to superstore.csv")
    ap.add_argument("--out", type=str, default="data/in", help="Output folder for monthly dumps")
    return ap.parse_args()


def main() -> int:
    args = parse_args()
    in_path = Path(args.input)
    out_dir = Path(args.out)

    if not in_path.exists():
        raise FileNotFoundError(f"Missing input file: {in_path}")

    ensure_dir(out_dir)

    df = pd.read_csv(in_path)

    # Superstore columns commonly include:
    # Order Date, State, Region, Product Name, Sales, Profit, Quantity
    df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")
    df = df[df["Order Date"].notna()].copy()

    # Build columns that match our project's aliases:
    # Transaction Date -> date
    # Sales($) -> revenue
    # COGS -> cost (computed from Sales - Profit)
    # Qty -> units
    # State -> region
    # SKU -> product
    revenue = pd.to_numeric(df["Sales"], errors="coerce")
    profit = pd.to_numeric(df["Profit"], errors="coerce")
    qty = pd.to_numeric(df.get("Quantity", pd.NA), errors="coerce")

    out = pd.DataFrame(
        {
            "Transaction Date": df["Order Date"],
            " State ": df.get("State", df.get("Region", pd.NA)),  # intentionally messy spacing
            "SKU": df.get("Product Name", df.get("Sub-Category", pd.NA)),
            "Sales($)": revenue,
            "COGS": revenue - profit,  # cost derived from profit
            "Qty": qty,
        }
    )

    out["month"] = out["Transaction Date"].dt.to_period("M").astype(str)

    # Write each month as its own dump (mix CSV + XLSX to test both ingest paths)
    for m, g in out.groupby("month"):
        g2 = g.drop(columns=["month"]).copy()
        # alternate file type by month number for realism
        month_num = int(m.split("-")[1])
        if month_num % 2 == 0:
            path = out_dir / f"{m}_dump.xlsx"
            g2.to_excel(path, index=False)
        else:
            path = out_dir / f"{m}_dump.csv"
            g2.to_csv(path, index=False)

    print(f"✅ Wrote monthly dumps to: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
