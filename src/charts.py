from __future__ import annotations

from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def _safe_latest_month(trends: pd.DataFrame) -> pd.Timestamp | None:
    if "month" not in trends.columns:
        return None
    s = pd.to_datetime(trends["month"], errors="coerce")
    s = s.dropna()
    if s.empty:
        return None
    return s.max()


def save_line_chart(x, y, title: str, ylabel: str, out_path: Path) -> None:
    fig = plt.figure()
    plt.plot(x, y)
    plt.title(title)
    plt.xlabel("Month")
    plt.ylabel(ylabel)
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    fig.savefig(out_path, dpi=200)
    plt.close(fig)


def save_bar_chart(labels, values, title: str, ylabel: str, out_path: Path, top_n: int = 12) -> None:
    # keep top N by value
    s = pd.Series(values, index=pd.Index(labels, dtype=str)).dropna()
    if s.empty:
        return
    s = s.sort_values(ascending=False).head(top_n)

    fig = plt.figure()
    plt.bar(s.index.astype(str), s.values)
    plt.title(title)
    plt.xlabel("")
    plt.ylabel(ylabel)
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    fig.savefig(out_path, dpi=200)
    plt.close(fig)


def generate_charts(
    trends: pd.DataFrame,
    drill_region: pd.DataFrame,
    drill_product: pd.DataFrame,
    out_dir: Path,
) -> list[Path]:
    """
    Creates PNG charts and returns the list of generated file paths.
    Uses:
      - trends (monthly KPI table)
      - drilldowns (month x region/product)
    """
    ensure_dir(out_dir)
    created: list[Path] = []

    # ----- Trend charts -----
    t = trends.copy()
    if "month" in t.columns:
        t["month"] = pd.to_datetime(t["month"], errors="coerce")

    if "month" in t.columns and t["month"].notna().any():
        t = t.sort_values("month")
        x = t["month"]

        if "revenue" in t.columns:
            p = out_dir / "trend_revenue.png"
            save_line_chart(x, t["revenue"], "Revenue Trend", "Revenue", p)
            created.append(p)

        if "gross_profit" in t.columns:
            p = out_dir / "trend_gross_profit.png"
            save_line_chart(x, t["gross_profit"], "Gross Profit Trend", "Gross Profit", p)
            created.append(p)

        if "margin" in t.columns:
            p = out_dir / "trend_margin.png"
            save_line_chart(x, t["margin"], "Margin Trend", "Margin", p)
            created.append(p)

    # ----- Latest month bar charts -----
    latest = _safe_latest_month(t)
    if latest is not None:
        # Revenue by Region (latest month)
        if not drill_region.empty and {"month", "region", "revenue"}.issubset(drill_region.columns):
            dr = drill_region.copy()
            dr["month"] = pd.to_datetime(dr["month"], errors="coerce")
            dr_latest = dr[dr["month"] == latest].copy()
            if not dr_latest.empty:
                p = out_dir / "latest_month_revenue_by_region.png"
                save_bar_chart(
                    labels=dr_latest["region"].fillna("Unknown"),
                    values=dr_latest["revenue"],
                    title=f"Revenue by Region ({latest:%Y-%m})",
                    ylabel="Revenue",
                    out_path=p,
                )
                created.append(p)

        # Revenue by Product (latest month)
        if not drill_product.empty and {"month", "product", "revenue"}.issubset(drill_product.columns):
            dp = drill_product.copy()
            dp["month"] = pd.to_datetime(dp["month"], errors="coerce")
            dp_latest = dp[dp["month"] == latest].copy()
            if not dp_latest.empty:
                p = out_dir / "latest_month_revenue_by_product.png"
                save_bar_chart(
                    labels=dp_latest["product"].fillna("Unknown"),
                    values=dp_latest["revenue"],
                    title=f"Revenue by Product ({latest:%Y-%m})",
                    ylabel="Revenue",
                    out_path=p,
                )
                created.append(p)

    return created
