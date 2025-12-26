from __future__ import annotations

from pathlib import Path
from datetime import datetime

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image,
    PageBreak,
    KeepTogether,
)
from reportlab.lib.styles import getSampleStyleSheet


def _fmt_num(x) -> str:
    if pd.isna(x):
        return ""
    if isinstance(x, (int, float)):
        return f"{x:,.2f}"
    return str(x)


def _fmt_currency(x, currency_code: str) -> str:
    if pd.isna(x):
        return "N/A"
    try:
        v = float(x)
    except Exception:
        return "N/A"
    code = (currency_code or "AUD").upper().strip()
    return f"{code} {v:,.2f}"


def _fmt_pct(x) -> str:
    if pd.isna(x):
        return "N/A"
    try:
        v = float(x)
    except Exception:
        return "N/A"
    return f"{v*100:.2f}%"


def _fmt_bps_from_ratio(x) -> str:
    if pd.isna(x):
        return "N/A"
    try:
        v = float(x)
    except Exception:
        return "N/A"
    bps = v * 10000.0
    sign = "+" if bps > 0 else ""
    return f"{sign}{bps:,.0f} bps"


def _kpi_from_summary(summary: pd.DataFrame) -> dict[str, float | str]:
    return dict(zip(summary["metric"], summary["value"]))


def _safe_latest_month(trends: pd.DataFrame) -> pd.Timestamp | None:
    if trends is None or trends.empty or "month" not in trends.columns:
        return None
    s = pd.to_datetime(trends["month"], errors="coerce").dropna()
    if s.empty:
        return None
    return s.max()


def _last_value_by_month(df: pd.DataFrame, col: str) -> float | pd.NA:
    if df is None or df.empty or col not in df.columns:
        return pd.NA
    d = df.copy()
    if "month" in d.columns:
        d["month"] = pd.to_datetime(d["month"], errors="coerce")
        d = d.sort_values("month")
    s = pd.to_numeric(d[col], errors="coerce").dropna()
    if s.empty:
        return pd.NA
    return float(s.iloc[-1])


def _top_dim_latest_month(
    drill: pd.DataFrame, dim_col: str, value_col: str, latest: pd.Timestamp | None
) -> tuple[str, float | pd.NA]:
    if latest is None or drill is None or drill.empty:
        return ("", pd.NA)
    need = {"month", dim_col, value_col}
    if not need.issubset(drill.columns):
        return ("", pd.NA)

    d = drill.copy()
    d["month"] = pd.to_datetime(d["month"], errors="coerce")
    d = d[d["month"] == latest].copy()
    if d.empty:
        return ("", pd.NA)

    d[value_col] = pd.to_numeric(d[value_col], errors="coerce")
    d = d.dropna(subset=[value_col])
    if d.empty:
        return ("", pd.NA)

    r = d.loc[d[value_col].idxmax()]
    return (str(r.get(dim_col, "Unknown")), float(r[value_col]))


def _build_exec_insights(
    *,
    summary: pd.DataFrame,
    trends: pd.DataFrame,
    variance: pd.DataFrame,
    drill_region: pd.DataFrame,
    drill_product: pd.DataFrame,
    currency_code: str,
) -> list[str]:
    k = _kpi_from_summary(summary)
    latest = _safe_latest_month(trends)
    latest_label = latest.strftime("%Y-%m") if latest is not None else "N/A"

    rev_mom_pct = _last_value_by_month(variance, "revenue_mom_pct")
    gp_mom_pct = _last_value_by_month(variance, "gross_profit_mom_pct")
    margin_mom_abs = _last_value_by_month(variance, "margin_mom_abs")

    top_region, top_region_rev = _top_dim_latest_month(drill_region, "region", "revenue", latest)
    top_product, top_product_rev = _top_dim_latest_month(drill_product, "product", "revenue", latest)

    lines: list[str] = []
    lines.append(f"Latest month in dataset: {latest_label}")

    lines.append(f"Revenue MoM: {_fmt_pct(rev_mom_pct)}" if pd.notna(rev_mom_pct) else "Revenue MoM: N/A")
    lines.append(f"Gross Profit MoM: {_fmt_pct(gp_mom_pct)}" if pd.notna(gp_mom_pct) else "Gross Profit MoM: N/A")
    lines.append(f"Margin change (MoM): {_fmt_bps_from_ratio(margin_mom_abs)}" if pd.notna(margin_mom_abs) else "Margin change (MoM): N/A")

    if top_region:
        lines.append(f"Top region by revenue ({latest_label}): {top_region} ({_fmt_currency(top_region_rev, currency_code)})")
    else:
        lines.append(f"Top region by revenue ({latest_label}): N/A")

    if top_product:
        lines.append(f"Top product by revenue ({latest_label}): {top_product} ({_fmt_currency(top_product_rev, currency_code)})")
    else:
        lines.append(f"Top product by revenue ({latest_label}): N/A")

    revenue = k.get("revenue", pd.NA)
    gp = k.get("gross_profit", pd.NA)
    margin = k.get("margin", pd.NA)

    lines.append(f"Total revenue: {_fmt_currency(revenue, currency_code)}")
    lines.append(f"Total gross profit: {_fmt_currency(gp, currency_code)}")
    lines.append(f"Overall margin: {_fmt_pct(margin)}")

    return lines


def write_pdf_report(
    out_path: Path,
    summary: pd.DataFrame,
    trends: pd.DataFrame,
    variance: pd.DataFrame,
    drill_region: pd.DataFrame,
    drill_product: pd.DataFrame,
    chart_paths: list[Path],
    source_label: str,
    currency_code: str = "AUD",
    report_title: str = "Analyst Reporting Pack",
    report_subtitle: str = "",
    notes: list[str] | None = None,
    warnings: list[str] | None = None,
) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)

    styles = getSampleStyleSheet()
    story = []

    # ---- Title page ----
    story.append(Paragraph(report_title or "Analyst Reporting Pack", styles["Title"]))
    if report_subtitle:
        story.append(Spacer(1, 0.15 * cm))
        story.append(Paragraph(report_subtitle, styles["Heading2"]))

    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph(f"Source: <b>{source_label}</b>", styles["Normal"]))
    story.append(Paragraph(f"Currency: <b>{(currency_code or 'AUD').upper().strip()}</b>", styles["Normal"]))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles["Normal"]))
    story.append(Spacer(1, 0.35 * cm))

    if warnings:
        story.append(Paragraph("Warnings", styles["Heading2"]))
        for w in warnings[:10]:
            story.append(Paragraph(f"• {w}", styles["Normal"]))
        if len(warnings) > 10:
            story.append(Paragraph(f"(+ {len(warnings) - 10} more)", styles["Normal"]))
        story.append(Spacer(1, 0.25 * cm))

    # ---- KPI block ----
    story.append(Paragraph("Key KPIs", styles["Heading2"]))
    k = _kpi_from_summary(summary)

    revenue = k.get("revenue", pd.NA)
    gp = k.get("gross_profit", pd.NA)
    margin = k.get("margin", pd.NA)
    units = k.get("units", pd.NA)
    rows_loaded = k.get("rows_loaded", pd.NA)

    kpi_rows = [
        ["Metric", "Value"],
        ["Revenue", _fmt_currency(revenue, currency_code)],
        ["Gross Profit", _fmt_currency(gp, currency_code)],
        ["Margin", _fmt_pct(margin)],
        ["Units", _fmt_num(units)],
        ["Rows Loaded", _fmt_num(rows_loaded)],
    ]

    kpi_tbl = Table(kpi_rows, colWidths=[6 * cm, 10 * cm])
    kpi_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("PADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(kpi_tbl)
    story.append(Spacer(1, 0.25 * cm))

    # ---- Executive Summary (Insights) ----
    story.append(Paragraph("Executive Summary", styles["Heading2"]))
    insights = _build_exec_insights(
        summary=summary,
        trends=trends,
        variance=variance,
        drill_region=drill_region,
        drill_product=drill_product,
        currency_code=currency_code,
    )
    for line in insights[:10]:
        story.append(Paragraph(f"• {line}", styles["Normal"]))

    # ---- Commentary from config.yaml ----
    clean_notes = [str(x).strip() for x in (notes or []) if str(x).strip()]
    if clean_notes:
        story.append(Spacer(1, 0.2 * cm))
        story.append(Paragraph("Commentary", styles["Heading2"]))
        for n in clean_notes[:8]:
            story.append(Paragraph(f"• {n}", styles["Normal"]))
        if len(clean_notes) > 8:
            story.append(Paragraph(f"(+ {len(clean_notes) - 8} more)", styles["Normal"]))

    # Start charts on a fresh page so 2-per-page fits reliably
    story.append(PageBreak())

    # ---- Charts (2 per page) ----
    story.append(Paragraph("Charts", styles["Heading2"]))
    story.append(Spacer(1, 0.2 * cm))

    page_w, page_h = A4
    left_right = 1.7 * cm
    top_bottom = 1.7 * cm
    max_w = page_w - 2 * left_right

    usable_h = page_h - 2 * top_bottom
    max_h_each = (usable_h - 4.5 * cm) / 2

    chart_paths = [Path(p) for p in chart_paths if Path(p).exists()]

    for idx in range(0, len(chart_paths), 2):
        pair: list = []

        p1 = chart_paths[idx]
        pair.append(Paragraph(p1.name, styles["Normal"]))
        pair.append(Spacer(1, 0.12 * cm))
        img1 = Image(str(p1))
        img1.hAlign = "CENTER"
        img1._restrictSize(max_w, max_h_each)
        pair.append(img1)

        pair.append(Spacer(1, 0.35 * cm))

        if idx + 1 < len(chart_paths):
            p2 = chart_paths[idx + 1]
            pair.append(Paragraph(p2.name, styles["Normal"]))
            pair.append(Spacer(1, 0.12 * cm))
            img2 = Image(str(p2))
            img2.hAlign = "CENTER"
            img2._restrictSize(max_w, max_h_each)
            pair.append(img2)

        story.append(KeepTogether(pair))

        if idx + 2 < len(chart_paths):
            story.append(PageBreak())

    # ---- Variance snapshot (optional) ----
    if variance is not None and not variance.empty:
        story.append(PageBreak())
        story.append(Paragraph("Variance Snapshot (Top 12 Rows)", styles["Heading2"]))
        v = variance.head(12).copy()

        keep = [c for c in ["month", "revenue", "revenue_mom_abs", "revenue_mom_pct", "margin", "margin_mom_abs"] if c in v.columns]
        if keep:
            v = v[keep]

        if "month" in v.columns:
            v["month"] = pd.to_datetime(v["month"], errors="coerce").dt.strftime("%Y-%m")

        v_fmt = v.apply(lambda col: col.map(_fmt_num))
        table_data = [list(v_fmt.columns)] + v_fmt.values.tolist()

        tbl = Table(table_data, repeatRows=1)
        tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("PADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(tbl)

    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=A4,
        leftMargin=1.7 * cm,
        rightMargin=1.7 * cm,
        topMargin=1.7 * cm,
        bottomMargin=1.7 * cm,
        title=report_title or "Analyst Reporting Pack",
    )
    doc.build(story)
