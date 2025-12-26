from __future__ import annotations

import argparse
from pathlib import Path
import sys
import random
import pandas as pd

from src.ingest import read_all
from src.clean import clean
from src.kpis import build_tables
from src.export_excel import write_excel_pack
from src.charts import generate_charts
from src.pdf_report import write_pdf_report
from src.config import resolve_config
from src.quality import build_quality_report, write_quality_excel
from src.runlog import write_run_log


def make_demo_inputs(input_dir: Path) -> None:
    random.seed(42)
    input_dir.mkdir(parents=True, exist_ok=True)

    regions = ["NSW", "VIC", "QLD", "WA"]
    products = ["Widget A", "Widget B", "Widget C"]

    def mk(month_start: str, n: int) -> pd.DataFrame:
        rows = []
        for _ in range(n):
            rows.append(
                {
                    "Transaction Date": month_start,
                    " State ": random.choice(regions),
                    "Sales($)": round(random.uniform(100, 1200), 2),
                    "COGS": round(random.uniform(40, 700), 2),
                    "Qty": random.randint(1, 20),
                    "SKU": random.choice(products),
                }
            )
        df = pd.DataFrame(rows)
        df.loc[3, "Sales($)"] = ""
        df.loc[7, "Transaction Date"] = "not a date"
        return df

    mk("2025-01-01", 60).to_csv(input_dir / "jan_dump.csv", index=False)
    mk("2025-02-01", 55).to_excel(input_dir / "feb_dump.xlsx", index=False)


def parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Analyst Reporting Automation Suite")

    # Default config.yaml so `python main.py` just works
    p.add_argument("--config", type=str, default="config.yaml", help="Path to config.yaml (default: config.yaml)")
    p.add_argument("--input", type=str, default=None, help="Override input_dir from config")
    p.add_argument("--out", type=str, default=None, help="Override out_dir from config")
    p.add_argument("--currency", type=str, default=None, help="Override currency_code from config (e.g. AUD/USD)")

    p.add_argument("--demo", action="store_true", help="Generate demo inputs into input folder then run")

    # Optional override for PDF only (config controls defaults)
    g = p.add_mutually_exclusive_group()
    g.add_argument("--pdf", action="store_true", help="Force PDF on (override config)")
    g.add_argument("--no-pdf", action="store_true", help="Force PDF off (override config)")

    return p.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)

    try:
        cfg = resolve_config(
            config_path=args.config,
            cli_input=args.input,
            cli_out=args.out,
            cli_currency=args.currency,
        )
    except Exception as e:
        print(f"ERROR loading config: {e}", file=sys.stderr)
        return 2

    # CLI override for PDF
    make_pdf = cfg.make_pdf
    if args.pdf:
        make_pdf = True
    if args.no_pdf:
        make_pdf = False

    input_dir = cfg.input_dir
    out_dir = cfg.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.demo:
        make_demo_inputs(input_dir)

    df_raw = read_all(input_dir)
    if df_raw.empty:
        print(f"ERROR: No input files found in {input_dir}", file=sys.stderr)
        return 2

    # Clean (apply aliases from config)
    result = clean(df_raw, aliases=cfg.aliases)

    # KPI tables
    tables = build_tables(result.df)

    # Optional artifacts
    if cfg.write_cleaned_csv:
        cleaned_csv = out_dir / "cleaned_data.csv"
        result.df.to_csv(cleaned_csv, index=False)
        print(f"🧼 Cleaned data saved: {cleaned_csv}")

    if cfg.write_quality_report:
        dq_path = out_dir / "data_quality.xlsx"
        qr = build_quality_report(result.df)
        try:
            write_quality_excel(dq_path, qr)
        except PermissionError:
            print(f"ERROR: Can't write {dq_path}. Close it if open in Excel, then re-run.", file=sys.stderr)
        else:
            print(f"✅ Data quality report created: {dq_path}")

    # Excel pack
    if cfg.write_excel_pack:
        out_path = out_dir / "report_pack.xlsx"
        write_excel_pack(
            out_path=out_path,
            summary=tables["Summary"],
            trends=tables["Trends"],
            variance=tables["Variance"],
            drill_region=tables["Drilldown_Region"],
            drill_product=tables["Drilldown_Product"],
            warnings=result.warnings,
            currency_code=cfg.currency_code,
        )
        print(f"📦 Excel pack created: {out_path}")
    else:
        out_path = out_dir / "report_pack.xlsx"

    # Charts (generate if either charts are requested OR PDF needs them)
    created: list[Path] = []
    if cfg.write_charts or make_pdf:
        charts_dir = out_dir / "charts"
        created = generate_charts(
            trends=tables["Trends"],
            drill_region=tables["Drilldown_Region"],
            drill_product=tables["Drilldown_Product"],
            out_dir=charts_dir,
        )

        if created:
            print("📊 Charts saved:")
            for p in created:
                print(" -", p)
        else:
            print("📊 No charts generated (missing/empty trend or drilldown data).")

    # PDF (driven by YAML)
    pdf_created = False
    if make_pdf:
        pdf_path = out_dir / "report.pdf"
        try:
            write_pdf_report(
                out_path=pdf_path,
                summary=tables["Summary"],
                trends=tables["Trends"],
                variance=tables["Variance"],
                drill_region=tables["Drilldown_Region"],
                drill_product=tables["Drilldown_Product"],
                chart_paths=created,
                source_label=str(input_dir),
                currency_code=cfg.currency_code,
                report_title=cfg.report_title,
                report_subtitle=cfg.report_subtitle,
                notes=cfg.notes,
                warnings=result.warnings,
            )
        except PermissionError:
            print(f"ERROR: Can't write {pdf_path}. Close it if open, then re-run.", file=sys.stderr)
        else:
            pdf_created = True
            print(f"📄 PDF report created: {pdf_path}")

    # Run log (driven by YAML)
    if cfg.write_run_log:
        log_path = out_dir / "run_log.txt"
        write_run_log(
            log_path,
            input_dir=str(input_dir),
            out_dir=str(out_dir),
            currency=cfg.currency_code,
            rows_loaded=len(df_raw),
            charts_count=len(created),
            pdf_created=pdf_created,
            warnings=result.warnings,
        )
        print(f"🧾 Run log written: {log_path}")

    print(f"✅ Loaded rows: {len(df_raw)} from folder: {input_dir}")
    print(f"💱 Currency code: {cfg.currency_code}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
