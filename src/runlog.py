from __future__ import annotations

from pathlib import Path
from datetime import datetime


def write_run_log(
    out_path: Path,
    *,
    input_dir: str,
    out_dir: str,
    currency: str,
    rows_loaded: int,
    charts_count: int,
    pdf_created: bool,
    warnings: list[str] | None,
) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)

    lines: list[str] = []
    lines.append("Analyst Reporting Automation Suite - Run Log")
    lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("")
    lines.append(f"Input dir: {input_dir}")
    lines.append(f"Output dir: {out_dir}")
    lines.append(f"Currency: {currency}")
    lines.append(f"Rows loaded: {rows_loaded}")
    lines.append(f"Charts generated: {charts_count}")
    lines.append(f"PDF created: {pdf_created}")
    lines.append("")
    lines.append("Warnings:")
    if warnings:
        for w in warnings:
            lines.append(f"- {w}")
    else:
        lines.append("- (none)")

    out_path.write_text("\n".join(lines), encoding="utf-8")
