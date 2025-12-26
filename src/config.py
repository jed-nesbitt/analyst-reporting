from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


@dataclass(frozen=True)
class AppConfig:
    input_dir: Path = Path("data/in")
    out_dir: Path = Path("out")
    currency_code: str = "AUD"
    aliases: dict[str, str] = field(default_factory=dict)

    # Report text
    report_title: str = "Analyst Reporting Pack"
    report_subtitle: str = ""
    notes: list[str] = field(default_factory=list)

    # Run toggles
    make_pdf: bool = True
    write_excel_pack: bool = True
    write_charts: bool = True
    write_cleaned_csv: bool = True
    write_quality_report: bool = True
    write_run_log: bool = True


def load_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise ValueError("Config must be a YAML mapping (key: value)")
    return data


def _as_bool(x: Any, default: bool) -> bool:
    if x is None:
        return default
    if isinstance(x, bool):
        return x
    if isinstance(x, (int, float)):
        return bool(x)
    s = str(x).strip().lower()
    if s in {"true", "yes", "y", "1", "on"}:
        return True
    if s in {"false", "no", "n", "0", "off"}:
        return False
    return default


def _parse_notes(raw_notes: Any) -> list[str]:
    if raw_notes is None:
        return []
    if isinstance(raw_notes, str):
        s = raw_notes.strip()
        return [s] if s else []
    if isinstance(raw_notes, list):
        out: list[str] = []
        for x in raw_notes:
            if x is None:
                continue
            s = str(x).strip()
            if s:
                out.append(s)
        return out
    raise ValueError("config.yaml notes must be a string or a list of strings")


def resolve_config(
    *,
    config_path: str | None,
    cli_input: str | None,
    cli_out: str | None,
    cli_currency: str | None,
) -> AppConfig:
    # If user doesn't provide a path, we default to config.yaml at project root
    path = Path(config_path or "config.yaml")
    raw = load_config(path)

    input_dir = Path(raw.get("input_dir", "data/in"))
    out_dir = Path(raw.get("out_dir", "out"))
    currency_code = str(raw.get("currency_code", "AUD")).upper().strip()

    aliases_raw = raw.get("aliases", {}) or {}
    if not isinstance(aliases_raw, dict):
        raise ValueError("config.yaml aliases must be a mapping")
    aliases = {str(k): str(v) for k, v in aliases_raw.items()}

    report_title = str(raw.get("report_title", "Analyst Reporting Pack")).strip() or "Analyst Reporting Pack"
    report_subtitle = str(raw.get("report_subtitle", "")).strip()
    notes = _parse_notes(raw.get("notes"))

    # toggles
    make_pdf = _as_bool(raw.get("make_pdf"), True)
    write_excel_pack = _as_bool(raw.get("write_excel_pack"), True)
    write_charts = _as_bool(raw.get("write_charts"), True)
    write_cleaned_csv = _as_bool(raw.get("write_cleaned_csv"), True)
    write_quality_report = _as_bool(raw.get("write_quality_report"), True)
    write_run_log = _as_bool(raw.get("write_run_log"), True)

    cfg = AppConfig(
        input_dir=input_dir,
        out_dir=out_dir,
        currency_code=currency_code,
        aliases=aliases,
        report_title=report_title,
        report_subtitle=report_subtitle,
        notes=notes,
        make_pdf=make_pdf,
        write_excel_pack=write_excel_pack,
        write_charts=write_charts,
        write_cleaned_csv=write_cleaned_csv,
        write_quality_report=write_quality_report,
        write_run_log=write_run_log,
    )

    # CLI overrides config (optional)
    input_dir_final = Path(cli_input) if cli_input else cfg.input_dir
    out_dir_final = Path(cli_out) if cli_out else cfg.out_dir
    currency_final = (cli_currency.upper().strip() if cli_currency else cfg.currency_code)

    return AppConfig(
        input_dir=input_dir_final,
        out_dir=out_dir_final,
        currency_code=currency_final,
        aliases=cfg.aliases,
        report_title=cfg.report_title,
        report_subtitle=cfg.report_subtitle,
        notes=cfg.notes,
        make_pdf=cfg.make_pdf,
        write_excel_pack=cfg.write_excel_pack,
        write_charts=cfg.write_charts,
        write_cleaned_csv=cfg.write_cleaned_csv,
        write_quality_report=cfg.write_quality_report,
        write_run_log=cfg.write_run_log,
    )
