from __future__ import annotations

from pathlib import Path
import pandas as pd


SUPPORTED = {".csv", ".xlsx", ".xls"}


def list_input_files(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        return []
    files = [p for p in input_dir.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED]
    return sorted(files)


def read_one_file(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        # try utf-8, then fallback
        try:
            df = pd.read_csv(path)
        except UnicodeDecodeError:
            df = pd.read_csv(path, encoding="latin-1")
    else:
        df = pd.read_excel(path)

    # Track lineage
    df["source_file"] = path.name
    return df


def read_all(input_dir: Path) -> pd.DataFrame:
    files = list_input_files(input_dir)
    if not files:
        return pd.DataFrame()

    dfs: list[pd.DataFrame] = []
    for f in files:
        df = read_one_file(f)
        dfs.append(df)

    return pd.concat(dfs, ignore_index=True)
