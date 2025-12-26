from __future__ import annotations

import re
from dataclasses import dataclass
import pandas as pd


def _snake(s: str) -> str:
    s = s.strip()
    s = re.sub(r"[^\w]+", "_", s)      # spaces/punct -> _
    s = re.sub(r"_+", "_", s)
    return s.lower().strip("_")


DEFAULT_ALIASES = {
    # date
    "date": "date",
    "month": "date",
    "period": "date",
    "transaction_date": "date",
    "order_date": "date",
    # revenue
    "revenue": "revenue",
    "sales": "revenue",
    "total_sales": "revenue",
    "amount": "revenue",
    "net_sales": "revenue",
    # cost
    "cost": "cost",
    "cogs": "cost",
    "total_cost": "cost",
    # units
    "units": "units",
    "qty": "units",
    "quantity": "units",
    # dimensions
    "region": "region",
    "state": "region",
    "product": "product",
    "sku": "product",
    "category": "product",
}


REQUIRED = ["date", "revenue"]  # minimal to build a usable report


@dataclass(frozen=True)
class CleanResult:
    df: pd.DataFrame
    warnings: list[str]


def standardise_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_snake(str(c)) for c in df.columns]
    return df


def apply_aliases(df: pd.DataFrame, aliases: dict[str, str] | None = None) -> pd.DataFrame:
    df = df.copy()
    aliases = aliases or DEFAULT_ALIASES

    ren: dict[str, str] = {}
    for c in df.columns:
        if c in aliases:
            ren[c] = aliases[c]
    df = df.rename(columns=ren)

    # If duplicates appear after aliasing (e.g. sales + revenue),
    # keep the first non-null across duplicates.
    # Example: two columns both renamed to "revenue".
    if df.columns.duplicated().any():
        # Build a new DF with merged duplicates
        out = pd.DataFrame()
        for col in pd.unique(df.columns):
            cols = [c for c in df.columns if c == col]
            if len(cols) == 1:
                out[col] = df[col]
            else:
                # combine_first left-to-right
                s = df[cols[0]]
                for k in cols[1:]:
                    s = s.combine_first(df[k])
                out[col] = s
        df = out

    return df


def coerce_types(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # ---- Date FIRST (robust: normal strings + Excel serials) ----
    if "date" in df.columns:
        s = df["date"]

        # 1) Normal parsing
        parsed = pd.to_datetime(s, errors="coerce")

        # 2) Excel serial day numbers (often show up as 44927 / "44927.0")
        num = pd.to_numeric(s, errors="coerce")
        excel_like = num.between(20000, 60000)  # ~1954 to ~2064

        if excel_like.any():
            parsed_excel = pd.to_datetime(
                num[excel_like],
                unit="D",
                origin="1899-12-30",
                errors="coerce",
            )
            parsed.loc[excel_like] = parsed_excel

        df["date"] = parsed

    # ---- Then clean strings (EXCLUDE date so we don't break it) ----
    obj_cols = [c for c in df.select_dtypes(include="object").columns if c != "date"]
    for c in obj_cols:
        df[c] = df[c].astype(str).str.strip()
        df.loc[df[c].isin(["", "nan", "None", "NULL", "null"]), c] = pd.NA

    # ---- Numerics ----
    for c in ["revenue", "cost", "units"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    return df



def validate(df: pd.DataFrame) -> list[str]:
    warnings: list[str] = []

    for r in REQUIRED:
        if r not in df.columns:
            warnings.append(f"Missing required column: '{r}'")
        else:
            if df[r].isna().mean() > 0.5:
                warnings.append(f"Column '{r}' has >50% missing values")

    if "date" in df.columns:
        bad = df["date"].isna().sum()
        if bad > 0:
            warnings.append(f"Rows with unparseable dates: {bad}")

    return warnings


def clean(df_raw: pd.DataFrame, aliases: dict[str, str] | None = None) -> CleanResult:
    df = standardise_columns(df_raw)
    df = apply_aliases(df, aliases=aliases)
    df = coerce_types(df)
    warnings = validate(df)
    return CleanResult(df=df, warnings=warnings)
