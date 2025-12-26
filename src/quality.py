from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import pandas as pd


@dataclass(frozen=True)
class QualityReport:
    overview: pd.DataFrame
    missing_by_col: pd.DataFrame
    duplicates: pd.DataFrame
    date_range: pd.DataFrame
    categorical_profile: pd.DataFrame


def _safe_nunique(s: pd.Series) -> int:
    try:
        return int(s.nunique(dropna=True))
    except Exception:
        return 0


def build_quality_report(df: pd.DataFrame) -> QualityReport:
    n_rows, n_cols = df.shape

    # Overview
    overview = pd.DataFrame(
        {
            "metric": ["rows", "columns"],
            "value": [n_rows, n_cols],
        }
    )

    # Missingness
    missing_by_col = pd.DataFrame(
        {
            "column": df.columns.astype(str),
            "dtype": df.dtypes.astype(str).values,
            "missing_count": df.isna().sum().values,
            "missing_pct": (df.isna().mean() * 100).round(2).values,
            "n_unique": [ _safe_nunique(df[c]) for c in df.columns ],
        }
    ).sort_values(["missing_count", "missing_pct"], ascending=False)

    # Duplicates (whole-row)
    dup_count = int(df.duplicated().sum())
    duplicates = pd.DataFrame(
        {
            "metric": ["duplicate_rows"],
            "value": [dup_count],
        }
    )

    # Date range (if date exists)
    if "date" in df.columns:
        d = pd.to_datetime(df["date"], errors="coerce")
        date_range = pd.DataFrame(
            {
                "metric": ["min_date", "max_date", "rows_with_valid_date", "rows_with_invalid_date"],
                "value": [
                    str(d.min()) if d.notna().any() else "",
                    str(d.max()) if d.notna().any() else "",
                    int(d.notna().sum()),
                    int(d.isna().sum()),
                ],
            }
        )
    else:
        date_range = pd.DataFrame({"metric": ["date_column_present"], "value": [False]})

    # Categorical profile (top categories for key dims if present)
    cat_cols = [c for c in ["region", "product", "source_file"] if c in df.columns]
    rows = []
    for c in cat_cols:
        vc = df[c].fillna("Unknown").astype(str).value_counts().head(10)
        for k, v in vc.items():
            rows.append({"column": c, "category": k, "count": int(v)})

    categorical_profile = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["column", "category", "count"])

    return QualityReport(
        overview=overview,
        missing_by_col=missing_by_col,
        duplicates=duplicates,
        date_range=date_range,
        categorical_profile=categorical_profile,
    )


def write_quality_excel(out_path: Path, qr: QualityReport) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        qr.overview.to_excel(w, sheet_name="Overview", index=False)
        qr.date_range.to_excel(w, sheet_name="DateRange", index=False)
        qr.duplicates.to_excel(w, sheet_name="Duplicates", index=False)
        qr.missing_by_col.to_excel(w, sheet_name="Missingness", index=False)
        qr.categorical_profile.to_excel(w, sheet_name="TopCategories", index=False)
