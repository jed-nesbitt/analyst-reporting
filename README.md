# Analyst Reporting Automation Suite

A practical Python automation tool that converts messy monthly CSV/XLSX “data dumps” into stakeholder-ready reporting outputs — Excel pack, charts, PDF, data quality report, cleaned export, and a run log — driven by `config.yaml`.

## Outputs
- `out/report_pack.xlsx` (ExecutiveSummary, Summary, Trends, Variance, Drilldowns)
- `out/charts/*.png`
- `out/report.pdf`
- `out/data_quality.xlsx`
- `out/cleaned_data.csv`
- `out/run_log.txt`

## Quick start
```bash
pip install -r requirements.txt
python main.py
