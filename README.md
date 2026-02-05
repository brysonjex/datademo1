# datademo1

## JE sample analysis workflow

This repo includes a GitHub Actions workflow that produces basic summary statistics and Benford's Law analysis for `je_samples.xlsx`.

### What it generates

The workflow runs `scripts/analyze_je_samples.py` and writes output files to `analysis_output/`:

- `summary.md`: high-level overview with per-sheet row/column counts.
- `sheet_summary.csv`: row/column counts per sheet.
- `<sheet>_column_summary.csv`: descriptive stats for all columns.
- `<sheet>_numeric_stats.csv`: numeric summary stats.
- `<sheet>_date_ranges.csv`: date range summaries.

It also runs `scripts/benford_analysis.py` and writes output files to `benford_output/`:

- `summary.md`: high-level Benford summary.
- `benford_detail.csv`: per-digit counts and proportions per numeric column.
- `benford_summary.csv`: chi-square and MAD per numeric column.

### How to run

1. Push changes to `je_samples.xlsx` (or run the workflow manually via **Actions → JE Sample Analysis → Run workflow**).
2. Download the `je-analysis-output` and `je-benford-output` artifacts from the workflow run to view the outputs.

### Running locally

```bash
python -m pip install -r requirements.txt
python scripts/analyze_je_samples.py --input je_samples.xlsx --output analysis_output
python scripts/benford_analysis.py --input je_samples.xlsx --output benford_output
```
