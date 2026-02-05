import argparse
import math
from pathlib import Path

import pandas as pd


def leading_digit(value: float) -> int | None:
    if pd.isna(value):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if number == 0:
        return None
    number = abs(number)
    while number < 1:
        number *= 10
    digit = int(str(number)[0])
    if digit == 0:
        return None
    return digit


def expected_benford_distribution() -> dict[int, float]:
    return {digit: math.log10(1 + 1 / digit) for digit in range(1, 10)}


def analyze_numeric_column(series: pd.Series, sheet_name: str, column_name: str) -> tuple[pd.DataFrame, dict]:
    digits = series.map(leading_digit).dropna().astype(int)
    total = int(digits.shape[0])
    expected = expected_benford_distribution()

    counts = digits.value_counts().reindex(range(1, 10), fill_value=0).sort_index()
    proportions = counts / total if total > 0 else counts.astype(float)

    detail_rows = []
    for digit in range(1, 10):
        detail_rows.append(
            {
                "sheet": sheet_name,
                "column": column_name,
                "digit": digit,
                "count": int(counts[digit]),
                "proportion": float(proportions[digit]) if total > 0 else 0.0,
                "expected_proportion": expected[digit],
                "difference": float(proportions[digit]) - expected[digit] if total > 0 else -expected[digit],
            }
        )

    if total > 0:
        expected_counts = pd.Series({digit: expected[digit] * total for digit in range(1, 10)})
        chi_square = float(((counts - expected_counts) ** 2 / expected_counts).sum())
        mad = float((proportions - pd.Series(expected)).abs().mean())
    else:
        chi_square = 0.0
        mad = 0.0

    summary = {
        "sheet": sheet_name,
        "column": column_name,
        "total_values": total,
        "chi_square": chi_square,
        "mad": mad,
    }

    return pd.DataFrame(detail_rows), summary


def markdown_table(df: pd.DataFrame) -> str:
    headers = [str(col) for col in df.columns]
    rows = df.astype(str).values.tolist()
    header_line = "| " + " | ".join(headers) + " |"
    separator_line = "| " + " | ".join(["---"] * len(headers)) + " |"
    row_lines = ["| " + " | ".join(row) + " |" for row in rows]
    return "\n".join([header_line, separator_line, *row_lines])


def main() -> None:
    parser = argparse.ArgumentParser(description="Run Benford's Law analysis on JE samples.")
    parser.add_argument("--input", default="je_samples.xlsx", help="Path to the Excel file.")
    parser.add_argument("--output", default="benford_output", help="Directory to write analysis output.")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    excel = pd.ExcelFile(input_path)
    detail_frames = []
    summary_rows = []

    for sheet_name in excel.sheet_names:
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        numeric_cols = df.select_dtypes(include="number")
        if numeric_cols.empty:
            continue
        for column_name in numeric_cols.columns:
            detail_df, summary = analyze_numeric_column(numeric_cols[column_name], sheet_name, column_name)
            detail_frames.append(detail_df)
            summary_rows.append(summary)

    detail_df = pd.concat(detail_frames, ignore_index=True) if detail_frames else pd.DataFrame(
        columns=["sheet", "column", "digit", "count", "proportion", "expected_proportion", "difference"]
    )
    summary_df = pd.DataFrame(summary_rows) if summary_rows else pd.DataFrame(
        columns=["sheet", "column", "total_values", "chi_square", "mad"]
    )

    top_deviations = summary_df.sort_values("mad", ascending=False).head(10)

    report_lines = [
        "# Benford Analysis Report",
        "",
        f"Input file: `{input_path}`",
        "",
        "## What this report is",
        "Benford's Law describes how often each leading digit (1 through 9) appears in many real-world datasets.",
        "For example, a leading digit of **1** is expected about **30.1%** of the time, while **9** is expected about **4.6%**.",
        "Large deviations from these expected rates can indicate unusual patterns worth reviewing.",
        "",
        "## How to read the results",
        "- **Leading digit**: the first non-zero digit of a number (e.g., 0.045 → 4, 1200 → 1).",
        "- **Observed proportion**: how often that digit appears in the data.",
        "- **Expected proportion**: Benford's Law expectation for that digit.",
        "- **Difference**: observed minus expected (positive means the digit appears more than expected).",
        "- **MAD (Mean Absolute Deviation)**: average absolute difference across digits; higher values mean larger overall deviation.",
        "- **Chi-square**: another deviation metric; higher values suggest larger differences from expectations.",
        "",
        "## Top columns by deviation (MAD)",
        markdown_table(top_deviations) if not top_deviations.empty else "No numeric data available.",
        "",
        "## Detailed digit breakdown",
        markdown_table(detail_df) if not detail_df.empty else "No numeric data available.",
        "",
        "Report generated by `scripts/benford_analysis.py`.",
    ]

    (output_dir / "benford_report.md").write_text("\n".join(report_lines))


if __name__ == "__main__":
    main()
