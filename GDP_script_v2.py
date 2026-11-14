#!/usr/bin/env python3
"""
GDP Research Verification Tool — adapted to the supplied BEA workbook.

Input:
    Section1all_xls.xlsx, BEA Table 1.1.1, quarterly worksheet T10101-Q.

Important methodology:
    Table 1.1.1 reports percent change from the preceding period,
    already expressed at seasonally adjusted annual rates (SAAR).
    Therefore, do NOT calculate growth from GDP levels or annualize
    the reported values again.
"""

from __future__ import annotations

import json
import math
import statistics
import sys
from datetime import datetime, timezone
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    print("Missing dependency. Install it with:")
    print("py -m pip install openpyxl")
    raise SystemExit(1)


# ----------------------------- CASE SETTINGS -----------------------------

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "Section1all_xls.xlsx"

SHEET_NAME = "T10101-Q"
CASE_ID = "US_REAL_GDP_Q2_2026"
TARGET_PERIOD = "2026Q2"

# This is the published value in the supplied workbook (Table 1.1.1).
OFFICIAL_GROWTH = 1.5

ROUNDING_DECIMALS = 1
HISTORY_QUARTERS = 20
ANOMALY_Z_THRESHOLD = 2.5
OUTPUT_DIR = BASE_DIR / "gdp_research_output"


# ----------------------------- EXCEL HELPERS -----------------------------

def normalize(value) -> str:
    return "" if value is None else str(value).strip().lower()


def parse_quarter_header(value):
    """Parse headers such as 2026Q2 and return (year, quarter)."""
    text = str(value).strip().upper().replace(" ", "")
    if len(text) == 6 and text[:4].isdigit() and text[4] == "Q":
        year = int(text[:4])
        quarter = int(text[5])
        if 1900 <= year <= 2100 and 1 <= quarter <= 4:
            return year, quarter
    return None


def find_gdp_row(ws):
    """Find the GDP row using the BEA series description."""
    for row in ws.iter_rows():
        for cell in row:
            if normalize(cell.value) == "gross domestic product":
                return cell.row
    raise RuntimeError(
        f"Could not find 'Gross domestic product' in worksheet '{ws.title}'."
    )


def extract_gdp_growth(ws):
    """
    Extract quarterly annualized GDP growth rates from Table 1.1.1.
    Quarter labels are in row 8 and GDP data are in the matching row.
    """
    gdp_row = find_gdp_row(ws)

    # Find the row containing compact quarterly headers (e.g. 1947Q2).
    header_row = None
    best_score = 0
    for row_number in range(1, ws.max_row + 1):
        score = sum(
            parse_quarter_header(cell.value) is not None
            for cell in ws[row_number]
        )
        if score > best_score:
            best_score = score
            header_row = row_number

    if header_row is None or best_score < 2:
        raise RuntimeError("Could not identify quarterly headers (e.g. 2026Q2).")

    observations = {}
    for column in range(1, ws.max_column + 1):
        parsed = parse_quarter_header(ws.cell(header_row, column).value)
        if parsed is None:
            continue

        year, quarter = parsed
        value = ws.cell(gdp_row, column).value
        if value is None:
            continue

        try:
            numeric_value = float(value)
        except (TypeError, ValueError):
            continue

        period = f"{year}Q{quarter}"
        observations[period] = numeric_value

    if not observations:
        raise RuntimeError("No quarterly GDP growth observations were extracted.")

    return dict(sorted(observations.items())), {
        "worksheet": ws.title,
        "gdp_row": gdp_row,
        "quarter_header_row": header_row,
        "series_code": ws.cell(gdp_row, 3).value,
    }


def load_bea_excel(path: Path):
    if not path.exists():
        raise FileNotFoundError(
            f"Excel file not found: {path}\n"
            "Place Section1all_xls.xlsx in the same folder as this script."
        )

    workbook = load_workbook(filename=path, read_only=True, data_only=True)
    try:
        if SHEET_NAME not in workbook.sheetnames:
            raise RuntimeError(
                f"Required worksheet '{SHEET_NAME}' was not found. "
                f"Available sheets include: {', '.join(workbook.sheetnames[:10])}"
            )
        ws = workbook[SHEET_NAME]
        data, layout = extract_gdp_growth(ws)
        return data, layout, ws.title
    finally:
        workbook.close()


# ----------------------------- VALIDATION -----------------------------

def validate_observations(data):
    problems = []

    if TARGET_PERIOD not in data:
        problems.append(f"Required quarter {TARGET_PERIOD} is missing.")

    for period, value in data.items():
        if not math.isfinite(value):
            problems.append(f"Non-finite GDP growth value for {period}.")
        elif value < -100 or value > 100:
            problems.append(
                f"Unusually large GDP growth value for {period}: {value}%. "
                "Check the source data."
            )

    return problems


# ----------------------------- HISTORICAL CONTEXT -----------------------------

def historical_context(data, target_rate):
    periods = sorted(data)
    prior_periods = [period for period in periods if period < TARGET_PERIOD]
    prior_periods = prior_periods[-HISTORY_QUARTERS:]

    if len(prior_periods) < 8:
        return {
            "observations_used": len(prior_periods),
            "warning": (
                "Too few prior quarters for a meaningful historical "
                "comparison. Interpret results cautiously."
            ),
        }

    values = [data[period] for period in prior_periods]
    mean_value = statistics.mean(values)
    median_value = statistics.median(values)
    stdev_value = statistics.stdev(values)
    z_score = (
        (target_rate - mean_value) / stdev_value
        if stdev_value else None
    )

    return {
        "observations_used": len(values),
        "period_start": prior_periods[0],
        "period_end": prior_periods[-1],
        "historical_mean_pct": mean_value,
        "historical_median_pct": median_value,
        "historical_sample_std_dev": stdev_value,
        "target_z_score": z_score,
        "anomaly_screen_flag": (
            abs(z_score) >= ANOMALY_Z_THRESHOLD
            if z_score is not None else None
        ),
        "interpretation": (
            "A screening flag is not proof of a statistical error, "
            "data manipulation, or misleading statement."
        ),
    }


# ----------------------------- REPORTING -----------------------------

def write_report(results):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    json_path = OUTPUT_DIR / f"{CASE_ID}_results.json"
    txt_path = OUTPUT_DIR / f"{CASE_ID}_report.txt"

    with json_path.open("w", encoding="utf-8") as file:
        json.dump(results, file, indent=2, ensure_ascii=False)

    primary = results["primary_verification"]
    context = results["historical_context"]

    lines = [
        f"GDP RESEARCH VERIFICATION — {CASE_ID}",
        "=" * 68,
        f"Generated UTC: {results['generated_utc']}",
        f"Data source: {results['data_source']}",
        f"Workbook: {results['input_file']}",
        f"Worksheet: {results['worksheet']}",
        "",
        "1. PRIMARY VERIFICATION — REPORTED ANNUALIZED GROWTH",
        "-" * 68,
        f"Target quarter: {primary['target_period']}",
        f"Reported real GDP growth: {primary['reported_growth_pct']:.6f}%",
        f"Official comparison figure: {OFFICIAL_GROWTH:.1f}%",
        f"Calculated result rounded: {primary['rounded_growth_pct']:.1f}%",
        f"Difference vs official: {primary['difference_vs_official_pp']:.6f} pp",
        f"Match after rounding: {primary['matches_after_rounding']}",
        "",
        "2. HISTORICAL CONTEXT",
        "-" * 68,
        json.dumps(context, indent=2, ensure_ascii=False),
        "",
        "3. METHODOLOGY AND DATA LIMITATIONS",
        "-" * 68,
        results["methodology"],
        results["vintage_limitation"],
        "",
        "4. DATA QUALITY",
        "-" * 68,
        json.dumps(results["data_quality"], indent=2, ensure_ascii=False),
        "",
        "Interpretation boundary:",
        "This tool checks the extracted reported rate, workbook structure, "
        "historical unusualness and agreement with the comparison figure. "
        "It does not independently validate the underlying national accounts "
        "or determine whether government statements are truthful or misleading.",
    ]

    txt_path.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    print(f"\nSaved JSON: {json_path.resolve()}")
    print(f"Saved report: {txt_path.resolve()}")


# ----------------------------- MAIN -----------------------------

def main():
    print("GDP Research Verification Tool — adapted for BEA Table 1.1.1")
    print("Reading quarterly GDP growth rates from Excel...\n")

    try:
        data, layout, worksheet = load_bea_excel(EXCEL_FILE)
        problems = validate_observations(data)
        if problems:
            raise RuntimeError("Data validation failed:\n" + "\n".join(problems))

        reported_rate = data[TARGET_PERIOD]
        rounded = round(reported_rate, ROUNDING_DECIMALS)

        primary = {
            "target_period": TARGET_PERIOD,
            "reported_growth_pct": reported_rate,
            "official_comparison_growth_pct": OFFICIAL_GROWTH,
            "rounded_growth_pct": rounded,
            "difference_vs_official_pp": reported_rate - OFFICIAL_GROWTH,
            "matches_after_rounding": rounded == OFFICIAL_GROWTH,
            "rate_is_already_annualized": True,
        }

        results = {
            "case_id": CASE_ID,
            "generated_utc": datetime.now(timezone.utc).isoformat(),
            "series_id": "BEA Table 1.1.1 — A191RL",
            "series_description": (
                "Percent change from preceding period in real gross domestic "
                "product, seasonally adjusted at annual rates"
            ),
            "data_source": "U.S. Bureau of Economic Analysis, Table 1.1.1",
            "input_file": str(EXCEL_FILE),
            "worksheet": worksheet,
            "excel_layout_detected": layout,
            "primary_verification": primary,
            "historical_context": historical_context(data, reported_rate),
            "methodology": (
                "The supplied workbook is Table 1.1.1. Its GDP row contains "
                "percent changes from the preceding quarter, already expressed "
                "at seasonally adjusted annual rates (SAAR). The reported rate "
                "is read directly. Do not calculate growth from GDP levels and "
                "do not annualize this rate again."
            ),
            "vintage_limitation": (
                "The workbook states that the data were published August 26, "
                "2026, and the file was created August 25, 2026. This workbook "
                "therefore reflects the vintage supplied in this file; it "
                "should not be described as a later October 29, 2026 revision."
            ),
            "data_quality": {
                "excel_file_found": True,
                "worksheet_identified": worksheet,
                "required_quarter_present": True,
                "all_extracted_values_finite": True,
                "historical_anomaly_threshold_z": ANOMALY_Z_THRESHOLD,
                "warnings": [
                    "Table 1.1.1 contains growth rates, not GDP levels.",
                    "The reported rates are already annualized (SAAR).",
                    "Historical anomaly flags are screening indicators only.",
                    "The workbook is BEA data, not an independent statistical source.",
                ],
            },
        }

        write_report(results)
        return 0

    except Exception as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        print(
            "\nNo verification conclusion has been produced. Check the Excel "
            "file location, worksheet structure, and target quarter.",
            file=sys.stderr,
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
