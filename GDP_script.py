
#!/usr/bin/env python3
"""
GDP Research Verification Tool — Version 2
Case: US real GDP, Q2 2026.

Input:
    BEA Table 1.1.6 Excel workbook (current revised data).

Important methodological limitation:
    The workbook contains revised data (October 29, 2026).
    It cannot reproduce the original August 26, 2026 release figure
    of 1.5%. The current revised figure is approximately 2.2%.

The tool checks arithmetic, data structure, historical context,
and agreement with the current revised published figure.
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

CASE_ID = "US_GDP_Q2_2026"

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "Table (2).xlsx"

TARGET_Q1 = "2026-01-01"
TARGET_Q2 = "2026-04-01"

# Current revised BEA figure, not the original August 26 release.
OFFICIAL_REVISED_GROWTH = 2.2

# Original August 26 release, retained for historical reference only.
ORIGINAL_RELEASE_GROWTH = 1.5

ROUNDING_DECIMALS = 1
HISTORY_QUARTERS = 20
ANOMALY_Z_THRESHOLD = 2.5

OUTPUT_DIR = BASE_DIR / "gdp_research_output"


# ----------------------------- EXCEL HELPERS -----------------------------

def normalize(value) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def parse_year(value):
    """Return a four-digit year when a cell contains a year."""
    if isinstance(value, (int, float)):
        year = int(value)
        if 1900 <= year <= 2100 and float(value) == year:
            return year

    text = normalize(value)

    if text.isdigit() and len(text) == 4:
        year = int(text)
        if 1900 <= year <= 2100:
            return year

    return None


def parse_quarter(value):
    """Recognize Q1, Q2, Q3, Q4 and common variants."""
    text = normalize(value).replace(" ", "")

    mapping = {
        "q1": 1,
        "1q": 1,
        "q2": 2,
        "2q": 2,
        "q3": 3,
        "3q": 3,
        "q4": 4,
        "4q": 4,
    }

    return mapping.get(text)


def quarter_to_date(year: int, quarter: int) -> str:
    month = (quarter - 1) * 3 + 1
    return f"{year}-{month:02d}-01"


def find_gdp_row(ws):
    """Find the row containing the GDP series label."""
    for row in ws.iter_rows():
        for cell in row:
            text = normalize(cell.value)

            if text == "gross domestic product":
                return cell.row

    raise RuntimeError(
        f"Could not find 'Gross domestic product' in worksheet '{ws.title}'."
    )


def find_quarter_header_row(ws, gdp_row):
    """Find the row above the GDP data with the most Q1-Q4 labels."""
    best_row = None
    best_score = 0

    for row_number in range(1, gdp_row):
        score = sum(
            1 for cell in ws[row_number]
            if parse_quarter(cell.value) is not None
        )

        if score > best_score:
            best_score = score
            best_row = row_number

    if best_row is None or best_score < 2:
        raise RuntimeError(
            "Could not identify the quarterly column headings in the workbook."
        )

    return best_row


def find_year_header_row(ws, quarter_row):
    """Find the closest preceding row containing year headings."""
    best_row = None
    best_score = 0

    for row_number in range(1, quarter_row):
        score = sum(
            1 for cell in ws[row_number]
            if parse_year(cell.value) is not None
        )

        if score > best_score:
            best_score = score
            best_row = row_number

    if best_row is None:
        raise RuntimeError(
            "Could not identify the year headings in the workbook."
        )

    return best_row


def extract_gdp_data(ws):
    """Extract quarterly GDP observations from the BEA worksheet."""
    gdp_row = find_gdp_row(ws)
    quarter_row = find_quarter_header_row(ws, gdp_row)
    year_row = find_year_header_row(ws, quarter_row)

    observations = {}

    current_year = None

    for column in range(1, ws.max_column + 1):
        year_value = ws.cell(year_row, column).value
        parsed_year = parse_year(year_value)

        # Forward-fill years across merged or grouped year headings.
        if parsed_year is not None:
            current_year = parsed_year

        quarter_value = ws.cell(quarter_row, column).value
        quarter = parse_quarter(quarter_value)

        if quarter is None or current_year is None:
            continue

        value = ws.cell(gdp_row, column).value

        if value is None:
            continue

        try:
            numeric_value = float(value)
        except (TypeError, ValueError):
            continue

        period = quarter_to_date(current_year, quarter)
        observations[period] = numeric_value

    if not observations:
        raise RuntimeError(
            "No quarterly GDP observations could be extracted. "
            "The workbook layout may differ from the expected BEA table."
        )

    return dict(sorted(observations.items())), {
        "worksheet": ws.title,
        "gdp_row": gdp_row,
        "year_header_row": year_row,
        "quarter_header_row": quarter_row,
    }


def load_bea_excel(path: Path):
    if not path.exists():
        raise FileNotFoundError(
            f"Excel file not found: {path}\n"
            "Check the file path and filename."
        )

    workbook = load_workbook(
        filename=path,
        read_only=False,
        data_only=True
    )

    errors = []

    for ws in workbook.worksheets:
        try:
            data, layout = extract_gdp_data(ws)
            workbook.close()
            return data, layout, ws.title
        except RuntimeError as exc:
            errors.append(f"{ws.title}: {exc}")

    workbook.close()

    raise RuntimeError(
        "Could not extract GDP data from any worksheet:\n"
        + "\n".join(errors)
    )


# ----------------------------- VALIDATION -----------------------------

def validate_observations(data):
    problems = []

    for period in (TARGET_Q1, TARGET_Q2):
        if period not in data:
            problems.append(
                f"Required quarter {period} is missing from the Excel data."
            )

    for period, value in data.items():
        if not math.isfinite(value):
            problems.append(f"Non-finite GDP value for {period}.")
        elif value <= 0:
            problems.append(f"Non-positive GDP value for {period}: {value}")

    if TARGET_Q1 in data and TARGET_Q2 in data:
        if data[TARGET_Q2] <= data[TARGET_Q1]:
            problems.append("Q2 GDP level is not greater than Q1 GDP level.")

    return problems


# ----------------------------- CALCULATIONS -----------------------------

def quarter_growth(data, current, previous):
    return (data[current] / data[previous] - 1.0) * 100.0


def annualized_growth(data, current, previous):
    return ((data[current] / data[previous]) ** 4 - 1.0) * 100.0


def historical_growth_rates(data):
    periods = sorted(data.keys())
    results = []

    for index in range(1, len(periods)):
        previous = periods[index - 1]
        current = periods[index]

        prev_dt = datetime.strptime(previous, "%Y-%m-%d")
        curr_dt = datetime.strptime(current, "%Y-%m-%d")

        month_gap = (
            (curr_dt.year - prev_dt.year) * 12
            + curr_dt.month - prev_dt.month
        )

        if month_gap != 3:
            continue

        results.append({
            "period": current[:7],
            "previous_period": previous[:7],
            "annualized_growth_pct": annualized_growth(
                data, current, previous
            ),
        })

    return results


def historical_context(data, target_rate):
    rates = historical_growth_rates(data)

    # Use only quarters before the target quarter.
    prior = [
        item for item in rates
        if item["period"] < TARGET_Q2[:7]
    ]

    prior = prior[-HISTORY_QUARTERS:]

    if len(prior) < 8:
        return {
            "observations_used": len(prior),
            "warning": (
                "Too few prior quarters for a meaningful historical "
                "comparison. Interpret results cautiously."
            ),
        }

    values = [
        item["annualized_growth_pct"]
        for item in prior
    ]

    mean_value = statistics.mean(values)
    median_value = statistics.median(values)
    stdev_value = statistics.stdev(values)

    z_score = (
        (target_rate - mean_value) / stdev_value
        if stdev_value else None
    )

    return {
        "observations_used": len(values),
        "period_start": prior[0]["period"],
        "period_end": prior[-1]["period"],
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
        "1. PRIMARY VERIFICATION — CURRENT REVISED DATA",
        "-" * 68,
        f"Q1 2026 GDP level: {primary['q1_level']:.3f}",
        f"Q2 2026 GDP level: {primary['q2_level']:.3f}",
        f"Quarterly growth: {primary['quarterly_growth_pct']:.6f}%",
        f"Annualized growth: {primary['annualized_growth_pct']:.6f}%",
        f"Current revised official figure: "
        f"{OFFICIAL_REVISED_GROWTH:.1f}%",
        f"Calculated result rounded: "
        f"{primary['rounded_calculated_growth_pct']:.1f}%",
        f"Difference before rounding: "
        f"{primary['difference_vs_official_percentage_points']:.6f} pp",
        f"Match after rounding: {primary['matches_after_rounding']}",
        "",
        "2. HISTORICAL CONTEXT",
        "-" * 68,
        json.dumps(context, indent=2, ensure_ascii=False),
        "",
        "3. DATA VINTAGE AND LIMITATIONS",
        "-" * 68,
        f"Original August 26, 2026 release figure: "
        f"{ORIGINAL_RELEASE_GROWTH:.1f}%",
        f"Current revised comparison figure: "
        f"{OFFICIAL_REVISED_GROWTH:.1f}%",
        results["vintage_limitation"],
        "",
        "4. DATA QUALITY",
        "-" * 68,
        json.dumps(
            results["data_quality"],
            indent=2,
            ensure_ascii=False
        ),
        "",
        "Interpretation boundary:",
        "The tool checks arithmetic, data structure, historical "
        "unusualness and agreement with the current revised figure. "
        "It does not independently validate the underlying national "
        "accounts or determine whether government statements are "
        "truthful or misleading.",
    ]

    txt_path.write_text("\n".join(lines), encoding="utf-8")

    print("\n".join(lines))
    print(f"\nSaved JSON: {json_path.resolve()}")
    print(f"Saved report: {txt_path.resolve()}")


# ----------------------------- MAIN -----------------------------

def main():
    print("GDP Research Verification Tool — Version 2")
    print("Reading BEA GDP data from Excel...\n")

    try:
        data, layout, worksheet = load_bea_excel(EXCEL_FILE)

        problems = validate_observations(data)

        if problems:
            raise RuntimeError(
                "Data validation failed:\n" + "\n".join(problems)
            )

        q_growth = quarter_growth(data, TARGET_Q2, TARGET_Q1)
        a_growth = annualized_growth(data, TARGET_Q2, TARGET_Q1)
        rounded = round(a_growth, ROUNDING_DECIMALS)

        primary = {
            "q1_period": TARGET_Q1,
            "q2_period": TARGET_Q2,
            "q1_level": data[TARGET_Q1],
            "q2_level": data[TARGET_Q2],
            "quarterly_growth_pct": q_growth,
            "annualized_growth_pct": a_growth,
            "official_revised_growth_pct": OFFICIAL_REVISED_GROWTH,
            "original_august_release_growth_pct": ORIGINAL_RELEASE_GROWTH,
            "rounded_calculated_growth_pct": rounded,
            "difference_vs_official_percentage_points": (
                a_growth - OFFICIAL_REVISED_GROWTH
            ),
            "matches_after_rounding": (
                rounded == OFFICIAL_REVISED_GROWTH
            ),
        }

        results = {
            "case_id": CASE_ID,
            "generated_utc": datetime.now(timezone.utc).isoformat(),
            "series_id": "BEA Table 1.1.6 — GDP row",
            "series_description": (
                "Real Gross Domestic Product, billions of chained "
                "2017 dollars, seasonally adjusted annual rates"
            ),
            "data_source": "U.S. Bureau of Economic Analysis, Table 1.1.6",
            "input_file": str(EXCEL_FILE),
            "worksheet": worksheet,
            "excel_layout_detected": layout,
            "primary_verification": primary,
            "historical_context": historical_context(data, a_growth),
            "vintage_limitation": (
                "The input workbook contains revised BEA data, last "
                "revised October 29, 2026. It is not the original "
                "August 26, 2026 data vintage. Consequently, this "
                "workbook can verify the revised Q2 2026 growth rate, "
                "but cannot reproduce the original 1.5% release figure. "
                "The original figure is included only as historical "
                "context and must be verified against an archived "
                "August 26 data vintage separately."
            ),
            "data_quality": {
                "excel_file_found": True,
                "worksheet_identified": worksheet,
                "required_quarters_present": True,
                "all_numeric_values_finite_and_positive": True,
                "historical_anomaly_threshold_z": ANOMALY_Z_THRESHOLD,
                "warnings": [
                    "The workbook contains revised data, not the original August 26 vintage.",
                    "The original 1.5% figure is not verified by this workbook.",
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
            "\nNo verification conclusion has been produced. "
            "Check the Excel file path, workbook structure, "
            "and required quarterly observations.",
            file=sys.stderr,
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
