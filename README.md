# U.S. GDP Verification Research — Q2 and Q3 2026

## Project Overview

This project examines the verification, revision, and communication of official U.S. Gross Domestic Product (GDP) growth estimates for the second and third quarters of 2026.

The research uses data published by the U.S. Bureau of Economic Analysis (BEA), independent manual calculations, Python-based computational verification, and an analysis of media reporting.

The central purpose is to assess the reproducibility of published GDP growth figures and examine how statistical revisions and data communication affect their interpretation.

## Research Objectives

- Verify published GDP growth rates using official BEA data.
- Recalculate annualized quarterly real GDP growth.
- Compare different BEA data vintages and subsequent revisions.
- Examine how media outlets report and frame GDP results.
- Identify the strengths and limitations of computational verification.

## Data and Key Findings

The project examines the following BEA GDP estimates:

- **Q2 2026:** The second estimate, released on August 26, 2026, reported annualized real GDP growth of 1.5%.
- **Q2 2026 revision:** The third estimate, released on September 30, 2026, revised growth to 2.2%.
- **Q3 2026:** The advance estimate, released on October 29, 2026, reported annualized real GDP growth of 3.4%.

The difference between the Q2 estimates reflects revisions to the underlying BEA national accounts data rather than necessarily indicating a computational error.

These observations demonstrate the importance of distinguishing between initial estimates, subsequent revisions, and different data vintages.

## Methodology

The verification process consists of five stages:

1. Obtaining official BEA GDP data and identifying the relevant data vintages.
2. Manually calculating annualized quarterly real GDP growth.
3. Comparing independent calculations with published BEA figures.
4. Running two Python scripts to examine computational reproducibility.
5. Reviewing media coverage and reporting language to assess how GDP results are communicated.

### Annualized GDP Growth Calculation

The annualized quarterly growth rate is calculated using:

\[
g_{\text{annualized}} =
\left[
\left(\frac{GDP_t}{GDP_{t-1}}\right)^4-1
\right]\times100
\]

Where:

- \(GDP_t\) represents real GDP in the current quarter.
- \(GDP_{t-1}\) represents real GDP in the preceding quarter.
- The exponent 4 annualizes the quarterly growth rate.
- The result is expressed as a percentage.

The calculations use comparable real GDP observations from the relevant BEA data vintage.

## Python Verification

The repository includes two Python scripts:

- `GDP_script.py`
- `GDP_script_v2.py`

Both scripts provide computational approaches to examining published GDP growth figures using BEA data.

Their execution and outputs are examined as part of the reproducibility process.

## Repository Structure

The main project files are maintained in the repository:

- `GDP_script.py` — First GDP verification script.
- `GDP_script_v2.py` — Alternative GDP verification script.
- `Section1all_xls.xlsx` — Supporting BEA data.
- `Table (2).xlsx` — Additional supporting data.
- `US_GDP_Verification_Research_Report.pdf` — Full research report.
- `README.md` — Project documentation.

## Important Limitations

Manual calculations and Python scripts use the same underlying BEA source data. Consequently, they provide computational verification but do not independently validate the underlying national accounts or the original statistical data collection process.

Furthermore, BEA GDP estimates are subject to revision as additional source information becomes available.

The Q3 2026 figure examined in this research is an advance estimate and should not be treated as a final estimate.

The distinction between computational reproducibility and independent statistical validation is central to interpreting the project's findings.

## Main Research Report

The full research report is available in:

`US_GDP_Verification_Research_Report.pdf`

It provides further details on the data, calculations, verification procedures, findings, and research limitations.

## Official Data Sources

**U.S. Bureau of Economic Analysis (BEA)**

- [BEA Official GDP Data](https://www.bea.gov/data/gdp/gross-domestic-product)
- [GDP Second Estimate, Q2 2026 — August 26, 2026](https://www.bea.gov/)
- [GDP Third Estimate, Q2 2026 — September 30, 2026](https://www.bea.gov/)
- [GDP Advance Estimate, Q3 2026 — October 29, 2026](https://www.bea.gov/news/2026/gdp-advance-estimate-3rd-quarter-2026)

---

**Research focus:** GDP data verification, statistical revisions, computational reproducibility, and economic data communication.
