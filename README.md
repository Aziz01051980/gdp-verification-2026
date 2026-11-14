# U.S. GDP Verification Research — Q2 and Q3 2026

## Project Overview

This project examines the verification and communication of official U.S. Gross Domestic Product (GDP) growth figures for the second and third quarters of 2026.

The research uses data published by the U.S. Bureau of Economic Analysis (BEA), manual calculations, Python scripts, and media reporting analysis.

## Research Objectives

- Verify published GDP growth rates using BEA data.
- Recalculate annualized quarterly real GDP growth.
- Compare different BEA data vintages and revisions.
- Examine how media outlets report and frame GDP results.
- Identify the strengths and limitations of computational verification.

## Data and Findings

The project examines two BEA data releases:

- **Q2 2026:** The August 26, 2026 release reported annualized real GDP growth of 1.5%.
- **Q3 2026:** The October 29, 2026 advance estimate reported annualized real GDP growth of 3.4%, with Q2 revised to 2.2%.

The difference between the two Q2 figures reflects a revision to the BEA estimate, rather than a calculation error.

## Methodology

The verification process includes:

1. Obtaining official BEA GDP data.
2. Manually calculating annualized quarterly growth.
3. Comparing calculations with published BEA figures.
4. Reviewing Python scripts used for verification.
5. Comparing media coverage and reporting language.

The annualized growth calculation is:

\[
\left(\frac{\text{Current Quarter GDP}}{\text{Previous Quarter GDP}}\right)^4 - 1
\]

## Repository Structure

- `report/` — Final research report.
- `data/` — Official BEA data files.
- `scripts/` — Python verification scripts.
- `results/` — Calculation results and findings.

## Important Limitation

Manual calculations and Python scripts use the same underlying BEA data. They provide computational checks, but do not independently validate the underlying national accounts.

The Q3 2026 figure is an advance estimate and may be revised.

## Main Report

See the `report/` folder for the full research report.

## Source

U.S. Bureau of Economic Analysis (BEA): https://www.bea.gov/
