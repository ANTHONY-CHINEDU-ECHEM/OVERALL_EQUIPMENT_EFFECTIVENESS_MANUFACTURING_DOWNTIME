# Manufacturing OVERALL EQUIPMENT EFFECTIVENESS (OEE) and Downtime Root Cause Command Workbook

An Excel command workbook that unifies downtime reporting into one governed taxonomy, calculates OEE correctly by line and shift, and surfaces root causes, fingerprints and shift anomalies for a 12 line, 3 shift CNC operation. Built from a supplied 15,394 row production extract through a fully documented, eleven step cleaning procedure.


## Business problem

A precision parts manufacturer has OEE stuck around 61 percent. Downtime is logged in incompatible formats, so root causes are invisible and improvement budgets are allocated by guesswork.

## Results at a glance

<table>
  <tr><th>Measure</th><th>Result</th></tr>
  <tr><td>Governed OEE (availability x performance x quality, time weighted)</td><td>41.4 percent</td></tr>
  <tr><td>Reported OEE in the legacy field</td><td>63.0 percent</td></tr>
  <tr><td>Reporting gap</td><td>21.5 points overstated</td></tr>
  <tr><td>Availability, performance, quality</td><td>70.0, 69.7 and 84.9 percent</td></tr>
  <tr><td>Downtime reasons needed for 80 percent of lost minutes</td><td>7 of 8</td></tr>
  <tr><td>Stored reason codes that contradict the description</td><td>87 percent</td></tr>
</table>

## Success metrics

<table>
  <tr><th>Catalogue metric</th><th>Result</th><th>Status</th></tr>
  <tr><td>OEE validated against three months of line data</td><td>36 of 36 line months reproduce an independent Python calculation exactly</td><td>Met</td></tr>
  <tr><td>80 percent of losses traced to five or fewer causes</td><td>7 causes needed; downtime is spread evenly and no line level concentration exists</td><td>Not met (data has no vital few)</td></tr>
  <tr><td>Drill down within two clicks</td><td>Line then shift selection reveals root causes</td><td>Met</td></tr>
  <tr><td>Handles 50,000 plus events</td><td>Tested at 61,000 events: 3 to 8 seconds per selector change versus 1 to 2 seconds at 15,250</td><td>Partly met (linear scaling)</td></tr>
</table>

## Workbook structure

<table>
  <tr><th>Sheet</th><th>Purpose</th></tr>
  <tr><td>01 Executive Summary</td><td>KPIs, success metrics, live findings and recommendations</td></tr>
  <tr><td>02 OEE Command Center</td><td>Plant, line, shift and period selectors; KPI tiles; loss waterfall; monthly trend; line by shift heat map</td></tr>
  <tr><td>03 Root Cause Pareto</td><td>Reason, Six Big Losses and line hotspot Pareto with the 80 percent test</td></tr>
  <tr><td>04 Downtime Fingerprint</td><td>Radar of each line's cause mix, fingerprint distance and a chi square test of independence</td></tr>
  <tr><td>05 Shift Variance Detector</td><td>All 36 line and shift combinations scored against a trailing 13 week baseline</td></tr>
  <tr><td>06 Drill Down Navigator</td><td>Plant to line to shift to root cause</td></tr>
  <tr><td>07 Weekly OEE Scorecard</td><td>One page printable scorecard with a preset print area for PDF export</td></tr>
  <tr><td>08 Validation and Assumptions</td><td>Three month validation, scale test, assumptions and limitations</td></tr>
  <tr><td>M1 and M2</td><td>Governed downtime taxonomy and the time weighted OEE engine</td></tr>
  <tr><td>C1 to C11</td><td>Data cleaning procedure, one step per sheet</td></tr>
</table>

## Method

1. Clean the extract (see docs/data_cleaning_procedure.md) and replace the stored reason code and planned or unplanned category with a governed taxonomy keyed on the operator's description.
2. Compute time layers per record: operating time, net run time and fully productive time. For any group, OEE equals the sum of fully productive time over the sum of planned time, so OEE equals A x P x Q exactly at every level.
3. Rank losses with Pareto analysis at reason, loss class and line by reason level.
4. Test whether line fingerprints differ with a chi square test on event counts (12 lines x 8 reasons).
5. Score each line and shift week with a z score against its previous 13 weeks.

## Key findings

* The legacy OEE field overstates performance by about 21.5 points and cannot be rebuilt from its own components; it should be retired.
* Availability and performance losses dominate; quality is the strongest pillar.
* There is no vital few: the largest reason holds 12.8 percent of downtime and the five worst line and reason hotspots together hold 6.2 percent. Line fingerprints do not differ (p = 0.579), so the loss is systemic and calls for a network wide programme.

## How to use

Open the workbook in Microsoft Excel 2010 or later. Use the teal bordered selectors on 02, 03, 04, 05, 06 and 07. Blue text on yellow marks inputs, black text marks formulas. The selectors are in cell drop down lists (data validation), which behave like combo boxes in every Excel version; a Form Control combo box can be linked to the same cells if preferred.

## Reproducing the build

```
pip install pandas numpy openpyxl scipy
python3 src/p2_build.py
```

Run from the repository root. The script reads data/raw, writes workbook/, and the workbook recalculates when opened.

## Repository structure

```
README.md
workbook/      finished Excel workbook
data/raw/      supplied dataset, unchanged
src/           cleaning engine, sheet writers and build script
docs/          data cleaning procedure and methodology
```

## Limitations

machine_id, operator_id and fault_code are unique per record, so fingerprints are built at line level, the finest grain with repeated observations. Weekly line and shift OEE rests on about three records, so a two sigma flag should prompt a review rather than a conclusion.

## Author

Anthony Chinedu Echem
