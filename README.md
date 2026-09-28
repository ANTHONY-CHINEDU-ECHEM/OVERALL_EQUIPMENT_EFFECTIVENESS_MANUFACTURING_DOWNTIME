# Manufacturing OEE and Downtime Root Cause Command Workbook

An Excel command workbook that unifies downtime reporting into one governed taxonomy, calculates Overall Equipment Effectiveness (OEE) correctly by line and shift, and surfaces root causes, downtime fingerprints and shift anomalies for a 12 line, 3 shift CNC machining operation. The workbook is built from a supplied 15,394 row production extract through a fully documented, eleven step cleaning procedure, and it runs entirely in native Excel with no add ins, macros or external services.

## Project brief

In precision machining, capacity is bought once and paid for every hour of every shift. A CNC line that stands idle, runs below its rated speed or produces parts that fail inspection still carries its full cost of depreciation, labour, energy and floor space. Overall Equipment Effectiveness is the standard measure of how much of that paid for capacity turns into good parts, combining availability (was the machine running when it was scheduled to), performance (did it run at its ideal rate) and quality (were the parts right first time). The manufacturer in this study runs 12 CNC lines across three shifts and reports OEE stuck at around 61 percent, well short of the level generally regarded as world class, and months of effort have failed to move it.

The reason improvement has stalled is not a lack of effort but a lack of reliable information. Supervisors record downtime in six incompatible formats, with free text descriptions, inconsistent reason codes and no shared definition of what counts as planned or unplanned. When the same stoppage can be recorded as a tooling issue on one shift, a setup delay on the next and an unexplained stop on the third, no one can say with confidence which causes are costing the most time. Improvement budgets are consequently allocated by the loudest voice in the room or the most recent breakdown, and a headline OEE figure is reported upwards without anyone being able to trace how it was calculated or whether it can be trusted.

This project was commissioned to give the plant a single, governed version of the truth about lost capacity. The brief called for a command workbook that standardises every downtime event into one controlled taxonomy, calculates availability, performance, quality and OEE correctly at every level from plant to line to shift, and isolates the true root causes behind each type of loss through Pareto analysis. It also asked for two capabilities rarely attempted in a spreadsheet: a downtime fingerprint that compares the cause mix of each line, and a shift variance detector that flags any line and shift combination drifting away from its own recent baseline. The workbook had to be fast enough for daily use, navigable in two clicks, validated against an independent calculation, and transparent enough that plant managers, engineers and finance could audit every number it produces.

## Objectives

* Replace inconsistent reason codes and categories with one governed downtime taxonomy keyed on the operator's own description.
* Calculate availability, performance, quality and OEE so that the components multiply exactly to the composite at every level of aggregation.
* Identify whether a small number of causes drive most of the lost time, using Pareto analysis at reason, loss class and line hotspot level.
* Compare the downtime fingerprint of each line and test statistically whether lines genuinely differ.
* Flag line and shift combinations that deviate from a trailing 13 week baseline.
* Provide drill down navigation from plant to line to shift to root cause, and a printable weekly scorecard.

## What makes this workbook different

Most OEE spreadsheets average percentages across lines and shifts, which quietly distorts the result whenever run times differ. This workbook works in time rather than percentages. For every record it computes operating time, net run time and fully productive time, so that for any group OEE equals the sum of fully productive time divided by the sum of planned time. Because the same time layers feed every component, availability multiplied by performance multiplied by quality equals OEE exactly at plant, line, shift and week level. On top of this engine sit a governed taxonomy, a chi square test of line fingerprints and a z score based anomaly detector, all in native worksheet formulas.

## Data

The supplied production extract contains 15,394 rows covering downtime events and production records across 12 lines and 3 shifts, including planned time, run time, output, rejects, a stored reason code, a planned or unplanned category, an operator description and a legacy OEE field. Cleaning followed eleven documented steps, one per worksheet (C1 to C11).

<img width="1102" height="340" alt="Screenshot 2026-09-28 at 23 06 11" src="https://github.com/user-attachments/assets/5c10521c-8c84-4a44-b41f-72ef36fb3d92" />


Three data issues shaped the analysis and are worth stating plainly:

* 87 percent of stored reason codes contradict the operator's written description of the same event. The stored code and the planned or unplanned category were therefore replaced with a governed taxonomy derived from the description.
* The legacy OEE field cannot be rebuilt from its own availability, performance and quality components, so it was retained only for comparison.
* machine_id, operator_id and fault_code are unique on every record, so they carry no repeated observations. Fingerprints are therefore built at line level, the finest grain the data can support.

<img width="1144" height="292" alt="Screenshot 2026-09-28 at 23 07 21" src="https://github.com/user-attachments/assets/51945ad4-beb0-4d85-a178-6f46cd067369" />

## Results at a glance

<table>
  <tr><th>Measure</th><th>Result</th></tr>
  <tr><td>Governed OEE (availability x performance x quality, time weighted)</td><td>41.4 percent</td></tr>
  <tr><td>Reported OEE in the legacy field</td><td>63.0 percent</td></tr>
  <tr><td>Reporting gap</td><td>21.5 points overstated</td></tr>
  <tr><td>Availability</td><td>70.0 percent</td></tr>
  <tr><td>Performance</td><td>69.7 percent</td></tr>
  <tr><td>Quality</td><td>84.9 percent</td></tr>
  <tr><td>Downtime reasons needed for 80 percent of lost minutes</td><td>7 of 8</td></tr>
  <tr><td>Share of downtime held by the largest single reason</td><td>12.8 percent</td></tr>
  <tr><td>Share held by the five worst line and reason hotspots combined</td><td>6.2 percent</td></tr>
  <tr><td>Stored reason codes that contradict the description</td><td>87 percent</td></tr>
  <tr><td>Line fingerprint difference (chi square test)</td><td>Not significant, p = 0.579</td></tr>
</table>

## Success metrics

<table>
  <tr><th>Target</th><th>Result</th><th>Status</th></tr>
  <tr><td>OEE validated against three months of line data</td><td>36 of 36 line months reproduce an independent Python calculation exactly</td><td>Met</td></tr>
  <tr><td>80 percent of losses traced to five or fewer causes</td><td>7 causes needed; downtime is spread evenly and no line level concentration exists</td><td>Not met (the data has no vital few)</td></tr>
  <tr><td>Drill down reachable within two clicks</td><td>Line then shift selection reveals root causes</td><td>Met</td></tr>
  <tr><td>Handles 50,000 plus events without lag</td><td>Tested at 61,000 events: 3 to 8 seconds per selector change, compared with 1 to 2 seconds at 15,250</td><td>Partly met (scaling is linear)</td></tr>
</table>

## Findings in detail

### 1. The plant is performing far worse than it believes

The governed calculation puts plant OEE at 41.4 percent, against 63.0 percent in the legacy reporting field. That gap of 21.5 points is not a rounding difference; it means roughly a third of the performance leadership believed it had does not exist. The legacy figure cannot be reconstructed from its own components, which suggests it has been averaged, overwritten or calculated inconsistently over time. The practical consequence is significant. Targets, bonuses, capacity plans and investment cases built on the reported figure have all been anchored to a level of effectiveness the plant is not achieving. Put simply, fewer than half of the scheduled machine hours are currently producing good parts at the ideal rate.

<img width="550" height="349" alt="Screenshot 2026-09-28 at 23 08 18" src="https://github.com/user-attachments/assets/d49e0642-a731-4951-8abd-afb51b18bca3" />

### 2. Availability and performance are where the capacity is being lost

Breaking OEE into its components shows availability at 70.0 percent and performance at 69.7 percent, with quality at 84.9 percent. Around 30 percent of scheduled time is lost to stoppages, and when machines are running they deliver only about 70 percent of their ideal output. These two pillars together account for most of the gap between the plant and a healthy OEE, so they should receive the bulk of improvement attention. Quality is the strongest of the three, but it should not be dismissed: in precision machining, losing around 15 percent of output to scrap or rework is still a material cost, and it compounds the other losses because every rejected part has already consumed machine time.

<img width="616" height="350" alt="Screenshot 2026-09-28 at 23 09 26" src="https://github.com/user-attachments/assets/ee3e2dce-1147-4cfe-80a4-75de73f9837a" />

### 3. The downtime records cannot currently be trusted

In 87 percent of events, the stored reason code contradicts what the operator actually wrote about the stoppage. Any Pareto chart, maintenance plan or supplier conversation built on those codes has been pointing at the wrong causes. This is arguably the most important finding for the organisation, because it explains why years of reporting have not translated into improvement. The governed taxonomy in this workbook corrects the historical record, but the lasting fix is at the point of capture: a single controlled list of reasons, used consistently by every shift.

<img width="700" height="311" alt="Screenshot 2026-09-28 at 23 10 10" src="https://github.com/user-attachments/assets/2148d45e-acc3-41fd-9381-6fba555cb52a" />


### 4. There is no vital few, and that changes the improvement strategy

The classic expectation, and the brief's target, was that 80 percent of lost time would trace to five or fewer causes. The data does not support that. Seven of the eight reasons are needed to reach 80 percent of lost minutes, the largest single reason holds only 12.8 percent of downtime, and the five worst line and reason hotspots combined hold just 6.2 percent. This is a genuine finding, not a failure of the analysis. It tells leadership that targeting one or two causes or one problem line will not move OEE meaningfully, and that funding a handful of point fixes would likely disappoint.

<img width="782" height="302" alt="Screenshot 2026-09-28 at 23 11 06" src="https://github.com/user-attachments/assets/b4aace95-f70e-470e-aaa2-00bfc5a13c06" />

### 5. Every line has the same downtime fingerprint

The fingerprint analysis compares the mix of downtime causes across all 12 lines, and a chi square test of independence on event counts (12 lines by 8 reasons) finds no significant difference between them (p = 0.579). No line is an outlier with its own distinctive problem. Taken together with the flat Pareto, this points to losses that are systemic, rooted in shared practices such as changeover methods, maintenance routines, material supply and shift handovers, rather than in individual machines. The appropriate response is a plant wide improvement programme, for example standardised work, a structured changeover reduction approach and a common autonomous maintenance routine, rather than line by line firefighting.

<img width="778" height="258" alt="Screenshot 2026-09-28 at 23 12 13" src="https://github.com/user-attachments/assets/9fb73bed-46b9-4a6a-9f20-09d894ed4bfc" />

### 6. Shift anomalies should prompt questions, not conclusions

The shift variance detector scores all 36 line and shift combinations each week against their own trailing 13 week baseline and flags deviations beyond two standard deviations. This gives supervisors an early signal when a crew or line drifts. However, weekly line and shift OEE rests on around three records, so individual flags are sensitive to single events. The detector is best used as a prompt for a conversation at the next production meeting rather than as evidence of poor performance by a particular shift.

<img width="701" height="361" alt="Screenshot 2026-09-28 at 23 13 17" src="https://github.com/user-attachments/assets/ee255ecb-ed47-4aa3-9e60-1ca9e52ac56e" />

<img width="773" height="394" alt="Screenshot 2026-09-28 at 23 14 21" src="https://github.com/user-attachments/assets/86b1f305-88e4-43c3-bcd4-9772dce028b3" />

<img width="589" height="373" alt="Screenshot 2026-09-28 at 23 15 22" src="https://github.com/user-attachments/assets/c8467c4c-0cfc-4593-b287-80e417cfff83" />

### 7. The calculation is proven, and the workbook scales predictably

The time weighted OEE engine reproduces an independent Python calculation exactly for all 36 line months in the three month validation window, which gives the plant a figure it can defend to finance and to customers. Performance testing at 61,000 events, four times the supplied volume, showed selector changes taking 3 to 8 seconds compared with 1 to 2 seconds at current volume. Scaling is linear and the workbook remains usable, but a plant generating events at that rate would benefit from moving the engine to Power Pivot or a database in time.

<img width="1204" height="357" alt="Screenshot 2026-09-28 at 23 16 23" src="https://github.com/user-attachments/assets/7aa12d99-4f2f-47c9-8eec-d612c571d1ef" />

<img width="1219" height="486" alt="Screenshot 2026-09-28 at 23 17 49" src="https://github.com/user-attachments/assets/73394020-55dc-4f98-9c79-f11a4b27797b" />


## Recommendations

1. Retire the legacy OEE field immediately and report the governed, time weighted OEE as the single official figure, resetting targets to the true baseline of 41.4 percent.
2. Introduce one controlled downtime reason list at the point of capture across all shifts, with the free text description kept alongside for audit.
3. Focus improvement on availability and performance, where the largest share of capacity is lost, starting with changeover reduction and speed loss analysis.
4. Fund a plant wide improvement programme rather than a few targeted fixes, since losses are evenly spread and no line stands out.
5. Use the weekly scorecard and shift variance detector as standing items in production meetings, treating flags as prompts for review.
6. Capture machine, operator and fault identifiers consistently so that future fingerprints can be built at machine level.

## Method

1. **Clean and govern.** Eleven documented steps (see docs/data_cleaning_procedure.md). The stored reason code and planned or unplanned category are replaced with a governed taxonomy keyed on the operator's description.
2. **Build the time layers.** For every record, operating time, net run time and fully productive time are computed. For any group, OEE equals the sum of fully productive time over the sum of planned time, so availability multiplied by performance multiplied by quality equals OEE exactly at every level.
3. **Rank losses.** Pareto analysis at reason, Six Big Losses class and line by reason level, each with an explicit 80 percent test.
4. **Compare fingerprints.** A radar view of each line's cause mix, a fingerprint distance measure and a chi square test of independence on event counts (12 lines by 8 reasons).
5. **Detect anomalies.** Each line and shift week is scored with a z score against its previous 13 weeks.
6. **Validate.** Line month OEE is compared with an independent Python calculation over three months, and the workbook is scale tested at 61,000 events.

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

## How to use

Open the workbook in Microsoft Excel 2010 or later; calculation is automatic.

* Use the teal bordered selectors on sheets 02, 03, 04, 05, 06 and 07 to choose plant, line, shift and period.
* To drill down, select a line and then a shift on sheet 06 to reveal the root causes behind its losses.
* Formatting convention: blue text on yellow marks inputs, and black text marks formulas.
* Sheet 07 has a preset print area, so the weekly scorecard can be exported to PDF directly.

The selectors are in cell drop down lists (data validation), which behave like combo boxes in every Excel version. A Form Control combo box can be linked to the same cells if preferred.

## Reproducing the build

The workbook is generated by Python, so every sheet can be rebuilt from the raw data.

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

## Assumptions and limitations

* **Fingerprint grain.** machine_id, operator_id and fault_code are unique per record, so fingerprints are built at line level, the finest grain with repeated observations.
* **Governed taxonomy.** Reasons are derived from operator descriptions. Where a description is vague, the classification reflects the best available reading rather than a confirmed cause.
* **Thin weekly samples.** Weekly line and shift OEE rests on about three records, so a two sigma flag should prompt a review rather than a conclusion.
* **Brief versus extract.** The brief cites OEE of around 61 percent; the legacy field in the supplied extract averages 63.0 percent. Both overstate the governed figure of 41.4 percent.
* **Scale.** Response time grows linearly with event volume; beyond roughly 50,000 events a data model or database back end is advisable.
