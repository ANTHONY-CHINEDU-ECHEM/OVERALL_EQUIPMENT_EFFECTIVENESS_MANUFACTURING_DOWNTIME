# Data Cleaning Procedure, Project 2

This document mirrors the eleven cleaning sheets (C1 to C11) in the workbook. Every figure below was produced by the same Python engine that wrote those sheets, and each sheet repeats the key counts as live Excel formulas so the reconciliation can be checked inside the workbook.

## Guiding principles

1. No row is deleted unless it is completely empty or an exact copy of an earlier row.
2. No value is changed silently: every changed cell is traceable through a flag or audit column (key_repaired_flag, outlier_fields, imputed_fields, and the cross field flags).
3. Where two fields contradict each other, one is declared authoritative, the decision is written down, and the other is kept for audit.
4. Dates, identifiers, survey answers and legacy model outputs are never imputed.

## Step summary

<table>
  <tr><th>Step</th><th>Sheet</th><th>Action</th><th>Rows in</th><th>Rows out</th><th>Cells changed</th></tr>
  <tr><td>S1</td><td>C2 Structural Integrity</td><td>Removed fully blank rows</td><td>15,394</td><td>15,372</td><td>0</td></tr>
  <tr><td>S2</td><td>C2 Structural Integrity</td><td>Removed exact duplicate records</td><td>15,372</td><td>15,250</td><td>0</td></tr>
  <tr><td>S3</td><td>C3 Key Collision Repair</td><td>Repaired duplicate record_id values</td><td>15,250</td><td>15,250</td><td>91</td></tr>
  <tr><td>S4</td><td>C4 Text Standardization</td><td>Collapsed categorical variants to one standard label</td><td>15,250</td><td>15,250</td><td>2,859</td></tr>
  <tr><td>S5</td><td>C5 Boolean Normalization</td><td>Converted 12 token encodings to TRUE/FALSE</td><td>15,250</td><td>15,250</td><td>61,000</td></tr>
  <tr><td>S6</td><td>C6 Date Standardization</td><td>Parsed four mixed date formats to true Excel dates</td><td>15,250</td><td>15,250</td><td>15,098</td></tr>
  <tr><td>S7</td><td>C7 Numeric Validation</td><td>Nullified impossible and extreme values</td><td>15,250</td><td>15,250</td><td>897</td></tr>
  <tr><td>S8</td><td>C8 Cross Field Consistency</td><td>Tested business rules between related fields</td><td>15,250</td><td>15,250</td><td>0</td></tr>
  <tr><td>S9</td><td>C9 Missing Value Treatment</td><td>Filled or explicitly retained every missing value</td><td>15,250</td><td>15,250</td><td>7,149</td></tr>
  <tr><td>S10</td><td>C10 Feature Engineering</td><td>Derived governed taxonomy, OEE time layers and time keys</td><td>15,250</td><td>15,250</td><td>213,500</td></tr>
  <tr><td>S11</td><td>Clean Data</td><td>Published analysis ready table</td><td>15,250</td><td>15,250</td><td>0</td></tr>
</table>

## C1 Data profile

The raw export holds 15,394 rows and 45 columns. Each column was measured for blanks, distinct values and distinct values after normalising case, whitespace, trailing periods and underscores. A gap between those two distinct counts proves the column carries formatting variants.

## C2 Structural integrity

22 rows were completely blank and were removed. 122 rows were exact copies of an earlier row and were removed, keeping the first occurrence. The original Excel row number of every removed row is listed on the sheet.

## C3 Key collision repair

After deduplication, 91 values of record_id were still shared by different records. The collision free anchor machine_id carries the true record sequence, so 91 records were re keyed from it instead of being deleted. The live check on the sheet confirms that no key in Clean Data is duplicated.

Related identifiers that repeat for legitimate reasons are flagged but not changed:

* work_order_number: 150 records flagged. Flagged in work_order_number_conflict_flag. Work orders can span records, but these repeats carry different lines, SKUs and shifts, so they are routed to planning for review rather than rewritten.
* batch_lot_number: 302 records flagged. Flagged in batch_lot_number_conflict_flag. Lot reuse across records can be legitimate (one lot, several runs); retained unchanged for traceability.

## C4 Text standardisation

9 categorical columns were standardised. 155 non standard spellings covering 2,859 cells were mapped to one governed label. Each raw value was normalised (trim, drop trailing period, underscores to spaces, collapse spaces, ignore case) and mapped to the most frequent clean spelling. Identifier columns were trimmed and upper cased.

<table>
  <tr><th>Issue type</th><th>Variants</th><th>Cells</th></tr>
  <tr><td>Inconsistent letter case</td><td>68</td><td>1,550</td></tr>
  <tr><td>Leading or trailing whitespace</td><td>38</td><td>676</td></tr>
  <tr><td>Trailing period</td><td>38</td><td>500</td></tr>
  <tr><td>Underscore used as separator</td><td>11</td><td>133</td></tr>
</table>

## C5 Boolean normalisation

4 flag columns used twelve encodings of yes and no (true, TRUE, yes, Yes, Y, 1 and their negatives). They were converted to native TRUE and FALSE. No unrecognised tokens were found.

## C6 Date standardisation

Dates arrived as text in four formats. Unambiguous patterns were parsed directly; slash dates where both parts are 12 or below were resolved as month first (the convention shown in the data dictionary) and flagged in an _ambiguous_flag column.

<table>
  <tr><th>Column</th><th>Detected pattern</th><th>Records</th><th>Parsed</th></tr>
  <tr><td>timestamp_start</td><td>Day Month abbrev (DD Mon YYYY)</td><td>3,062</td><td>3,062</td></tr>
  <tr><td>timestamp_start</td><td>ISO 8601 (YYYY MM DD)</td><td>3,038</td><td>3,038</td></tr>
  <tr><td>timestamp_start</td><td>Slash ISO (YYYY/MM/DD)</td><td>2,968</td><td>2,968</td></tr>
  <tr><td>timestamp_start</td><td>Ambiguous NN/NN/YYYY resolved as MM/DD/YYYY</td><td>2,402</td><td>2,402</td></tr>
  <tr><td>timestamp_start</td><td>MM/DD/YYYY (day > 12, unambiguous)</td><td>1,819</td><td>1,819</td></tr>
  <tr><td>timestamp_start</td><td>DD/MM/YYYY (day > 12, unambiguous)</td><td>1,809</td><td>1,809</td></tr>
  <tr><td>timestamp_start</td><td>Missing</td><td>152</td><td>0</td></tr>
</table>

## C7 Numeric validation

Each numeric column was tested against physical or definitional limits and, where no hard maximum exists, against the Tukey outer fence (Q3 plus 3 x IQR). Failing cells were blanked (never the whole row) and then treated in C9.

<table>
  <tr><th>Column</th><th>Unit</th><th>Minimum</th><th>Applied maximum</th><th>Below minimum</th><th>Above maximum</th><th>Rationale</th></tr>
  <tr><td>units_produced</td><td>units</td><td>0.00</td><td>9,427.25</td><td>0</td><td>45</td><td>Counts cannot be negative; magnitude errors above the fence (up to 25,200) removed</td></tr>
  <tr><td>units_scrapped</td><td>units</td><td>0.00</td><td>499.00</td><td>0</td><td>45</td><td>Counts cannot be negative; inflated values removed by fence</td></tr>
  <tr><td>downtime_minutes</td><td>minutes</td><td>0.00</td><td>167.32</td><td>120</td><td>109</td><td>Cannot be negative or exceed the longest planned shift; sentinel 240 caught by fence</td></tr>
  <tr><td>oee_pct</td><td>%</td><td>0.00</td><td>100.00</td><td>0</td><td>0</td><td>Percentage scale</td></tr>
  <tr><td>availability_pct</td><td>%</td><td>0.00</td><td>100.00</td><td>0</td><td>0</td><td>Percentage scale</td></tr>
  <tr><td>performance_pct</td><td>%</td><td>0.00</td><td>100.00</td><td>0</td><td>0</td><td>Percentage scale</td></tr>
  <tr><td>quality_pct</td><td>%</td><td>0.00</td><td>100.00</td><td>0</td><td>0</td><td>Percentage scale</td></tr>
  <tr><td>planned_production_time_min</td><td>minutes</td><td>1.00</td><td>1,440.00</td><td>0</td><td>0</td><td>Within one day</td></tr>
  <tr><td>actual_run_time_min</td><td>minutes</td><td>0.00</td><td>1,440.00</td><td>0</td><td>0</td><td>Within one day</td></tr>
  <tr><td>ideal_cycle_time_sec</td><td>seconds</td><td>0.10</td><td></td><td>0</td><td>0</td><td>Must be positive</td></tr>
  <tr><td>actual_cycle_time_sec</td><td>seconds</td><td>0.10</td><td>265.97</td><td>0</td><td>91</td><td>Must be positive; extreme cycle times (to 600 s) removed by fence</td></tr>
  <tr><td>repair_time_minutes</td><td>minutes</td><td>0.00</td><td>126.03</td><td>83</td><td>63</td><td>Negative repair impossible; sentinel 168 caught by fence</td></tr>
  <tr><td>six_sigma_defect_count</td><td>defects</td><td>0.00</td><td></td><td>0</td><td>0</td><td>Count</td></tr>
  <tr><td>defects_per_million_opportunities</td><td>DPMO</td><td>0.00</td><td>1,000,000.00</td><td>0</td><td>0</td><td>Definitional range</td></tr>
  <tr><td>changeover_time_minutes</td><td>minutes</td><td>0.00</td><td>69.53</td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
  <tr><td>energy_consumption_kwh</td><td>kWh</td><td>0.00</td><td>1,058.70</td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
  <tr><td>scrap_cost_usd</td><td>USD</td><td>0.00</td><td>11,355.04</td><td>0</td><td>84</td><td>Magnitude errors (about 12,000 and above) removed by fence</td></tr>
  <tr><td>line_speed_units_per_hr</td><td>units/h</td><td>0.00</td><td></td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
  <tr><td>target_units_per_hr</td><td>units/h</td><td>0.10</td><td></td><td>0</td><td>0</td><td>Must be positive</td></tr>
  <tr><td>rework_units</td><td>units</td><td>0.00</td><td>228.00</td><td>0</td><td>0</td><td>Count</td></tr>
  <tr><td>temperature_c</td><td>deg C</td><td>0.00</td><td>103.07</td><td>35</td><td>76</td><td>Machine operating envelope; sub zero and above 120 are sensor faults</td></tr>
  <tr><td>vibration_mm_s</td><td>mm/s</td><td>0.00</td><td>8.28</td><td>67</td><td>79</td><td>Velocity RMS is a magnitude and cannot be negative</td></tr>
  <tr><td>pressure_psi</td><td>psi</td><td>0.00</td><td>150.91</td><td>0</td><td>0</td><td>Gauge pressure in operating range</td></tr>
  <tr><td>humidity_pct</td><td>%</td><td>0.00</td><td>100.00</td><td>0</td><td>0</td><td>Percentage scale</td></tr>
</table>

## C8 Cross field consistency

<table>
  <tr><th>Rule</th><th>Records failing</th><th>Fail rate</th><th>Decision</th></tr>
  <tr><td>Reported OEE not equal to A x P x Q</td><td>13,773</td><td>90.3 percent</td><td>OEE is defined as A x P x Q. The governed OEE is recomputed from the three components and used everywhere; oee_pct is kept only to quantify the reporting gap (08 Validation).</td></tr>
  <tr><td>Run time exceeds planned time</td><td>5,782</td><td>37.9 percent</td><td>Availability from times would exceed 100 percent. availability_pct (the controlled OEE input) is authoritative; run time is not used to derive availability.</td></tr>
  <tr><td>Scrap exceeds output</td><td>418</td><td>2.7 percent</td><td>Physically impossible. quality_pct is authoritative for OEE; records flagged and unit counts excluded from any yield calculation.</td></tr>
  <tr><td>Reason code contradicts description</td><td>13,335</td><td>87.4 percent</td><td>The description is the operator's statement of what happened and is authoritative. Code, category and loss class are rebuilt from the governed taxonomy (M1).</td></tr>
  <tr><td>Category contradicts description</td><td>7,106</td><td>46.6 percent</td><td>Same resolution: governed_category replaces downtime_category in every analysis.</td></tr>
  <tr><td>Actual cycle faster than ideal</td><td>5,531</td><td>36.3 percent</td><td>performance_pct is authoritative and already bounded at 100; cycle times are used descriptively only.</td></tr>
</table>

## C9 Missing value treatment

<table>
  <tr><th>Column</th><th>Blanks before</th><th>Strategy</th><th>Filled</th><th>Blanks after</th></tr>
  <tr><td>units_produced</td><td>197</td><td>Median within production_line (overall median if group empty)</td><td>197</td><td>0</td></tr>
  <tr><td>units_scrapped</td><td>350</td><td>Median within production_line (overall median if group empty)</td><td>350</td><td>0</td></tr>
  <tr><td>downtime_minutes</td><td>534</td><td>Median within production_line (overall median if group empty)</td><td>534</td><td>0</td></tr>
  <tr><td>availability_pct</td><td>305</td><td>Median within production_line (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>performance_pct</td><td>305</td><td>Median within production_line (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>quality_pct</td><td>305</td><td>Median within production_line (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>actual_run_time_min</td><td>305</td><td>Median within production_line (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>actual_cycle_time_sec</td><td>396</td><td>Median within production_line (overall median if group empty)</td><td>396</td><td>0</td></tr>
  <tr><td>repair_time_minutes</td><td>604</td><td>Median within production_line (overall median if group empty)</td><td>604</td><td>0</td></tr>
  <tr><td>six_sigma_defect_count</td><td>305</td><td>Median within production_line (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>defects_per_million_opportunities</td><td>305</td><td>Median within production_line (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>changeover_time_minutes</td><td>305</td><td>Median within production_line (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>energy_consumption_kwh</td><td>305</td><td>Median within production_line (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>scrap_cost_usd</td><td>389</td><td>Median within production_line (overall median if group empty)</td><td>389</td><td>0</td></tr>
  <tr><td>rework_units</td><td>305</td><td>Median within production_line (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>temperature_c</td><td>416</td><td>Median within production_line (overall median if group empty)</td><td>416</td><td>0</td></tr>
  <tr><td>vibration_mm_s</td><td>451</td><td>Median within production_line (overall median if group empty)</td><td>451</td><td>0</td></tr>
  <tr><td>pressure_psi</td><td>305</td><td>Median within production_line (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>humidity_pct</td><td>305</td><td>Median within plant_site (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>oee_pct</td><td>305</td><td>Left blank: reported legacy metric; kept only to reconcile against the governed OEE</td><td>0</td><td>305</td></tr>
  <tr><td>shift_supervisor</td><td>305</td><td>Explicit category 'Not Recorded'</td><td>305</td><td>0</td></tr>
  <tr><td>product_sku</td><td>152</td><td>Explicit category 'Not Recorded'</td><td>152</td><td>0</td></tr>
  <tr><td>timestamp_start</td><td>152</td><td>Left blank: event dates are never imputed; record excluded from weekly series only</td><td>0</td><td>152</td></tr>
</table>

## C10 Feature engineering

Derived fields are documented on the C10 sheet with their definition, worksheet equivalent and a live populated count.

## C11 Audit log and reconciliation

The final sheet lists every step with rows in, rows out and cells changed, and reconciles live: rows in Raw Data, less blank rows, less exact duplicates, must equal rows in Clean Data. The check returns PASS in the delivered workbook.
