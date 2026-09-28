"""Project 2: Manufacturing OEE and Downtime Root Cause. Cleaning pipeline."""
import glob
import numpy as np
import pandas as pd
from lib.cleaning import Cleaner

SRC = sorted(glob.glob("data/raw/*.xlsx"))[0]

TAXONOMY = [  # governed code, reason, category, six big loss, OEE pillar
    ("DT01", "Mechanical Failure", "Unplanned", "Equipment failure", "Availability"),
    ("DT02", "Electrical Fault", "Unplanned", "Equipment failure", "Availability"),
    ("DT03", "Tooling Issue", "Unplanned", "Equipment failure", "Availability"),
    ("DT04", "Changeover", "Planned", "Setup and adjustment", "Availability"),
    ("DT05", "Material Shortage", "Unplanned", "Idling and minor stops", "Performance"),
    ("DT06", "Quality Hold", "Unplanned", "Process defects", "Quality"),
    ("DT07", "Planned Maintenance", "Planned", "Planned stop", "Availability"),
    ("DT08", "Operator Break", "Planned", "Planned stop", "Availability"),
]
TAX = pd.DataFrame(TAXONOMY, columns=["code", "reason", "category", "big_loss", "pillar"])

BOOL = ["root_cause_identified_flag", "shift_target_met_flag", "quality_hold_flag", "downtime_taxonomy_standardized_flag"]
NUM = lambda **k: dict(**k)
CFG = dict(
    key="record_id", anchor="machine_id", key_prefix="MFG-", key_width=7,
    flag_identity=[("work_order_number", "Flagged in work_order_number_conflict_flag. Work orders can span records, but these repeats carry different lines, SKUs and shifts, so they are routed to planning for review rather than rewritten."),
                   ("batch_lot_number", "Flagged in batch_lot_number_conflict_flag. Lot reuse across records can be legitimate (one lot, several runs); retained unchanged for traceability.")],
    categorical=["plant_site", "production_line", "shift", "shift_supervisor", "product_sku", "downtime_reason_code",
                 "downtime_reason_desc", "downtime_category", "maintenance_type"],
    canonical_override={"production_line": {f"line {x}".casefold(): f"Line-{x.upper()}" for x in
                                             ["a1", "a2", "b1", "b2", "c1", "c2", "d1", "d2", "e1", "e2", "f1", "f2"]}},
    id_text=["record_id", "work_order_number", "machine_id", "operator_id", "batch_lot_number", "fault_code", "technician_id"],
    bool=BOOL, dates=["timestamp_start"],
    numeric={
        "units_produced": NUM(lo=0, unit="units", rationale="Counts cannot be negative; magnitude errors above the fence (up to 25,200) removed"),
        "units_scrapped": NUM(lo=0, unit="units", rationale="Counts cannot be negative; inflated values removed by fence"),
        "downtime_minutes": NUM(lo=0, hi=960, unit="minutes", rationale="Cannot be negative or exceed the longest planned shift; sentinel 240 caught by fence"),
        "oee_pct": NUM(lo=0, hi=100, fence=False, unit="%", rationale="Percentage scale"),
        "availability_pct": NUM(lo=0, hi=100, fence=False, unit="%", rationale="Percentage scale"),
        "performance_pct": NUM(lo=0, hi=100, fence=False, unit="%", rationale="Percentage scale"),
        "quality_pct": NUM(lo=0, hi=100, fence=False, unit="%", rationale="Percentage scale"),
        "planned_production_time_min": NUM(lo=1, hi=1440, fence=False, unit="minutes", rationale="Within one day"),
        "actual_run_time_min": NUM(lo=0, hi=1440, fence=False, unit="minutes", rationale="Within one day"),
        "ideal_cycle_time_sec": NUM(lo=0.1, fence=False, unit="seconds", rationale="Must be positive"),
        "actual_cycle_time_sec": NUM(lo=0.1, unit="seconds", rationale="Must be positive; extreme cycle times (to 600 s) removed by fence"),
        "repair_time_minutes": NUM(lo=0, unit="minutes", rationale="Negative repair impossible; sentinel 168 caught by fence"),
        "six_sigma_defect_count": NUM(lo=0, fence=False, unit="defects", integer=True, rationale="Count"),
        "defects_per_million_opportunities": NUM(lo=0, hi=1000000, fence=False, unit="DPMO", rationale="Definitional range"),
        "changeover_time_minutes": NUM(lo=0, unit="minutes", rationale="Cannot be negative"),
        "energy_consumption_kwh": NUM(lo=0, unit="kWh", rationale="Cannot be negative"),
        "scrap_cost_usd": NUM(lo=0, unit="USD", rationale="Magnitude errors (about 12,000 and above) removed by fence"),
        "line_speed_units_per_hr": NUM(lo=0, fence=False, unit="units/h", rationale="Cannot be negative"),
        "target_units_per_hr": NUM(lo=0.1, fence=False, unit="units/h", rationale="Must be positive"),
        "rework_units": NUM(lo=0, unit="units", rationale="Count"),
        "temperature_c": NUM(lo=0, hi=120, unit="deg C", rationale="Machine operating envelope; sub zero and above 120 are sensor faults"),
        "vibration_mm_s": NUM(lo=0, unit="mm/s", rationale="Velocity RMS is a magnitude and cannot be negative"),
        "pressure_psi": NUM(lo=0, unit="psi", rationale="Gauge pressure in operating range"),
        "humidity_pct": NUM(lo=0, hi=100, fence=False, unit="%", rationale="Percentage scale"),
    },
    impute={
        "units_produced": ("median", "production_line"), "units_scrapped": ("median", "production_line"),
        "downtime_minutes": ("median", "production_line"), "availability_pct": ("median", "production_line"),
        "performance_pct": ("median", "production_line"), "quality_pct": ("median", "production_line"),
        "actual_run_time_min": ("median", "production_line"), "actual_cycle_time_sec": ("median", "production_line"),
        "repair_time_minutes": ("median", "production_line"), "six_sigma_defect_count": ("median", "production_line"),
        "defects_per_million_opportunities": ("median", "production_line"), "changeover_time_minutes": ("median", "production_line"),
        "energy_consumption_kwh": ("median", "production_line"), "scrap_cost_usd": ("median", "production_line"),
        "rework_units": ("median", "production_line"), "temperature_c": ("median", "production_line"),
        "vibration_mm_s": ("median", "production_line"), "pressure_psi": ("median", "production_line"),
        "humidity_pct": ("median", "plant_site"),
        "oee_pct": ("leave", "reported legacy metric; kept only to reconcile against the governed OEE"),
        "shift_supervisor": ("constant", "Not Recorded"), "product_sku": ("constant", "Not Recorded"),
        "timestamp_start": ("leave", "event dates are never imputed; record excluded from weekly series only"),
    },
)

FEATURES = [
    ("reason", "Governed downtime reason (standardised description, the authoritative field)", "=downtime_reason_desc after C4"),
    ("governed_code", "Governed reason code from the taxonomy on M1", "=INDEX(tax_code,MATCH(reason,tax_reason,0))"),
    ("governed_category", "Planned or Unplanned from the taxonomy (replaces random stored category)", "=INDEX(tax_category,MATCH(reason,tax_reason,0))"),
    ("six_big_loss", "Six Big Losses class from the taxonomy", "=INDEX(tax_bigloss,MATCH(reason,tax_reason,0))"),
    ("oee_governed_pct", "Governed OEE = availability x performance x quality", "=availability_pct*performance_pct*quality_pct/10000"),
    ("oee_gap_pts", "Reported oee_pct minus governed OEE (percentage points)", "=oee_pct-oee_governed_pct"),
    ("avail_time_min", "Operating time implied by availability = planned time x A", "=planned_production_time_min*availability_pct/100"),
    ("perf_time_min", "Net run time = operating time x P", "=avail_time_min*performance_pct/100"),
    ("fpt_min", "Fully productive time = net run time x Q; sum(fpt)/sum(planned) is the time weighted OEE", "=perf_time_min*quality_pct/100"),
    ("event_date", "Start date of the production record", "=INT(timestamp_start)"),
    ("week_start", "Monday of the production week", "=event_date-WEEKDAY(event_date,3)"),
    ("month_start", "First day of the production month", "=DATE(YEAR(d),MONTH(d),1)"),
    ("line_shift", "Line and shift key for variance detection", "=production_line&\" | \"&shift"),
    ("scrap_exceeds_output_flag", "Units scrapped greater than units produced (see C8)", "=units_scrapped>units_produced"),
]


def run():
    raw = pd.read_excel(SRC, sheet_name="Dataset", dtype=object)
    dd = pd.read_excel(SRC, sheet_name="Data Dictionary", header=None, dtype=object)
    cl = Cleaner(raw, CFG)
    cl.profile(); cl.structural(); cl.key_collisions(); cl.text(); cl.booleans(); cl.dates(); cl.numerics()
    rmap = dict(zip(TAX.reason.str.casefold(), TAX.code))
    cmap = dict(zip(TAX.reason.str.casefold(), TAX.category))
    checks = [
        dict(Check="Reported OEE not equal to A x P x Q", Description="|oee_pct minus availability x performance x quality| greater than 1 point",
             mask=lambda d: (d.oee_pct - d.availability_pct * d.performance_pct * d.quality_pct / 1e4).abs() > 1, flag="oee_reconciliation_flag",
             Decision="OEE is defined as A x P x Q. The governed OEE is recomputed from the three components and used everywhere; oee_pct is kept only to quantify the reporting gap (08 Validation)."),
        dict(Check="Run time exceeds planned time", Description="actual_run_time_min greater than planned_production_time_min",
             mask=lambda d: d.actual_run_time_min > d.planned_production_time_min,
             Decision="Availability from times would exceed 100 percent. availability_pct (the controlled OEE input) is authoritative; run time is not used to derive availability."),
        dict(Check="Scrap exceeds output", Description="units_scrapped greater than units_produced",
             mask=lambda d: d.units_scrapped > d.units_produced, flag="scrap_exceeds_output_flag",
             Decision="Physically impossible. quality_pct is authoritative for OEE; records flagged and unit counts excluded from any yield calculation."),
        dict(Check="Reason code contradicts description", Description="Stored downtime_reason_code differs from the governed code for the stated description",
             mask=lambda d: d.downtime_reason_code.str.upper() != d.downtime_reason_desc.str.casefold().map(rmap), flag="reason_code_conflict_flag",
             Decision="The description is the operator's statement of what happened and is authoritative. Code, category and loss class are rebuilt from the governed taxonomy (M1)."),
        dict(Check="Category contradicts description", Description="Stored Planned/Unplanned differs from the governed category for the description",
             mask=lambda d: d.downtime_category != d.downtime_reason_desc.str.casefold().map(cmap),
             Decision="Same resolution: governed_category replaces downtime_category in every analysis."),
        dict(Check="Actual cycle faster than ideal", Description="actual_cycle_time_sec below ideal_cycle_time_sec (implies performance above 100 percent)",
             mask=lambda d: d.actual_cycle_time_sec < d.ideal_cycle_time_sec,
             Decision="performance_pct is authoritative and already bounded at 100; cycle times are used descriptively only."),
    ]
    cl.cross_field(checks)
    cl.missing()
    df = cl.df
    df["reason"] = df.downtime_reason_desc
    key = df.reason.str.casefold()
    df["governed_code"] = key.map(rmap)
    df["governed_category"] = key.map(cmap)
    df["six_big_loss"] = key.map(dict(zip(TAX.reason.str.casefold(), TAX.big_loss)))
    df["oee_governed_pct"] = df.availability_pct * df.performance_pct * df.quality_pct / 1e4
    df["oee_gap_pts"] = df.oee_pct - df.oee_governed_pct
    df["avail_time_min"] = df.planned_production_time_min * df.availability_pct / 100
    df["perf_time_min"] = df.avail_time_min * df.performance_pct / 100
    df["fpt_min"] = df.perf_time_min * df.quality_pct / 100
    df["event_date"] = df.timestamp_start.dt.normalize()
    df["week_start"] = df.event_date - pd.to_timedelta(df.event_date.dt.weekday, unit="D")
    df["month_start"] = df.event_date.dt.to_period("M").dt.to_timestamp()
    df["line_shift"] = df.production_line + " | " + df["shift"]
    df["scrap_exceeds_output_flag"] = df.units_scrapped > df.units_produced
    assert df.governed_code.notna().all()
    cl._log("S10", "C10 Feature Engineering", "Derived governed taxonomy, OEE time layers and time keys", len(df), len(df), len(FEATURES) * len(df), f"{len(FEATURES)} features added")
    cl._log("S11", "Clean Data", "Published analysis ready table", len(df), len(df), 0, "Clean Data sheet, with named ranges for model formulas")
    cl.df = df
    return raw, dd, cl


if __name__ == "__main__":
    raw, dd, cl = run()
    d = cl.df
    print(d.shape, pd.DataFrame(cl.audit)[["Step", "RowsOut", "CellsChanged"]].T)
    print(cl.numeric_rules[["Column", "TukeyFence", "AppliedMax", "BelowMin", "AboveMax"]])
    print(cl.cross[["Check", "RecordsFailing", "FailRate"]])
    for c in ["plant_site", "production_line", "shift", "reason", "maintenance_type", "product_sku"]:
        print(c, d[c].value_counts().to_dict())
    print("governed OEE (time weighted)", d.fpt_min.sum() / d.planned_production_time_min.sum(), "reported mean", d.oee_pct.mean())
    print(d.groupby("reason").downtime_minutes.sum().sort_values(ascending=False) / d.downtime_minutes.sum())
    print("dates", d.event_date.min(), d.event_date.max(), d.event_date.isna().sum())
