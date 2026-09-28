"""Build the Project 2 workbook: Manufacturing OEE and Downtime Root Cause Command Workbook."""
import os, sys
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, BarChart, RadarChart, Reference
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule, DataBarRule
from lib.xl import *
from lib import cleansheets as CS
from p2_clean import run, FEATURES, TAX

OUT = "workbook/P2_Manufacturing_OEE_Downtime_Command_Workbook.xlsx"
SCALE = sys.argv[1] if len(sys.argv) > 1 else "61,000 events: 3 to 8 s per selector"

raw, dd, cl = run()
df = cl.df
orig = list(raw.columns)
extra = ["key_repaired_flag", "work_order_number_conflict_flag", "batch_lot_number_conflict_flag", "timestamp_start_ambiguous_flag",
         "outlier_fields", "outlier_count", "oee_reconciliation_flag", "reason_code_conflict_flag", "imputed_fields", "imputed_count"] + [f[0] for f in FEATURES]
clean = df[["source_row"] + orig + extra].copy()
TILE = int(os.environ.get("TILE", "1"))
if TILE > 1:
    parts = []
    for k in range(TILE):
        t = clean.copy(); t["record_id"] = t["record_id"] + ("" if k == 0 else f"T{k}"); parts.append(t)
    clean = pd.concat(parts, ignore_index=True)
    OUT = "/tmp/p2_scale.xlsx"
PLANTS = sorted(clean.plant_site.unique())
LINES = sorted(clean.production_line.unique())
SHIFTS = ["Day", "Swing", "Night"]
REASONS = list(TAX.reason)
LOSSES = ["Equipment failure", "Setup and adjustment", "Idling and minor stops", "Process defects", "Planned stop"]
END = clean.event_date.max()
weeks_all = sorted(clean.week_start.dropna().unique())
weeks_all = [pd.Timestamp(w) for w in weeks_all]

wb = Workbook(); wb.remove(wb.active)
ws_c = wb.create_sheet("Clean Data")
dump_df(ws_c, clean, fmts={c: "yyyy-mm-dd" for c in ["timestamp_start", "event_date", "week_start", "month_start"]}, header_fill=TEAL)
ws_r = wb.create_sheet("Raw Data"); dump_df(ws_r, raw, header_fill=GREY)
ws_d = wb.create_sheet("Data Dictionary")
for r in dd.itertuples(index=False):
    ws_d.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in r])
for k, w in zip("ABCD", [34, 22, 70, 40]):
    ws_d.column_dimensions[k].width = w
for row in ws_d.iter_rows():
    for c in row:
        c.alignment = Alignment(wrap_text=True, vertical="top"); c.font = F(9)
ws_d["A1"].font = F(14, True, NAVY)
ctx = CS.Ctx(wb, cl, raw, clean, len(raw))
for c in clean.columns:
    add_name(wb, "cd_" + c, ctx.cr(c))
N = lambda c: "cd_" + c

# lists
wl = wb.create_sheet("Lists")
cols = {"A": ["All Plants"] + PLANTS, "B": ["All Lines"] + LINES, "C": ["All Shifts"] + SHIFTS, "D": LINES, "E": SHIFTS,
        "F": ["Last 4 weeks", "Last 13 weeks", "Last 26 weeks", "Last 52 weeks", "All data"], "G": [4, 13, 26, 52, 999],
        "H": ["All", "Planned", "Unplanned"], "I": REASONS}
for k, vals in cols.items():
    wl[f"{k}1"] = k
    for i, v in enumerate(vals):
        wl[f"{k}{i + 2}"] = v
wk_desc = [w for w in reversed(weeks_all) if w <= END][:104]
wl["J1"] = "weeks"
for i, w in enumerate(wk_desc):
    wl.cell(i + 2, 10, w.to_pydatetime()).number_format = "yyyy-mm-dd"
wl.sheet_state = "hidden"
LS = {k: f"=Lists!${k}$2:${k}${1 + len(v)}" for k, v in cols.items()}
LS["J"] = f"=Lists!$J$2:$J${1 + len(wk_desc)}"
PERIOD_LOOK = "Lists!$F$2:$G$6"

names = ["Cover", "01 Executive Summary", "02 OEE Command Center", "03 Root Cause Pareto", "04 Downtime Fingerprint",
         "05 Shift Variance Detector", "06 Drill Down Navigator", "07 Weekly OEE Scorecard", "08 Validation & Assumptions",
         "M1 Downtime Taxonomy", "M2 OEE Engine"]
S = {n: wb.create_sheet(n) for n in names}


def crit_block(ws, row, col, plant, line, shift, period, extra_label=None):
    """Write hidden criteria helpers; return dict of absolute refs."""
    sh = f"'{ws.title}'"
    items = [("crit plant", f'=IF({plant}="All Plants","*",{plant})'),
             ("crit line", f'=IF({line}="All Lines","*",{line})'),
             ("crit shift", f'=IF({shift}="All Shifts","*",{shift})'),
             ("period end", f"=MAX({N('event_date')})"),
             ("period start", f'=IF({period}="All data",DATE(2000,1,1),{L(col + 1)}{row + 3}-7*VLOOKUP({period},{PERIOD_LOOK},2,FALSE)+1)')]
    ref = {}
    for i, (a, f_) in enumerate(items):
        ws.cell(row + i, col, a).font = F(8, False, GREY)
        c = ws.cell(row + i, col + 1, f_); c.font = F(8, False, GREY)
        if "period" in a:
            c.number_format = "yyyy-mm-dd"
        ref[a.split()[1]] = f"{sh}!${L(col + 1)}${row + i}"
    return ref


def crit(ref, use=("plant", "line", "shift", "period")):
    parts = []
    if "plant" in use:
        parts += [N("plant_site"), ref["plant"]]
    if "line" in use:
        parts += [N("production_line"), ref["line"]]
    if "shift" in use:
        parts += [N("shift"), ref["shift"]]
    if "period" in use:
        parts += [N("event_date"), f'">="&{ref["start"]}', N("event_date"), f'"<="&{ref["end"]}']
    return ",".join(parts)


def S_(col, cr):
    return f"SUMIFS({N(col)},{cr})"


# ================================================================= M1 Taxonomy
ws = S["M1 Downtime Taxonomy"]
setup(ws, "M1  Governed Downtime Taxonomy", "Approach step 1. One controlled list of downtime reasons, each mapped to a planned or unplanned category, a Six Big Losses class and an OEE pillar.", cols=12, width_last=13)
ws.column_dimensions["C"].width = 22; ws.column_dimensions["E"].width = 22
r = section(ws, 5, 2, "Governed taxonomy (single source of truth)", 7)
header_row(ws, r, 2, ["Code", "Reason", "Category", "Six Big Losses class", "OEE pillar", "Events (live)", "Downtime minutes (live)"])
T_TOP = r + 1
for i, t in enumerate(TAX.itertuples()):
    rr = T_TOP + i
    vals = [t.code, t.reason, t.category, t.big_loss, t.pillar, f"=COUNTIF({N('reason')},C{rr})", f"=SUMIF({N('reason')},C{rr},{N('downtime_minutes')})"]
    for j, v in enumerate(vals):
        body_cell(ws.cell(rr, 2 + j, v), "#,##0" if j >= 5 else None, i % 2 == 1)
T_END = T_TOP + len(TAX) - 1
for nm, col in (("tax_code", "B"), ("tax_reason", "C"), ("tax_category", "D"), ("tax_bigloss", "E"), ("tax_pillar", "F")):
    add_name(wb, nm, f"'M1 Downtime Taxonomy'!${col}${T_TOP}:${col}${T_END}")
tr = T_END + 1
for j, v in enumerate(["Total", "", "", "", "", f"=SUM(G{T_TOP}:G{T_END})", f"=SUM(H{T_TOP}:H{T_END})"]):
    c = ws.cell(tr, 2 + j, v); c.font = F(9, True, NAVY); c.fill = fill(TEAL_L); c.border = BORDER; c.number_format = "#,##0"
ws.cell(tr + 1, 2, "Records without a governed code (live)").font = F(9)
c = ws.cell(tr + 1, 7, f'=COUNTBLANK({N("governed_code")})'); calc_style(c, "#,##0", True, NAVY)
ws.cell(tr + 1, 8, f'=IF(G{tr + 1}=0,"PASS","FAIL")'); CS.PASS_RULES(ws, f"H{tr + 1}")
TAXCOV = f"'M1 Downtime Taxonomy'!$G${tr + 1}"
block_note(ws, 6, 10, 14, 13, "Mapping rationale. Equipment failures (mechanical, electrical, tooling) are unplanned availability losses. Changeover is a planned setup loss. "
           "Material shortage stops a running line and is classed as idling. Quality hold is a process defect loss. Planned maintenance and operator breaks are planned stops. "
           "This follows the Six Big Losses structure; the mapping is a governed decision and can be edited here, and every sheet updates through the named ranges tax_*.")
r = tr + 3
r = section(ws, r, 2, "Proof that the stored codes and categories cannot be trusted (live)", 11)
note(ws, r, 2, "Each row is a governed reason (from the description); each column is the reason code stored in the raw file. If codes were reliable the counts would sit on the diagonal. They are spread evenly, so the stored code carries no information and is replaced.", span=11)
r += 1
header_row(ws, r, 2, ["Governed reason"] + [t for t in TAX.code] + ["Agreement"])
for i, rs in enumerate(REASONS):
    rr = r + 1 + i
    body_cell(ws.cell(rr, 2, rs), None, i % 2 == 1)
    for j, code in enumerate(TAX.code):
        body_cell(ws.cell(rr, 3 + j, f'=COUNTIFS({N("reason")},$B{rr},{N("downtime_reason_code")},"{code}")'), "#,##0", i % 2 == 1)
    body_cell(ws.cell(rr, 11, f"=INDEX(C{rr}:J{rr},{i + 1})/SUM(C{rr}:J{rr})"), "0.0%", i % 2 == 1)
ws.conditional_formatting.add(f"C{r + 1}:J{r + 8}", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="7FB3AE"))
r += 10
header_row(ws, r, 2, ["Governed category", "Stored: Planned", "Stored: Unplanned", "Agreement"])
for i, cat in enumerate(["Planned", "Unplanned"]):
    rr = r + 1 + i
    body_cell(ws.cell(rr, 2, cat))
    body_cell(ws.cell(rr, 3, f'=COUNTIFS({N("governed_category")},B{rr},{N("downtime_category")},"Planned")'), "#,##0")
    body_cell(ws.cell(rr, 4, f'=COUNTIFS({N("governed_category")},B{rr},{N("downtime_category")},"Unplanned")'), "#,##0")
    body_cell(ws.cell(rr, 5, f"={'C' if i == 0 else 'D'}{rr}/SUM(C{rr}:D{rr})"), "0.0%")
r += 4
r = section(ws, r, 2, "Raw description variants mapped to the governed reason", 6)
tm = cl.text_map[cl.text_map.Column == "downtime_reason_desc"].sort_values(["StandardValue", "Records"], ascending=[True, False])
write_table(ws, r, 2, tm[["RawValue", "StandardValue", "Records", "IssueType"]].rename(columns={"RawValue": "Raw description", "StandardValue": "Governed reason", "IssueType": "Issue"}), fmts={"Records": "#,##0"})

# ================================================================= M2 OEE engine
ws = S["M2 OEE Engine"]
setup(ws, "M2  OEE Calculation Engine (time weighted)", "Availability, performance, quality and OEE for every line and shift, computed from time layers so that OEE = A x P x Q holds exactly at every level.", cols=16, width_last=10)
ws.column_dimensions["B"].width = 14
block_note(ws, 5, 2, 8, 17, "Method. For each record: operating time = planned time x A; net run time = operating time x P; fully productive time (FPT) = net run time x Q. "
           "For any group of records: A = sum(operating) / sum(planned), P = sum(net run) / sum(operating), Q = sum(FPT) / sum(net run), OEE = sum(FPT) / sum(planned). "
           "Averaging record level percentages would weight a 300 minute run the same as a 960 minute run and would break the identity OEE = A x P x Q; time weighting keeps it exact.")
r = 10
r = section(ws, r, 2, "Line by shift OEE (all data)", 15)
header_row(ws, r, 2, ["Line"] + [f"{s} OEE" for s in SHIFTS] + ["Line OEE", "Availability", "Performance", "Quality", "Planned hours", "Events", "Reported OEE (legacy)", "Gap (points)"])
E_TOP = r + 1
for i, ln in enumerate(LINES):
    rr = E_TOP + i
    body_cell(ws.cell(rr, 2, ln), None, i % 2 == 1)
    for j, sh in enumerate(SHIFTS):
        f_ = f'=SUMIFS({N("fpt_min")},{N("production_line")},$B{rr},{N("shift")},"{sh}")/SUMIFS({N("planned_production_time_min")},{N("production_line")},$B{rr},{N("shift")},"{sh}")'
        body_cell(ws.cell(rr, 3 + j, f_), "0.0%", i % 2 == 1)
    L_ = f"{N('production_line')},$B{rr}"
    vals = [f"=SUMIFS({N('fpt_min')},{L_})/SUMIFS({N('planned_production_time_min')},{L_})",
            f"=SUMIFS({N('avail_time_min')},{L_})/SUMIFS({N('planned_production_time_min')},{L_})",
            f"=SUMIFS({N('perf_time_min')},{L_})/SUMIFS({N('avail_time_min')},{L_})",
            f"=SUMIFS({N('fpt_min')},{L_})/SUMIFS({N('perf_time_min')},{L_})",
            f"=SUMIFS({N('planned_production_time_min')},{L_})/60", f"=COUNTIFS({L_})",
            f"=AVERAGEIFS({N('oee_pct')},{L_})/100", f"=(L{rr}-F{rr})*100"]
    fm = ["0.0%", "0.0%", "0.0%", "0.0%", "#,##0", "#,##0", "0.0%", "+0.0;-0.0"]
    for j, v in enumerate(vals):
        body_cell(ws.cell(rr, 6 + j, v), fm[j], i % 2 == 1)
E_END = E_TOP + len(LINES) - 1
tr = E_END + 1
tots = ["Plant network"] + [f'=SUMIFS({N("fpt_min")},{N("shift")},"{sh}")/SUMIFS({N("planned_production_time_min")},{N("shift")},"{sh}")' for sh in SHIFTS] + \
       [f"=SUM({N('fpt_min')})/SUM({N('planned_production_time_min')})", f"=SUM({N('avail_time_min')})/SUM({N('planned_production_time_min')})",
        f"=SUM({N('perf_time_min')})/SUM({N('avail_time_min')})", f"=SUM({N('fpt_min')})/SUM({N('perf_time_min')})",
        f"=SUM({N('planned_production_time_min')})/60", f"=COUNTA({N('record_id')})", f"=AVERAGE({N('oee_pct')})/100", f"=(L{tr}-F{tr})*100"]
fm = [None, "0.0%", "0.0%", "0.0%", "0.0%", "0.0%", "0.0%", "0.0%", "#,##0", "#,##0", "0.0%", "+0.0;-0.0"]
for j, v in enumerate(tots):
    c = ws.cell(tr, 2 + j, v); c.font = F(9, True, NAVY); c.fill = fill(TEAL_L); c.border = BORDER
    if fm[j]:
        c.number_format = fm[j]
ENG = {"oee": f"'M2 OEE Engine'!$F${tr}", "a": f"'M2 OEE Engine'!$G${tr}", "p": f"'M2 OEE Engine'!$H${tr}", "q": f"'M2 OEE Engine'!$I${tr}",
       "rep": f"'M2 OEE Engine'!$L${tr}", "gap": f"'M2 OEE Engine'!$M${tr}", "hours": f"'M2 OEE Engine'!$J${tr}", "row": tr}
ws.conditional_formatting.add(f"C{E_TOP}:F{E_END}", ColorScaleRule(start_type="min", start_color="F4A6A6", mid_type="percentile", mid_value=50, mid_color="FFF2CC", end_type="max", end_color="9BD3AE"))
ws.cell(tr + 1, 2, "Identity check: OEE equals A x P x Q at network level").font = F(9)
c = ws.cell(tr + 1, 6, f"=ABS(F{tr}-G{tr}*H{tr}*I{tr})"); calc_style(c, "0.000000")
ws.cell(tr + 1, 7, f'=IF(F{tr + 1}<0.000001,"PASS","FAIL")'); CS.PASS_RULES(ws, f"G{tr + 1}")
IDENT = f"'M2 OEE Engine'!$F${tr + 1}"
r = tr + 3
r = section(ws, r, 2, "Plant by line OEE (all data)", 15)
header_row(ws, r, 2, ["Plant"] + LINES + ["Plant OEE"])
P_TOP = r + 1
for i, p in enumerate(PLANTS):
    rr = P_TOP + i
    body_cell(ws.cell(rr, 2, p), None, i % 2 == 1)
    for j, ln in enumerate(LINES):
        f_ = f'=IFERROR(SUMIFS({N("fpt_min")},{N("plant_site")},$B{rr},{N("production_line")},"{ln}")/SUMIFS({N("planned_production_time_min")},{N("plant_site")},$B{rr},{N("production_line")},"{ln}"),"")'
        body_cell(ws.cell(rr, 3 + j, f_), "0.0%", i % 2 == 1)
    body_cell(ws.cell(rr, 15, f'=SUMIFS({N("fpt_min")},{N("plant_site")},$B{rr})/SUMIFS({N("planned_production_time_min")},{N("plant_site")},$B{rr})'), "0.0%", i % 2 == 1)
ws.conditional_formatting.add(f"C{P_TOP}:O{P_TOP + 3}", ColorScaleRule(start_type="min", start_color="F4A6A6", mid_type="percentile", mid_value=50, mid_color="FFF2CC", end_type="max", end_color="9BD3AE"))
ws.column_dimensions["B"].width = 20
ws.freeze_panes = "C12"

# ================================================================= 02 Command Center
ws = S["02 OEE Command Center"]
setup(ws, "OEE Command Center", "Real time plant view. Filter by plant, line, shift and period; every tile, chart and the heat map respond. OEE is the governed, time weighted A x P x Q.", cols=18, width_last=9.5)
ws.column_dimensions["B"].width = 12
selector(ws, "B6", LS["A"], "All Plants", "PLANT", "B5"); ws.merge_cells("B6:D6")
selector(ws, "F6", LS["B"], "All Lines", "LINE", "F5"); ws.merge_cells("F6:H6")
selector(ws, "J6", LS["C"], "All Shifts", "SHIFT", "J5"); ws.merge_cells("J6:L6")
selector(ws, "N6", LS["F"], "Last 13 weeks", "PERIOD", "N5"); ws.merge_cells("N6:P6")
ws.row_dimensions[6].height = 24
HC = 30  # helper col AD
ref = crit_block(ws, 5, HC, "$B$6", "$F$6", "$J$6", "$N$6")
CR = crit(ref)
hc = lambda i: f"${L(HC + 1)}${12 + i}"
helpers = [("planned", S_("planned_production_time_min", CR)), ("operating", S_("avail_time_min", CR)), ("net run", S_("perf_time_min", CR)),
           ("fpt", S_("fpt_min", CR)), ("downtime min", S_("downtime_minutes", CR)), ("events", f"COUNTIFS({CR})"),
           ("reported oee", f"IFERROR(AVERAGEIFS({N('oee_pct')},{CR})/100,0)"), ("scrap cost", S_("scrap_cost_usd", CR))]
for i, (a, f_) in enumerate(helpers):
    ws.cell(12 + i, HC, a).font = F(8, False, GREY)
    c = ws.cell(12 + i, HC + 1, "=" + f_); c.font = F(8, False, GREY)
A_ = f"{hc(1)}/{hc(0)}"; P_ = f"{hc(2)}/{hc(1)}"; Q_ = f"{hc(3)}/{hc(2)}"; O_ = f"{hc(3)}/{hc(0)}"
cards = [("OEE (GOVERNED)", f"=IFERROR({O_},0)", "0.0%", "A x P x Q, time weighted", TEAL),
         ("AVAILABILITY", f"=IFERROR({A_},0)", "0.0%", "Operating / planned time", NAVY),
         ("PERFORMANCE", f"=IFERROR({P_},0)", "0.0%", "Net run / operating time", NAVY),
         ("QUALITY", f"=IFERROR({Q_},0)", "0.0%", "Good / total output time", NAVY),
         ("DOWNTIME HOURS", f"={hc(4)}/60", "#,##0", f'Events in selection', AMBER),
         ("REPORTED OEE GAP", f"=({hc(6)}-IFERROR({O_},0))*100", '+0.0" pts";-0.0" pts"', "Legacy oee_pct minus governed", RED)]
for i, (a, b, fm, s_, col) in enumerate(cards):
    kpi(ws, 8, 2 + i * 3, a, b, fm, s_, 3, col)
ws.cell(10, 14, f'=TEXT({hc(5)},"#,##0")&" events in selection"').font = F(8, False, GREY, True)
# trend helper: last 26 weeks, filters plant/line/shift
TR0 = 22
ws.cell(TR0 - 1, HC, "Month"); ws.cell(TR0 - 1, HC + 1, "OEE"); ws.cell(TR0 - 1, HC + 2, "Availability"); ws.cell(TR0 - 1, HC + 3, "Performance"); ws.cell(TR0 - 1, HC + 4, "Quality")
last12 = [pd.Timestamp(m) for m in sorted(clean.month_start.dropna().unique())][-12:]
CR3 = crit(ref, ("plant", "line", "shift"))
for i, m in enumerate(last12):
    rr = TR0 + i
    mk = f"{N('month_start')},{L(HC)}{rr}"
    ws.cell(rr, HC, m.to_pydatetime()).number_format = "mmm yy"
    ws.cell(rr, HC + 1, f"=IFERROR(SUMIFS({N('fpt_min')},{CR3},{mk})/SUMIFS({N('planned_production_time_min')},{CR3},{mk}),0)")
    ws.cell(rr, HC + 2, f"=IFERROR(SUMIFS({N('avail_time_min')},{CR3},{mk})/SUMIFS({N('planned_production_time_min')},{CR3},{mk}),0)")
    ws.cell(rr, HC + 3, f"=IFERROR(SUMIFS({N('perf_time_min')},{CR3},{mk})/SUMIFS({N('avail_time_min')},{CR3},{mk}),0)")
    ws.cell(rr, HC + 4, f"=IFERROR(SUMIFS({N('fpt_min')},{CR3},{mk})/SUMIFS({N('perf_time_min')},{CR3},{mk}),0)")
# waterfall helper
WF0 = 52
ws.cell(WF0 - 1, HC, "Stage"); ws.cell(WF0 - 1, HC + 1, "Base"); ws.cell(WF0 - 1, HC + 2, "Share of planned time")
wf = [("Planned time", "0", "1"), ("Availability loss", f"IFERROR({A_},0)", f"1-IFERROR({A_},0)"),
      ("Performance loss", f"IFERROR({hc(2)}/{hc(0)},0)", f"IFERROR(({hc(1)}-{hc(2)})/{hc(0)},0)"),
      ("Quality loss", f"IFERROR({O_},0)", f"IFERROR(({hc(2)}-{hc(3)})/{hc(0)},0)"), ("Fully productive (OEE)", "0", f"IFERROR({O_},0)")]
for i, (a, b, c_) in enumerate(wf):
    ws.cell(WF0 + i, HC, a); ws.cell(WF0 + i, HC + 1, "=" + b); ws.cell(WF0 + i, HC + 2, "=" + c_)
# reasons helper sorted
RS0 = 60
ws.cell(RS0 - 1, HC, "Reason"); ws.cell(RS0 - 1, HC + 1, "Minutes (tie broken)"); ws.cell(RS0 - 1, HC + 3, "Sorted reason"); ws.cell(RS0 - 1, HC + 4, "Hours")
for i, rs in enumerate(REASONS):
    rr = RS0 + i
    ws.cell(rr, HC, rs)
    ws.cell(rr, HC + 1, f'=SUMIFS({N("downtime_minutes")},{CR},{N("reason")},{L(HC)}{rr})+{i + 1}/1000000')
    ws.cell(rr, HC + 3, f"=INDEX(${L(HC)}${RS0}:${L(HC)}${RS0 + 7},MATCH(LARGE(${L(HC + 1)}${RS0}:${L(HC + 1)}${RS0 + 7},{i + 1}),${L(HC + 1)}${RS0}:${L(HC + 1)}${RS0 + 7},0))")
    ws.cell(rr, HC + 4, f"=LARGE(${L(HC + 1)}${RS0}:${L(HC + 1)}${RS0 + 7},{i + 1})/60")
for rr in range(5, 70):
    for k in range(HC, HC + 5):
        ws.cell(rr, k).font = F(8, False, GREY)
# heat map line x shift, filters plant + period
section(ws, 12, 2, "Line by shift OEE heat map (plant and period filters)", 6)
header_row(ws, 13, 2, ["Line", "Day", "Swing", "Night", "Line OEE", "vs network"], height=22)
CRpp = crit(ref, ("plant", "period"))
for i, ln in enumerate(LINES):
    rr = 14 + i
    body_cell(ws.cell(rr, 2, ln)); ws.cell(rr, 2).font = F(9, True)
    for j, sh in enumerate(SHIFTS):
        f_ = f'=IFERROR(SUMIFS({N("fpt_min")},{CRpp},{N("production_line")},$B{rr},{N("shift")},"{sh}")/SUMIFS({N("planned_production_time_min")},{CRpp},{N("production_line")},$B{rr},{N("shift")},"{sh}"),"")'
        body_cell(ws.cell(rr, 3 + j, f_), "0.0%", align="center")
    body_cell(ws.cell(rr, 6, f'=IFERROR(SUMIFS({N("fpt_min")},{CRpp},{N("production_line")},$B{rr})/SUMIFS({N("planned_production_time_min")},{CRpp},{N("production_line")},$B{rr}),"")'), "0.0%", align="center")
    body_cell(ws.cell(rr, 7, f'=IFERROR((F{rr}-SUMIFS({N("fpt_min")},{CRpp})/SUMIFS({N("planned_production_time_min")},{CRpp}))*100,"")'), '+0.0;-0.0', align="center")
    ws.row_dimensions[rr].height = 16
ws.conditional_formatting.add("C14:F25", ColorScaleRule(start_type="min", start_color="F4A6A6", mid_type="percentile", mid_value=50, mid_color="FFF2CC", end_type="max", end_color="9BD3AE"))
ws.conditional_formatting.add("G14:G25", CellIsRule(operator="lessThan", formula=["0"], font=F(9, True, RED)))
ws.conditional_formatting.add("G14:G25", CellIsRule(operator="greaterThan", formula=["0"], font=F(9, True, GREEN)))
ws.conditional_formatting.add("B14:B25", FormulaRule(formula=["$B14=$F$6"], fill=fill(TEAL_L), font=F(9, True, TEAL)))
c1 = LineChart(); ctitle(c1, "Monthly OEE and components, last 12 months")
c1.add_data(Reference(ws, min_col=HC + 1, max_col=HC + 4, min_row=TR0 - 1, max_row=TR0 + 11), titles_from_data=True)
c1.set_categories(Reference(ws, min_col=HC, min_row=TR0, max_row=TR0 + 11))
for s_, colr, wdt in zip(c1.series, [NAVY, "7FB3AE", AMBER, "9AA7B8"], [32000, 15000, 15000, 15000]):
    s_.graphicalProperties.line.solidFill = colr; s_.graphicalProperties.line.width = wdt; s_.smooth = False
c1.y_axis.number_format = "0%"; c1.y_axis.scaling.min = 0; c1.y_axis.scaling.max = 1.0; c1.x_axis.number_format = "mmm yy"
c1.height = 7; c1.width = 17; c1.legend.position = "b"; c1.visible_cells_only = False
ws.add_chart(c1, "I12")
c2 = BarChart(); c2.type = "col"; c2.grouping = "stacked"; c2.overlap = 100; ctitle(c2, "Where planned time is lost (share of planned time)")
c2.add_data(Reference(ws, min_col=HC + 1, max_col=HC + 2, min_row=WF0 - 1, max_row=WF0 + 4), titles_from_data=True)
c2.set_categories(Reference(ws, min_col=HC, min_row=WF0, max_row=WF0 + 4))
c2.series[0].graphicalProperties.noFill = True; c2.series[0].graphicalProperties.line.noFill = True
c2.series[1].graphicalProperties.solidFill = TEAL
from openpyxl.chart.label import DataLabelList
c2.series[1].dLbls = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False, showLeaderLines=False, numFmt="0%")
c2.series[1].dLbls.position = "inEnd"
c2.y_axis.number_format = "0%"; c2.y_axis.scaling.max = 1; c2.y_axis.scaling.min = 0; c2.legend = None; c2.gapWidth = 40
c2.height = 7; c2.width = 12; c2.visible_cells_only = False
ws.add_chart(c2, "B28")
c3 = BarChart(); c3.type = "bar"; ctitle(c3, "Downtime hours by governed reason")
c3.add_data(Reference(ws, min_col=HC + 4, min_row=RS0 - 1, max_row=RS0 + 7), titles_from_data=True)
c3.set_categories(Reference(ws, min_col=HC + 3, min_row=RS0, max_row=RS0 + 7))
c3.series[0].graphicalProperties.solidFill = NAVY; c3.legend = None; c3.gapWidth = 40
c3.x_axis.scaling.orientation = "maxMin"
c3.height = 7; c3.width = 15; c3.visible_cells_only = False
ws.add_chart(c3, "J28")
note(ws, 43, 2, "Reading the command center. Tiles show the governed OEE and its three components for the current selection. The waterfall shows how planned time is lost to availability, performance and quality until only fully productive time remains; that final bar equals OEE. "
     "The heat map compares every line and shift under the plant and period filters, with the selected line highlighted. The reported OEE gap tile shows how far the legacy oee_pct field sits from the governed calculation.", span=18, height=44)
for k in range(HC, HC + 5):
    ws.column_dimensions[L(k)].hidden = True
print_fit(ws, "A1:T45")

# ================================================================= 03 Pareto
ws = S["03 Root Cause Pareto"]
setup(ws, "03  Root Cause Pareto Analysis", "Deliverable 2. Ranks governed downtime reasons and line level hotspots by lost minutes, and tests how many causes explain 80 percent of downtime.", cols=16, width_last=10)
ws.column_dimensions["B"].width = 22
selector(ws, "B6", LS["A"], "All Plants", "PLANT", "B5")
selector(ws, "D6", LS["B"], "All Lines", "LINE", "D5"); ws.merge_cells("D6:E6")
selector(ws, "G6", LS["C"], "All Shifts", "SHIFT", "G5"); ws.merge_cells("G6:H6")
selector(ws, "J6", LS["F"], "All data", "PERIOD", "J5"); ws.merge_cells("J6:K6")
selector(ws, "M6", LS["H"], "All", "CATEGORY", "M5"); ws.merge_cells("M6:N6")
ref = crit_block(ws, 5, 30, "$B$6", "$D$6", "$G$6", "$J$6")
ws.cell(4, 30, "crit category").font = F(8, False, GREY)
ws.cell(4, 31, '=IF($M$6="All","*",$M$6)').font = F(8, False, GREY)
CR = crit(ref) + f",{N('governed_category')},$AE$4"
r = section(ws, 8, 2, "Pareto of governed downtime reasons", 8)
header_row(ws, r, 2, ["Reason (ranked)", "Downtime hours", "Events", "Share", "Cumulative share", "Class", "Rank key"])
PT = r + 1
for i, rs in enumerate(REASONS):
    rr = PT + i
    ws.cell(rr, 30, rs).font = F(8, False, GREY)
    c = ws.cell(rr, 31, f'=SUMIFS({N("downtime_minutes")},{CR},{N("reason")},AD{rr})/60+{i + 1}/1000000'); c.font = F(8, False, GREY)
for i in range(8):
    rr = PT + i
    vals = [f"=INDEX($AD${PT}:$AD${PT + 7},MATCH(H{rr},$AE${PT}:$AE${PT + 7},0))", f"=ROUND(H{rr},2)",
            f'=COUNTIFS({CR},{N("reason")},B{rr})', f"=C{rr}/SUM($C${PT}:$C${PT + 7})",
            f"=SUM($E${PT}:E{rr})", f'=IF(OR(F{rr}<=0.8,{i}=0,F{rr - 1 if i else rr}<0.8),"Needed for 80%","Beyond 80%")',
            f"=LARGE($AE${PT}:$AE${PT + 7},{i + 1})"]
    fm = [None, "#,##0.0", "#,##0", "0.0%", "0.0%", None, "0.000"]
    for j, v in enumerate(vals):
        body_cell(ws.cell(rr, 2 + j, v), fm[j], i % 2 == 1)
    ws.cell(rr, 8).font = F(8, False, GREY)
ws.conditional_formatting.add(f"G{PT}:G{PT + 7}", CellIsRule(operator="equal", formula=['"Needed for 80%"'], fill=fill(AMBER_L), font=F(9, True, AMBER)))
ws.conditional_formatting.add(f"E{PT}:E{PT + 7}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=0.3, color="5EA8A0"))
r = PT + 9
ws.cell(r, 2, "Causes needed to explain 80 percent").font = F(10, True)
c = ws.cell(r, 5, f'=COUNTIF(F{PT}:F{PT + 7},"<0.8")+1'); calc_style(c, '0" of 8"', True, NAVY)
ws.cell(r + 1, 2, "Largest single cause share").font = F(10, True)
c = ws.cell(r + 1, 5, f"=E{PT}"); calc_style(c, "0.0%", True, NAVY)
ws.cell(r + 2, 2, "Concentration versus perfectly even spread").font = F(10, True)
c = ws.cell(r + 2, 5, f"=E{PT}/(1/8)"); calc_style(c, '0.00"x"', True, NAVY)
PAR = {"n80": f"'03 Root Cause Pareto'!$E${r}", "top": f"'03 Root Cause Pareto'!$E${r + 1}"}
ch = BarChart(); ch.type = "col"; ctitle(ch, "Downtime Pareto: hours and cumulative share")
ch.add_data(Reference(ws, min_col=3, min_row=PT - 1, max_row=PT + 7), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=PT, max_row=PT + 7))
ch.series[0].graphicalProperties.solidFill = NAVY; ch.gapWidth = 40; ch.y_axis.title = "Hours"; ch.y_axis.scaling.min = 0
lc = LineChart(); lc.add_data(Reference(ws, min_col=6, min_row=PT - 1, max_row=PT + 7), titles_from_data=True)
lc.y_axis.axId = 200; lc.y_axis.number_format = "0%"; lc.y_axis.scaling.min = 0; lc.y_axis.scaling.max = 1; lc.y_axis.crosses = "max"
lc.series[0].graphicalProperties.line.solidFill = AMBER; lc.series[0].graphicalProperties.line.width = 28000
lc.y_axis.delete = False; lc.y_axis.title = "Cumulative"; ch.y_axis.crosses = "min"
ch += lc; ch.height = 8.5; ch.width = 15.5; ch.legend.position = "b"
ws.add_chart(ch, "J8")
r += 4
r = section(ws, r, 2, "Six Big Losses view", 6)
header_row(ws, r, 2, ["Loss class", "Downtime hours", "Share"])
for i, lo in enumerate(LOSSES):
    rr = r + 1 + i
    body_cell(ws.cell(rr, 2, lo), None, i % 2 == 1)
    body_cell(ws.cell(rr, 3, f'=SUMIFS({N("downtime_minutes")},{CR},{N("six_big_loss")},B{rr})/60'), "#,##0.0", i % 2 == 1)
    body_cell(ws.cell(rr, 4, f"=C{rr}/SUM($C${r + 1}:$C${r + 5})"), "0.0%", i % 2 == 1)
ws.conditional_formatting.add(f"D{r + 1}:D{r + 5}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=0.5, color="E29578"))
r += 7
r = section(ws, r, 2, "Line level hotspots: top 15 line and reason combinations", 8)
note(ws, r, 2, "Second test for a vital few: split every reason by line (96 combinations, scored in hidden columns AH to AP) and rank the worst. If the top five combinations hold only a small share, concentration does not exist at line level either and the loss is systemic.", span=12)
r += 1
header_row(ws, r, 2, ["Rank", "Line", "Reason", "Downtime hours", "Share of selection", "Cumulative"])
G0 = 34  # column AH grid (hidden)
ws.cell(PT - 1, G0, "Grid").font = F(8, True, GREY)
for j, rs in enumerate(REASONS):
    ws.cell(PT - 1, G0 + 1 + j, rs).font = F(8, True, GREY)
for i, ln in enumerate(LINES):
    rr = PT + i
    ws.cell(rr, G0, ln).font = F(8, False, GREY)
    for j in range(8):
        c = ws.cell(rr, G0 + 1 + j, f'=SUMIFS({N("downtime_minutes")},{crit(ref, ("plant", "shift", "period"))},{N("governed_category")},$AE$4,{N("production_line")},${L(G0)}{rr},{N("reason")},{L(G0 + 1 + j)}${PT - 1})/60+{(i * 8 + j + 1)}/10000000')
        c.font = F(8, False, GREY); c.number_format = "0.0"
GR = f"${L(G0 + 1)}${PT}:${L(G0 + 8)}${PT + 11}"
for k in range(15):
    rr = r + 1 + k
    body_cell(ws.cell(rr, 2, k + 1), "0", k % 2 == 1)
    val = f"LARGE({GR},{k + 1})"
    body_cell(ws.cell(rr, 3, f"=INDEX(${L(G0)}${PT}:${L(G0)}${PT + 11},SUMPRODUCT(({GR}={val})*(ROW({GR})-{PT - 1})))"), None, k % 2 == 1)
    body_cell(ws.cell(rr, 4, f"=INDEX(${L(G0 + 1)}${PT - 1}:${L(G0 + 8)}${PT - 1},SUMPRODUCT(({GR}={val})*(COLUMN({GR})-{G0})))"), None, k % 2 == 1)
    body_cell(ws.cell(rr, 5, f"={val}"), "#,##0.0", k % 2 == 1)
    body_cell(ws.cell(rr, 6, f"=E{rr}/SUM({GR})"), "0.0%", k % 2 == 1)
    body_cell(ws.cell(rr, 7, f"=SUM($F${r + 1}:F{rr})"), "0.0%", k % 2 == 1)
ws.conditional_formatting.add(f"E{r + 1}:E{r + 15}", DataBarRule(start_type="min", end_type="max", color="E29578"))
ws.cell(r + 16, 2, "Top 5 hotspots: share of downtime").font = F(10, True)
c = ws.cell(r + 16, 6, f"=G{r + 5}"); calc_style(c, "0.0%", True, NAVY)
PAR["hot5"] = f"'03 Root Cause Pareto'!$F${r + 16}"
for k in range(30, 32):
    ws.column_dimensions[L(k)].hidden = True
for k in range(G0, G0 + 9):
    ws.column_dimensions[L(k)].hidden = True
ws.column_dimensions["G"].width = 15; ws.column_dimensions["D"].width = 19; ws.column_dimensions["H"].hidden = True
print_fit(ws, "A1:Q60", landscape=False)

# ================================================================= 04 Fingerprint
ws = S["04 Downtime Fingerprint"]
setup(ws, "04  Downtime Fingerprint Radar", "Deliverable 3. Each line's mix of downtime reasons compared with the network, with a chi square test of whether line fingerprints genuinely differ.", cols=14, width_last=10)
ws.column_dimensions["B"].width = 20
selector(ws, "B6", LS["D"], LINES[0], "LINE TO PROFILE", "B5")
selector(ws, "D6", LS["F"], "All data", "PERIOD", "D5"); ws.merge_cells("D6:E6")
ref = crit_block(ws, 5, 30, '"All Plants"', '"All Lines"', '"All Shifts"', "$D$6")
CRp = crit(ref, ("period",))
r = section(ws, 8, 2, "Share of downtime minutes by reason", 5)
header_row(ws, r, 2, ["Reason", "Selected line", "Network average", "Difference (points)", "Index (line / network)"])
FP = r + 1
for i, rs in enumerate(REASONS):
    rr = FP + i
    vals = [rs, f'=SUMIFS({N("downtime_minutes")},{CRp},{N("production_line")},$B$6,{N("reason")},B{rr})/SUMIFS({N("downtime_minutes")},{CRp},{N("production_line")},$B$6)',
            f'=SUMIFS({N("downtime_minutes")},{CRp},{N("reason")},B{rr})/SUMIFS({N("downtime_minutes")},{CRp})',
            f"=(C{rr}-D{rr})*100", f"=C{rr}/D{rr}"]
    fm = [None, "0.0%", "0.0%", "+0.0;-0.0", "0.00"]
    for j, v in enumerate(vals):
        body_cell(ws.cell(rr, 2 + j, v), fm[j], i % 2 == 1)
ws.conditional_formatting.add(f"E{FP}:E{FP + 7}", CellIsRule(operator="greaterThan", formula=["1.5"], font=F(9, True, RED)))
ws.conditional_formatting.add(f"F{FP}:F{FP + 7}", ColorScaleRule(start_type="num", start_value=0.7, start_color="9BD3AE", mid_type="num", mid_value=1, mid_color="FFFFFF", end_type="num", end_value=1.3, end_color="F4A6A6"))
rr = FP + 9
ws.cell(rr, 2, "Fingerprint distance from network").font = F(10, True)
c = ws.cell(rr, 5, f"=SUMPRODUCT(ABS(C{FP}:C{FP + 7}-D{FP}:D{FP + 7}))/2"); calc_style(c, "0.0%", True, NAVY)
ws.cell(rr + 1, 2, "Dominant reason on this line").font = F(10, True)
c = ws.cell(rr + 1, 5, f"=INDEX(B{FP}:B{FP + 7},MATCH(MAX(F{FP}:F{FP + 7}),F{FP}:F{FP + 7},0))"); calc_style(c, None, True, TEAL)
note(ws, rr + 2, 2, "Distance is the share of this line's downtime that would have to move between reasons to match the network mix (0 percent means identical).", span=5)
rc = RadarChart(); rc.type = "marker"; ctitle(rc, "Downtime fingerprint: selected line versus network")
rc.add_data(Reference(ws, min_col=3, max_col=4, min_row=FP - 1, max_row=FP + 7), titles_from_data=True)
rc.set_categories(Reference(ws, min_col=2, min_row=FP, max_row=FP + 7))
rc.series[0].graphicalProperties.line.solidFill = TEAL; rc.series[0].graphicalProperties.line.width = 28000
rc.series[1].graphicalProperties.line.solidFill = "9AA7B8"; rc.series[1].graphicalProperties.line.dashStyle = "dash"
rc.y_axis.number_format = "0%"; rc.y_axis.delete = False; rc.height = 9.5; rc.width = 14; rc.legend.position = "b"
ws.add_chart(rc, "H8")
r = FP + 14
r = section(ws, r, 2, "All line fingerprints (share of downtime minutes)", 12)
header_row(ws, r, 2, ["Line"] + [x.replace(" ", "\n", 1) for x in REASONS] + ["Distance", "Dominant"], height=36)
MT = r + 1
for i, ln in enumerate(LINES):
    rr = MT + i
    body_cell(ws.cell(rr, 2, ln), None, i % 2 == 1)
    for j, rs in enumerate(REASONS):
        body_cell(ws.cell(rr, 3 + j, f'=SUMIFS({N("downtime_minutes")},{CRp},{N("production_line")},$B{rr},{N("reason")},"{rs}")/SUMIFS({N("downtime_minutes")},{CRp},{N("production_line")},$B{rr})'), "0.0%", i % 2 == 1)
    body_cell(ws.cell(rr, 11, f"=SUMPRODUCT(ABS(C{rr}:J{rr}-$C${MT + 12}:$J${MT + 12}))/2"), "0.0%", i % 2 == 1)
    body_cell(ws.cell(rr, 12, f"=INDEX(tax_reason,MATCH(MAX(C{rr}:J{rr}),C{rr}:J{rr},0))"), None, i % 2 == 1)
rr = MT + 12
body_cell(ws.cell(rr, 2, "Network"), None); ws.cell(rr, 2).font = F(9, True, NAVY)
for j, rs in enumerate(REASONS):
    c = ws.cell(rr, 3 + j, f'=SUMIFS({N("downtime_minutes")},{CRp},{N("reason")},"{rs}")/SUMIFS({N("downtime_minutes")},{CRp})'); calc_style(c, "0.0%", True, NAVY); c.fill = fill(TEAL_L)
ws.conditional_formatting.add(f"C{MT}:J{MT + 11}", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="E29578"))
ws.conditional_formatting.add(f"B{MT}:B{MT + 11}", FormulaRule(formula=[f"$B{MT}=$B$6"], fill=fill(TEAL_L), font=F(9, True, TEAL)))
for k in range(3, 11):
    ws.column_dimensions[L(k)].width = 11
r = MT + 14
r = section(ws, r, 2, "Chi square test: do lines have different fingerprints? (event counts)", 12)
header_row(ws, r, 2, ["Line"] + [x.replace(" ", "\n", 1) for x in REASONS] + ["Row total"], height=36)
OB = r + 1
for i, ln in enumerate(LINES):
    rr = OB + i
    body_cell(ws.cell(rr, 2, ln), None, i % 2 == 1)
    for j, rs in enumerate(REASONS):
        body_cell(ws.cell(rr, 3 + j, f'=COUNTIFS({CRp},{N("production_line")},$B{rr},{N("reason")},"{rs}")'), "#,##0", i % 2 == 1)
    body_cell(ws.cell(rr, 11, f"=SUM(C{rr}:J{rr})"), "#,##0", i % 2 == 1)
rr = OB + 12
body_cell(ws.cell(rr, 2, "Column total"))
for j in range(9):
    c = ws.cell(rr, 3 + j, f"=SUM({L(3 + j)}{OB}:{L(3 + j)}{OB + 11})"); calc_style(c, "#,##0", True, NAVY)
EX = rr + 2
ws.cell(EX - 1, 2, "Expected counts under independence").font = F(9, True, GREY)
for i in range(12):
    for j in range(8):
        c = ws.cell(EX + i, 3 + j, f"=$K{OB + i}*{L(3 + j)}${OB + 12}/$K${OB + 12}"); c.number_format = "0.0"; c.font = F(8, False, GREY)
    ws.cell(EX + i, 2, f"=B{OB + i}").font = F(8, False, GREY)
sr = EX + 13
stats = [("Chi square statistic", f"=SUMPRODUCT((C{OB}:J{OB + 11}-C{EX}:J{EX + 11})^2/C{EX}:J{EX + 11})", "#,##0.00"),
         ("Degrees of freedom", "=(12-1)*(8-1)", "0"),
         ("p value", f"=CHIDIST(E{sr},E{sr + 1})", "0.0000"),
         ("Cramer's V (effect size)", f"=SQRT(E{sr}/(K{OB + 12}*(8-1)))", "0.000"),
         ("Conclusion", f'=IF(E{sr + 2}<0.05,"Line fingerprints differ significantly; target line specific causes","No significant difference between lines: downtime mix is systemic, so fix causes network wide")', None)]
for i, (a, b, fm) in enumerate(stats):
    ws.cell(sr + i, 2, a).font = F(10, True)
    c = ws.cell(sr + i, 5, b); calc_style(c, fm, True, NAVY)
FPS = {"p": f"'04 Downtime Fingerprint'!$E${sr + 2}", "v": f"'04 Downtime Fingerprint'!$E${sr + 3}", "concl": f"'04 Downtime Fingerprint'!$E${sr + 4}"}
note(ws, sr + 6, 2, "Why line level and not machine level. machine_id is unique on every one of the 15,250 records (C1), so no machine has more than one event and a machine fingerprint cannot be estimated. The line is the finest grain with repeated observations; the same sheet works unchanged for machines once the plant historian supplies a persistent asset identifier.", span=12, color=AMBER)
for k in (30, 31):
    ws.column_dimensions[L(k)].hidden = True
ws.column_dimensions["L"].width = 21
print_fit(ws, "A1:O75", landscape=False)

# ================================================================= 05 Shift variance detector
ws = S["05 Shift Variance Detector"]
setup(ws, "05  Shift Variance Anomaly Detector", "Deliverable 4. Weekly OEE for every line and shift tested against its own trailing 13 week baseline; weeks beyond the control limit are flagged.", cols=16, width_last=10)
ws.column_dimensions["B"].width = 16
selector(ws, "B6", LS["D"], LINES[0], "LINE", "B5")
selector(ws, "D6", LS["E"], "Night", "SHIFT", "D5"); ws.merge_cells("D6:E6")
selector(ws, "G6", LS["J"], wk_desc[1].to_pydatetime(), "WEEK TO SCAN", "G5"); ws.merge_cells("G6:H6"); ws["G6"].number_format = "dd mmm yyyy"
kv(ws, 5, 10, "Baseline weeks", 13, "0", True, span_label=2, comment="Catalogue brief: 13 week rolling baseline")
kv(ws, 6, 10, "Control limit (z)", 2.0, "0.0", True, span_label=2, comment="Two standard deviations: about 5 percent false alarm rate on a stable process")
kv(ws, 7, 10, "Minimum records per week", 2, "0", True, span_label=2, comment="Weeks with fewer records are not scored")
BW, ZL, MINR = "$L$5", "$L$6", "$L$7"
ws.cell(6, 14, '=$B$6&" | "&$D$6').font = F(8, False, WHITE)
KEY = "$N$6"
# scan matrix for all 36 line shifts, week selected
r = section(ws, 9, 2, "Network scan for the selected week: all 36 line and shift combinations", 12)
header_row(ws, r, 2, ["Line", "Shift", "Week OEE", "Baseline mean", "Baseline SD", "z score", "Records", "Status"])
SC = r + 1
combos = [(ln, sh) for ln in LINES for sh in SHIFTS]
HG = 20  # hidden grid from column T: 13 prior weeks OEE per combo
for i, (ln, sh) in enumerate(combos):
    rr = SC + i
    key = f'$B{rr}&" | "&$C{rr}'
    body_cell(ws.cell(rr, 2, ln), None, i % 2 == 1); body_cell(ws.cell(rr, 3, sh), None, i % 2 == 1)
    wk = f'{N("week_start")},$G$6'
    body_cell(ws.cell(rr, 4, f'=IFERROR(SUMIFS({N("fpt_min")},{N("line_shift")},{key},{wk})/SUMIFS({N("planned_production_time_min")},{N("line_shift")},{key},{wk}),"")'), "0.0%", i % 2 == 1)
    for k in range(13):
        wkk = f'{N("week_start")},$G$6-7*{k + 1}'
        c = ws.cell(rr, HG + k, f'=IF(COUNTIFS({N("line_shift")},{key},{wkk})<{MINR},"",SUMIFS({N("fpt_min")},{N("line_shift")},{key},{wkk})/SUMIFS({N("planned_production_time_min")},{N("line_shift")},{key},{wkk}))')
        c.font = F(8, False, GREY); c.number_format = "0.0%"
    grid = f"{L(HG)}{rr}:{L(HG + 12)}{rr}"
    body_cell(ws.cell(rr, 5, f'=IFERROR(AVERAGE({grid}),"")'), "0.0%", i % 2 == 1)
    body_cell(ws.cell(rr, 6, f'=IFERROR(STDEV({grid}),"")'), "0.0%", i % 2 == 1)
    body_cell(ws.cell(rr, 7, f'=IFERROR((D{rr}-E{rr})/F{rr},"")'), "+0.00;-0.00", i % 2 == 1)
    body_cell(ws.cell(rr, 8, f'=COUNTIFS({N("line_shift")},{key},{wk})'), "0", i % 2 == 1)
    body_cell(ws.cell(rr, 9, f'=IF(OR(H{rr}<{MINR},G{rr}=""),"Insufficient data",IF(G{rr}<=-{ZL},"LOW ANOMALY",IF(G{rr}>={ZL},"HIGH ANOMALY","In control")))'), None, i % 2 == 1, align="center")
SE = SC + len(combos) - 1
ws.conditional_formatting.add(f"I{SC}:I{SE}", CellIsRule(operator="equal", formula=['"LOW ANOMALY"'], fill=fill(RED), font=F(9, True, WHITE)))
ws.conditional_formatting.add(f"I{SC}:I{SE}", CellIsRule(operator="equal", formula=['"HIGH ANOMALY"'], fill=fill(GREEN), font=F(9, True, WHITE)))
ws.conditional_formatting.add(f"I{SC}:I{SE}", CellIsRule(operator="equal", formula=['"Insufficient data"'], font=F(9, False, GREY)))
ws.conditional_formatting.add(f"G{SC}:G{SE}", ColorScaleRule(start_type="num", start_value=-3, start_color="E57373", mid_type="num", mid_value=0, mid_color="FFFFFF", end_type="num", end_value=3, end_color="81C784"))
ws.conditional_formatting.add(f"B{SC}:C{SE}", FormulaRule(formula=[f"$N$6=$B{SC}&\" | \"&$C{SC}"], fill=fill(TEAL_L), font=F(9, True, TEAL)))
ws.cell(10, 12, "Scan summary").font = F(10, True, NAVY)
summ = [("Low anomalies", f'=COUNTIF(I{SC}:I{SE},"LOW ANOMALY")'), ("High anomalies", f'=COUNTIF(I{SC}:I{SE},"HIGH ANOMALY")'),
        ("In control", f'=COUNTIF(I{SC}:I{SE},"In control")'), ("Insufficient data", f'=COUNTIF(I{SC}:I{SE},"Insufficient data")')]
for i, (a, b) in enumerate(summ):
    ws.cell(11 + i, 12, a).font = F(9)
    c = ws.cell(11 + i, 14, b); calc_style(c, "0", True, NAVY)
DET = {"low": f"'05 Shift Variance Detector'!$N$11", "high": f"'05 Shift Variance Detector'!$N$12"}
block_note(ws, 16, 12, 30, 16, "How the detector works. For each line and shift, weekly OEE (time weighted) is compared with the mean and standard deviation of the same line and shift over the previous 13 weeks. "
           "A z score at or below minus the control limit is a LOW ANOMALY that warrants a shift review; at or above the limit is a HIGH ANOMALY worth learning from. "
           "Weeks with too few records are not scored, which prevents a single short run from raising a false alarm. "
           "With about three records per line, shift and week in this extract, weekly OEE is noisy; at a two sigma limit roughly one combination in twenty will flag by chance each week, so a flag should prompt a look, not a conclusion.")
# history for selected combo
r = SE + 3
r = section(ws, r, 2, "History for the selected line and shift", 12)
header_row(ws, r, 2, ["Week", "Records", "Weekly OEE", "Baseline mean", "Lower limit", "Upper limit", "z score", "Flag"])
H0 = r + 1
hist_weeks = [w for w in weeks_all if w <= END]
CH = 36  # hidden chart helper columns AJ:AM (NA for gaps so lines break cleanly)
for i, w in enumerate(hist_weeks):
    rr = H0 + i
    ws.cell(rr, 2, w.to_pydatetime())
    wk = f'{N("week_start")},B{rr}'
    ws.cell(rr, 3, f'=COUNTIFS({N("line_shift")},{KEY},{wk})')
    ws.cell(rr, 4, f'=IF(C{rr}<{MINR},"",SUMIFS({N("fpt_min")},{N("line_shift")},{KEY},{wk})/SUMIFS({N("planned_production_time_min")},{N("line_shift")},{KEY},{wk}))')
    if i >= 13:
        rng = f"D{rr - 13}:D{rr - 1}"
        ws.cell(rr, 5, f'=IFERROR(AVERAGE({rng}),"")')
        ws.cell(rr, 6, f'=IFERROR(E{rr}-{ZL}*STDEV({rng}),"")')
        ws.cell(rr, 7, f'=IFERROR(E{rr}+{ZL}*STDEV({rng}),"")')
        ws.cell(rr, 8, f'=IF(OR(D{rr}="",E{rr}=""),"",IFERROR((D{rr}-E{rr})/STDEV({rng}),""))')
        ws.cell(rr, 9, f'=IF(H{rr}="","",IF(H{rr}<=-{ZL},"LOW",IF(H{rr}>={ZL},"HIGH","")))')
    for k, fm in zip(range(2, 10), ["yyyy-mm-dd", "0", "0.0%", "0.0%", "0.0%", "0.0%", "+0.00;-0.00", None]):
        ws.cell(rr, k).font = F(9)
        if fm:
            ws.cell(rr, k).number_format = fm
    c = ws.cell(rr, CH, f'=IF(D{rr}="",0,D{rr})'); c.font = F(8, False, GREY)
    for k in range(1, 4):
        src = L(4 + k)
        if i < 13:
            continue
        prev = f"{L(CH + k)}{rr - 1}" if i > 13 else f"IFERROR(AVERAGE(D{H0}:D{rr - 1}),0)"
        c = ws.cell(rr, CH + k, f'=IF({src}{rr}="",{prev},{src}{rr})'); c.font = F(8, False, GREY)
HE = H0 + len(hist_weeks) - 1
for k, h in enumerate(["Weekly OEE", "Baseline mean", "Lower limit", "Upper limit"]):
    ws.cell(H0 - 1, CH + k, h)
    ws.column_dimensions[L(CH + k)].hidden = True
ws.conditional_formatting.add(f"I{H0}:I{HE}", CellIsRule(operator="equal", formula=['"LOW"'], fill=fill(RED), font=F(9, True, WHITE)))
ws.conditional_formatting.add(f"I{H0}:I{HE}", CellIsRule(operator="equal", formula=['"HIGH"'], fill=fill(GREEN), font=F(9, True, WHITE)))
ws.cell(r - 1, 11, "Weeks flagged LOW (history)").font = F(9, True)
c = ws.cell(r - 1, 14, f'=COUNTIF(I{H0}:I{HE},"LOW")'); calc_style(c, "0", True, RED)
ch = BarChart(); ch.type = "col"; ctitle(ch, "Selected line and shift: weekly OEE against 13 week control limits")
ch.add_data(Reference(ws, min_col=CH, min_row=H0 + 13, max_row=HE), titles_from_data=False)
ch.series[0].tx = None
from openpyxl.chart.series import SeriesLabel
ch.series[0].tx = SeriesLabel(v="Weekly OEE (no bar = no production)")
ch.series[0].graphicalProperties.solidFill = "8FA9C9"; ch.gapWidth = 30
ch.set_categories(Reference(ws, min_col=2, min_row=H0 + 13, max_row=HE))
lc = LineChart()
for k, (nm_, colr, dash) in enumerate([("Baseline mean", NAVY, None), ("Lower limit", RED, "dash"), ("Upper limit", GREEN, "dash")]):
    lc.add_data(Reference(ws, min_col=CH + 1 + k, min_row=H0 + 13, max_row=HE), titles_from_data=False)
    s_ = lc.series[-1]; s_.tx = SeriesLabel(v=nm_)
    s_.graphicalProperties.line.solidFill = colr; s_.graphicalProperties.line.width = 15000
    if dash:
        s_.graphicalProperties.line.dashStyle = dash
ch += lc
ch.y_axis.number_format = "0%"; ch.x_axis.number_format = "mmm yy"; ch.height = 8.5; ch.width = 20; ch.legend.position = "b"
ch.visible_cells_only = False
ws.add_chart(ch, f"K{H0}")
for k in range(HG, HG + 13):
    ws.column_dimensions[L(k)].hidden = True
ws.column_dimensions["N"].hidden = False
ws.column_dimensions["I"].width = 17
print_fit(ws, f"A1:Q{H0 + 24}", landscape=False, tall=2)

# ================================================================= 06 Drill down navigator
ws = S["06 Drill Down Navigator"]
setup(ws, "06  Line Level Drill Down Navigator", "Deliverable 5. Plant to line to shift to root cause in two selections. Each level lists every member, highlights the current path and shows the cause mix at the bottom.", cols=14, width_last=11)
ws.column_dimensions["B"].width = 22
selector(ws, "B6", LS["A"], "All Plants", "LEVEL 1  PLANT", "B5")
selector(ws, "D6", LS["B"], "All Lines", "LEVEL 2  LINE", "D5"); ws.merge_cells("D6:E6")
selector(ws, "G6", LS["C"], "All Shifts", "LEVEL 3  SHIFT", "G5"); ws.merge_cells("G6:H6")
selector(ws, "J6", LS["F"], "Last 52 weeks", "PERIOD", "J5"); ws.merge_cells("J6:K6")
ref = crit_block(ws, 5, 30, "$B$6", "$D$6", "$G$6", "$J$6")
HDR = ["OEE", "Availability", "Performance", "Quality", "Downtime hours", "Events"]


def level_rows(ws, top, members, member_col, crit_str, highlight):
    for i, m in enumerate(members):
        rr = top + i
        body_cell(ws.cell(rr, 2, m), None, i % 2 == 1)
        cr_ = f"{crit_str},{N(member_col)},$B{rr}"
        vals = [f'=IFERROR(SUMIFS({N("fpt_min")},{cr_})/SUMIFS({N("planned_production_time_min")},{cr_}),"")',
                f'=IFERROR(SUMIFS({N("avail_time_min")},{cr_})/SUMIFS({N("planned_production_time_min")},{cr_}),"")',
                f'=IFERROR(SUMIFS({N("perf_time_min")},{cr_})/SUMIFS({N("avail_time_min")},{cr_}),"")',
                f'=IFERROR(SUMIFS({N("fpt_min")},{cr_})/SUMIFS({N("perf_time_min")},{cr_}),"")',
                f"=SUMIFS({N('downtime_minutes')},{cr_})/60", f"=COUNTIFS({cr_})"]
        fm = ["0.0%", "0.0%", "0.0%", "0.0%", "#,##0.0", "#,##0"]
        for j, v in enumerate(vals):
            body_cell(ws.cell(rr, 3 + j, v), fm[j], i % 2 == 1)
    e = top + len(members) - 1
    ws.conditional_formatting.add(f"B{top}:H{e}", FormulaRule(formula=[f"$B{top}={highlight}"], fill=fill(TEAL_L), font=F(9, True, TEAL)))
    ws.conditional_formatting.add(f"C{top}:C{e}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=0.8, color="5EA8A0"))
    return e


r = section(ws, 8, 2, "Level 1: plants (period filter)", 7)
header_row(ws, r, 2, ["Plant"] + HDR)
e = level_rows(ws, r + 1, PLANTS, "plant_site", crit(ref, ("period",)), "$B$6")
r = section(ws, e + 2, 2, "Level 2: lines within the selected plant", 7)
header_row(ws, r, 2, ["Line"] + HDR)
e = level_rows(ws, r + 1, LINES, "production_line", crit(ref, ("plant", "period")), "$D$6")
r = section(ws, e + 2, 2, "Level 3: shifts within the selected plant and line", 7)
header_row(ws, r, 2, ["Shift"] + HDR)
e = level_rows(ws, r + 1, SHIFTS, "shift", crit(ref, ("plant", "line", "period")), "$G$6")
r = section(ws, e + 2, 2, "Level 4: root causes for the full selection", 7)
header_row(ws, r, 2, ["Reason", "Category", "Six Big Losses class", "Downtime hours", "Share", "Events", "Mean repair (min)"])
R4 = r + 1
CRall = crit(ref)
for i, rs in enumerate(REASONS):
    rr = R4 + i
    cr_ = f"{CRall},{N('reason')},$B{rr}"
    vals = [rs, f"=INDEX(tax_category,MATCH(B{rr},tax_reason,0))", f"=INDEX(tax_bigloss,MATCH(B{rr},tax_reason,0))",
            f"=SUMIFS({N('downtime_minutes')},{cr_})/60", f"=IFERROR(E{rr}/SUM($E${R4}:$E${R4 + 7}),0)", f"=COUNTIFS({cr_})",
            f'=IFERROR(AVERAGEIFS({N("repair_time_minutes")},{cr_}),"")']
    fm = [None, None, None, "#,##0.0", "0.0%", "#,##0", "0.0"]
    for j, v in enumerate(vals):
        body_cell(ws.cell(rr, 2 + j, v), fm[j], i % 2 == 1)
ws.conditional_formatting.add(f"F{R4}:F{R4 + 7}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=0.3, color="E29578"))
ws.cell(7, 2, f'="Current path:  "&$B$6&"  >  "&$D$6&"  >  "&$G$6&"   ("&$J$6&")"').font = F(10, True, TEAL)
section(ws, 8, 10, "Jump to", 4)
jumps = [("Pareto for deeper ranking", "03 Root Cause Pareto"), ("Fingerprint of a line", "04 Downtime Fingerprint"),
         ("Shift anomaly scan", "05 Shift Variance Detector"), ("Weekly scorecard", "07 Weekly OEE Scorecard"), ("Command center", "02 OEE Command Center")]
for i, (a, sh) in enumerate(jumps):
    c = ws.cell(9 + i, 10, a); c.hyperlink = f"#'{sh}'!A1"; c.font = Font(name=FONT, size=9, color=TEAL, underline="single")
block_note(ws, 15, 10, 26, 14, "Two selections to a root cause. Choose a line in Level 2 (one selection) and a shift in Level 3 (second selection); Level 4 immediately lists the root causes for that line and shift. "
           "Choosing a plant first narrows every level to that site. Leaving any selector on All keeps that level aggregated, so the same sheet answers plant, line, shift and combined questions.")
for k in (30, 31):
    ws.column_dimensions[L(k)].hidden = True
ws.column_dimensions["D"].width = 20
print_fit(ws, "A1:O62", landscape=False)

# ================================================================= 07 Weekly scorecard
ws = S["07 Weekly OEE Scorecard"]
setup(ws, "07  Weekly OEE Scorecard", "Deliverable 6. Print ready weekly scorecard by line. Choose the week and plant, then File > Export > PDF; the print area is preset to one landscape page.", cols=15, width_last=10)
ws.column_dimensions["B"].width = 12
selector(ws, "B6", LS["J"], wk_desc[1].to_pydatetime(), "WEEK STARTING", "B5"); ws.merge_cells("B6:C6"); ws["B6"].number_format = "dd mmm yyyy"
selector(ws, "E6", LS["A"], "All Plants", "PLANT", "E5"); ws.merge_cells("E6:G6")
ws["I6"] = '="Week "&TEXT($B$6,"dd mmm")&" to "&TEXT($B$6+6,"dd mmm yyyy")&"  |  "&$E$6'
ws["I6"].font = F(11, True, NAVY)
ws["AD5"] = '=IF($E$6="All Plants","*",$E$6)'; ws["AD5"].font = F(8, False, GREY)
PC = f"{N('plant_site')},$AD$5"
r = section(ws, 8, 2, "Line scorecard", 14)
header_row(ws, r, 2, ["Line", "OEE this week", "OEE prior week", "13 week baseline", "vs baseline (pts)", "Availability", "Performance", "Quality",
                      "Downtime hours", "Top reason this week", "Events", "Status"], height=34)
SCT = r + 1
wk = f'{N("week_start")},$B$6'
wkp = f'{N("week_start")},$B$6-7'
base = f'{N("week_start")},">="&($B$6-91),{N("week_start")},"<"&$B$6'
RG = 32  # hidden reason grid AF
for j, rs in enumerate(REASONS):
    ws.cell(SCT - 1, RG + j, rs).font = F(8, False, GREY)
for i, ln in enumerate(LINES + ["All lines"]):
    rr = SCT + i
    tot = ln == "All lines"
    lc_ = N("production_line") + "," + ('"*"' if tot else f"$B{rr}")
    cr_ = f"{PC},{lc_}"
    ff = lambda num, den, w: f'=IFERROR(SUMIFS({N(num)},{cr_},{w})/SUMIFS({N(den)},{cr_},{w}),"")'
    for j, rs in enumerate(REASONS):
        c = ws.cell(rr, RG + j, f'=SUMIFS({N("downtime_minutes")},{cr_},{wk},{N("reason")},"{rs}")+{j + 1}/1000000'); c.font = F(8, False, GREY)
    grid = f"${L(RG)}{rr}:${L(RG + 7)}{rr}"
    vals = [ln, ff("fpt_min", "planned_production_time_min", wk), ff("fpt_min", "planned_production_time_min", wkp),
            ff("fpt_min", "planned_production_time_min", base), f'=IFERROR((C{rr}-E{rr})*100,"")',
            ff("avail_time_min", "planned_production_time_min", wk), ff("perf_time_min", "avail_time_min", wk), ff("fpt_min", "perf_time_min", wk),
            f"=SUMIFS({N('downtime_minutes')},{cr_},{wk})/60",
            f'=IF(L{rr}=0,"",INDEX(${L(RG)}${SCT - 1}:${L(RG + 7)}${SCT - 1},MATCH(MAX({grid}),{grid},0)))',
            f"=COUNTIFS({cr_},{wk})",
            f'=IF(L{rr}=0,"No production",IF(F{rr}<=-5,"Below baseline",IF(F{rr}>=5,"Above baseline","On baseline")))']
    fm = [None, "0.0%", "0.0%", "0.0%", "+0.0;-0.0", "0.0%", "0.0%", "0.0%", "#,##0.0", None, "#,##0", None]
    for j, v in enumerate(vals):
        c = ws.cell(rr, 2 + j, v)
        body_cell(c, fm[j], i % 2 == 1, align="center" if j > 0 else None)
        if tot:
            c.font = F(9, True, NAVY); c.fill = fill(TEAL_L)
SCE = SCT + 12
ws.conditional_formatting.add(f"M{SCT}:M{SCE}", CellIsRule(operator="equal", formula=['"Below baseline"'], fill=fill(RED_L), font=F(9, True, RED)))
ws.conditional_formatting.add(f"M{SCT}:M{SCE}", CellIsRule(operator="equal", formula=['"Above baseline"'], fill=fill(GREEN_L), font=F(9, True, GREEN)))
ws.conditional_formatting.add(f"C{SCT}:C{SCE - 1}", ColorScaleRule(start_type="min", start_color="F4A6A6", mid_type="percentile", mid_value=50, mid_color="FFF2CC", end_type="max", end_color="9BD3AE"))
ws.column_dimensions["K"].width = 20
r = SCE + 2
r = section(ws, r, 2, "Shift summary for the week", 8)
header_row(ws, r, 2, ["Shift", "OEE", "13 week baseline", "vs baseline (pts)", "Downtime hours", "Events"])
for i, sh in enumerate(SHIFTS):
    rr = r + 1 + i
    cr_ = f'{PC},{N("shift")},"{sh}"'
    vals = [sh, f'=IFERROR(SUMIFS({N("fpt_min")},{cr_},{wk})/SUMIFS({N("planned_production_time_min")},{cr_},{wk}),"")',
            f'=IFERROR(SUMIFS({N("fpt_min")},{cr_},{base})/SUMIFS({N("planned_production_time_min")},{cr_},{base}),"")',
            f'=IFERROR((C{rr}-D{rr})*100,"")', f"=SUMIFS({N('downtime_minutes')},{cr_},{wk})/60", f"=COUNTIFS({cr_},{wk})"]
    fm = [None, "0.0%", "0.0%", "+0.0;-0.0", "#,##0.0", "#,##0"]
    for j, v in enumerate(vals):
        body_cell(ws.cell(rr, 2 + j, v), fm[j], i % 2 == 1, align="center" if j else None)
note(ws, r + 5, 2, "Status rule: Below baseline when the week is 5 or more OEE points under the line's 13 week baseline, Above baseline when 5 or more points over. Baseline is time weighted across the 13 weeks before the selected week.", span=13)
note(ws, r + 6, 2, "Automated export: with the print area preset, File > Export > Create PDF produces the one page scorecard. To automate fully, record a two line macro (set the week cell, then ExportAsFixedFormat) or schedule it from Power Automate; the sheet needs no manual edits between weeks.", span=13)
for k in list(range(RG, RG + 8)) + [30]:
    ws.column_dimensions[L(k)].hidden = True
ws.column_dimensions["H"].width = 12; ws.column_dimensions["M"].width = 14
print_fit(ws, "A1:Q36")

# ================================================================= 08 Validation & Assumptions
ws = S["08 Validation & Assumptions"]
setup(ws, "08  OEE Validation, Assumptions and Limitations", "Validation against the last three months of line data, an independent Python reference, scale testing and the assumption register.", cols=10, width_last=12)
for k, w in zip("BCDEFGHIJK", [16, 12, 14, 14, 14, 14, 14, 12, 40, 4]):
    ws.column_dimensions[k].width = w
r = section(ws, 5, 2, "V1  Governed OEE by line for the last three months: Excel versus independent Python reference", 9)
months = sorted(clean.month_start.dropna().unique())[-3:]
months = [pd.Timestamp(m) for m in months]
header_row(ws, r, 2, ["Line", "Month", "Excel governed OEE", "Python reference", "Difference", "Reported OEE (legacy)", "Reported gap (pts)", "Check"])
V1 = r + 1
i = 0
for ln in LINES:
    for m in months:
        rr = V1 + i
        sub = clean[(clean.production_line == ln) & (clean.month_start == m)]
        pyref = float(sub.fpt_min.sum() / sub.planned_production_time_min.sum())
        cr_ = f"{N('production_line')},$B{rr},{N('month_start')},$C{rr}"
        vals = [ln, m.to_pydatetime(), f"=SUMIFS({N('fpt_min')},{cr_})/SUMIFS({N('planned_production_time_min')},{cr_})", round(pyref, 10),
                f"=D{rr}-E{rr}", f"=AVERAGEIFS({N('oee_pct')},{cr_})/100", f"=(G{rr}-D{rr})*100", f'=IF(ABS(F{rr})<0.0001,"PASS","FAIL")']
        fm = [None, "mmm yyyy", "0.00%", "0.00%", "0.000000", "0.0%", "+0.0;-0.0", None]
        for j, v in enumerate(vals):
            body_cell(ws.cell(rr, 2 + j, v), fm[j], i % 2 == 1, align="center" if j else None)
        i += 1
V1E = V1 + i - 1
CS.PASS_RULES(ws, f"I{V1}:I{V1E}")
r = V1E + 2
ws.cell(r, 2, "Line months validated").font = F(10, True)
c = ws.cell(r, 5, f'=COUNTIF(I{V1}:I{V1E},"PASS")&" of "&COUNTA(I{V1}:I{V1E})'); calc_style(c, None, True, NAVY)
ws.cell(r + 1, 2, "Mean reported gap, last three months").font = F(10, True)
c = ws.cell(r + 1, 5, f"=AVERAGE(H{V1}:H{V1E})"); calc_style(c, '+0.0" pts";-0.0" pts"', True, RED)
VAL = {"v1": f"'08 Validation & Assumptions'!$E${r}", "gap3": f"'08 Validation & Assumptions'!$E${r + 1}"}
r += 3
r = section(ws, r, 2, "Validation log", 9)
header_row(ws, r, 2, ["ID", "Test", "", "Result (live)", "", "Status", "", "", "Method"])
for a_, b_ in ((3, 4), (5, 6), (7, 9)):
    ws.merge_cells(start_row=r, start_column=a_, end_row=r, end_column=b_)
tests = [("V1", "Governed OEE reproduces an independent Python calculation (3 months x 12 lines)", f"={VAL['v1']}", None, f'=IF(LEFT(E{{r}},2)=RIGHT(E{{r}},2),"PASS","FAIL")', "Excel SUMIFS time weighting versus pandas groupby, tolerance 0.01 points"),
         ("V2", "OEE identity holds at network level (OEE = A x P x Q)", f"={IDENT}", "0.000000", f'=IF(E{{r}}<0.000001,"PASS","FAIL")', "Absolute difference between OEE and product of components"),
         ("V3", "Every record carries a governed reason code", f"={TAXCOV}", "#,##0", f'=IF(E{{r}}=0,"PASS","FAIL")', "Blank governed codes after taxonomy mapping"),
         ("V4", "80 percent of downtime traced to five or fewer causes", f"={PAR['n80']}", '0" causes"', f'=IF(E{{r}}<=5,"MET","NOT MET")', "Reason level Pareto, all data (03). See limitation L1"),
         ("V5", "Drill down from plant to root cause", "2 selections", None, '="MET"', "Line then shift selector on 06 reveals Level 4 root causes"),
         ("V6", "Handles 50,000+ events without degradation", SCALE, None, '="PARTLY MET"', "Clean Data tiled 4x to 61,000 events in a test copy. Selector changes: 3 to 8 s at 61,000 versus 1 to 2 s at 15,250 (LibreOffice headless, single thread); full hard recalculation 88 s. Scaling is linear; for sustained 50,000+ volumes load Clean Data into the Excel Data Model (Power Pivot)."),
         ("V7", "Cleaning row reconciliation", "C11_PLACEHOLDER", "#,##0", f'=IF(E{{r}}=0,"PASS","FAIL")', "Raw rows less blanks and duplicates equals Clean Data rows"),
         ("V8", "Reported OEE agrees with governed OEE", f"={ENG['gap']}", '+0.0" pts";-0.0" pts"', f'=IF(ABS(E{{r}})<=1,"PASS","FINDING")', "Legacy oee_pct mean minus governed OEE, all data")]
VT = r + 1
for i, (a, b, c_, fm, st, m) in enumerate(tests):
    rr = VT + i
    ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=4)
    ws.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=6)
    ws.merge_cells(start_row=rr, start_column=7, end_row=rr, end_column=9)
    for k, v, f_ in ((2, a, None), (3, b, None), (5, c_, fm), (7, st.format(r=rr), None), (10, m, None)):
        cc = ws.cell(rr, k, v); body_cell(cc, f_, i % 2 == 1, align="wrap" if k in (3, 10) else "center")
    ws.row_dimensions[rr].height = row_h(m, 38) if len(m) > 80 else 32
    if a == "V7":
        V7ROW = rr
CS.PASS_RULES(ws, f"G{VT}:G{VT + 7}")
for lab_, col_, fc_ in (("MET", GREEN_L, GREEN), ("NOT MET", RED_L, RED), ("FINDING", AMBER_L, AMBER), ("PARTLY MET", AMBER_L, AMBER)):
    ws.conditional_formatting.add(f"G{VT}:G{VT + 7}", CellIsRule(operator="equal", formula=[f'"{lab_}"'], fill=fill(col_), font=F(9, True, fc_)))
r = VT + 9
r = section(ws, r, 2, "Assumption register", 9)
asm = [("A1", "OEE definition", "A x P x Q, time weighted by planned production time (Nakajima). The stored oee_pct is not used as a metric."),
       ("A2", "Authoritative OEE inputs", "availability_pct, performance_pct and quality_pct as recorded, after validation in C7 and imputation by line median in C9."),
       ("A3", "Downtime taxonomy", "Description is authoritative; governed code, category and Six Big Losses class from M1. Editable mapping."),
       ("A4", "Baseline window", "13 weeks preceding the scored week (catalogue brief). Editable on 05."),
       ("A5", "Control limit", "2 standard deviations (editable on 05). About 5 percent false alarms per stable line and shift."),
       ("A6", "Minimum sample", "Weeks with fewer than 2 records for a line and shift are not scored (editable on 05)."),
       ("A7", "Scorecard status band", "Plus or minus 5 OEE points versus the 13 week baseline."),
       ("A8", "Records without a start date", "152 records (1.0 percent) have no parseable date; included in all data totals, excluded from weekly and period views.")]
header_row(ws, r, 2, ["ID", "Assumption", "", "Detail", "", "", "", "", ""])
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4); ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=10)
for i, (a, b, c_) in enumerate(asm):
    rr = r + 1 + i
    ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=4); ws.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=10)
    for k, v in ((2, a), (3, b), (5, c_)):
        body_cell(ws.cell(rr, k, v), None, i % 2 == 1, align="wrap")
    ws.row_dimensions[rr].height = 28
r = r + len(asm) + 2
r = section(ws, r, 2, "Limitations", 9)
lims = ["L1  Downtime in this extract is spread almost evenly across the eight governed reasons (each between about 11.7 and 12.9 percent of minutes), so seven of eight reasons are needed to reach 80 percent, and splitting by line does not create concentration (top five line and reason hotspots hold about 6 percent). The Pareto target is not met because the data contains no vital few; the correct response is a systemic programme on availability and performance losses, not a cause specific one.",
        "L2  The legacy oee_pct field averages about 21 points above the governed A x P x Q. Either the field is computed with a different, undocumented method or the component fields are mis-recorded; both must be resolved with the MES owner before OEE is used for incentives.",
        "L3  machine_id, operator_id and fault_code are unique per record, so asset level fingerprints and operator analysis are not possible from this extract.",
        "L4  Stored reason codes and categories are statistically unrelated to the descriptions (M1); any historic report built on those codes should be treated as unreliable."]
for t in lims:
    r = note(ws, r, 2, t, span=9, color=INK)
print_fit(ws, "A1:K90", landscape=False)

# ================================================================= 01 Executive summary
ws = S["01 Executive Summary"]
setup(ws, "Executive Summary: OEE and Downtime", "One page view for the operations leadership review. All figures are live from Clean Data.", cols=12, width_last=10.5)
kp = [("GOVERNED OEE", f"={ENG['oee']}", "0.0%", "All data, A x P x Q", TEAL), ("REPORTED OEE (LEGACY)", f"={ENG['rep']}", "0.0%", "Stored oee_pct average", RED),
      ("REPORTING GAP", f"={ENG['gap']}", '+0.0" pts";-0.0" pts"', "Legacy minus governed", AMBER), ("LINES x SHIFTS", "12 x 3", "@", "4 plants, 15,250 records", NAVY)]
for i, (a, b, fm, s_, col) in enumerate(kp):
    kpi(ws, 5, 2 + i * 3, a, b, fm, s_, 3, col)
kp2 = [("AVAILABILITY", f"={ENG['a']}", "0.0%", "Operating over planned time", NAVY), ("PERFORMANCE", f"={ENG['p']}", "0.0%", "Speed losses", NAVY),
       ("QUALITY", f"={ENG['q']}", "0.0%", "Defect and rework losses", NAVY), ("CAUSES FOR 80% OF DOWNTIME", f"={PAR['n80']}", '0" of 8"', "Reason level Pareto", AMBER)]
for i, (a, b, fm, s_, col) in enumerate(kp2):
    kpi(ws, 9, 2 + i * 3, a, b, fm, s_, 3, col)
r = 13
r = section(ws, r, 2, "Success metrics scorecard", 12)
header_row(ws, r, 2, ["Success metric (catalogue)", "", "", "Target", "", "Result", "", "Status", "", "Evidence", "", ""])
for k in (2, 5, 7, 9, 11):
    ws.merge_cells(start_row=r, start_column=k, end_row=r, end_column=k + (2 if k in (2, 11) else 1))
sm = [("OEE calculation validated against three months of line data", "3 months", f"={VAL['v1']}", f'=IF(LEFT(G{{r}},2)=RIGHT(G{{r}},2),"MET","NOT MET")', "08 V1: 36 line months reproduce an independent Python calculation"),
      ("80 percent of downtime losses traced to five or fewer causes", "5 or fewer", f"={PAR['n80']}", f'=IF(G{{r}}<=5,"MET","NOT MET")', "Reason level Pareto; causes are evenly spread (08 L1)"),
      ("Drill down navigation within two clicks", "2 clicks", "2 selections", '="MET"', "06 Drill Down Navigator"),
      ("Workbook handles 50,000+ events", "50,000+", SCALE, '="PARTLY MET"', "Works at 61,000 events; response slows linearly to 3 to 8 s per selector (08 V6)")]
for i, (a, b, c_, e_, ev) in enumerate(sm):
    rr = r + 1 + i
    for k, v, f_ in ((2, a, None), (5, b, None), (7, c_, '0" causes"' if "n80" in str(c_) or "Pareto" in str(c_) else None), (9, e_.format(r=rr), None), (11, ev, None)):
        cc = ws.cell(rr, k, v); body_cell(cc, f_, i % 2 == 1, align="center" if k in (5, 7, 9) else "wrap")
        ws.merge_cells(start_row=rr, start_column=k, end_row=rr, end_column=k + (2 if k in (2, 11) else 1))
    ws.row_dimensions[rr].height = 34
ws.cell(r + 2, 7).number_format = '0" causes"'
for lab_, col_ in (("MET", GREEN_L), ("NOT MET", RED_L), ("PARTLY MET", AMBER_L)):
    ws.conditional_formatting.add(f"I{r + 1}:I{r + 4}", CellIsRule(operator="equal", formula=[f'"{lab_}"'], fill=fill(col_), font=F(9, True, INK)))
r += 6
r = section(ws, r, 2, "Key findings (live)", 12)
finds = [f'="1.  Governed OEE is "&TEXT({ENG["oee"]},"0.0%")&", not the "&TEXT({ENG["rep"]},"0.0%")&" reported by the legacy field. The "&TEXT({ENG["gap"]},"0.0")&" point gap means current OEE reporting overstates performance and cannot be reconciled to its own components."',
         f'="2.  Availability ("&TEXT({ENG["a"]},"0.0%")&") and performance ("&TEXT({ENG["p"]},"0.0%")&") drive the loss; quality ("&TEXT({ENG["q"]},"0.0%")&") is the strongest pillar. Speed and stoppage losses deserve the first improvement cycle."',
         f'="3.  There is no vital few: "&{PAR["n80"]}&" of 8 reasons are needed to reach 80 percent, the largest reason holds only "&TEXT({PAR["top"]},"0.0%")&", and the five worst line and reason hotspots together hold just "&TEXT({PAR["hot5"]},"0.0%")&" of downtime."',
         f'="4.  Fingerprint test: "&{FPS["concl"]}&" (p = "&TEXT({FPS["p"]},"0.000")&", Cramer V = "&TEXT({FPS["v"]},"0.000")&")."',
         '="5.  Governance: stored reason codes and planned/unplanned categories disagree with the descriptions on most records (C8, M1). The governed taxonomy replaces them and should be enforced at the point of entry."']
for f_ in finds:
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=13)
    c = ws.cell(r, 2, f_); c.font = F(10); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 42
    r += 1
r += 1
r = section(ws, r, 2, "Recommendations", 12)
recs = ["Treat the loss as systemic: with no vital few cause and no line specific fingerprint, launch a network wide availability and performance programme (standard changeover, planned maintenance adherence, speed loss tracking) rather than chasing single causes.",
        "Retire the legacy oee_pct field and publish only the governed A x P x Q; agree the calculation with the MES owner and lock it in the data model.",
        "Enforce the eight reason governed taxonomy as a drop down at the HMI so code, category and description can no longer disagree.",
        "Run the weekly scorecard (07) in the Monday production meeting and review every LOW anomaly from the detector (05) within 48 hours.",
        "Capture a persistent asset identifier so the fingerprint radar can move from line level to machine level."][:5]
for i, t in enumerate(recs):
    r = note(ws, r, 2, f"{i + 1}.  {t}", span=12, color=INK, size=10, height=row_h(t, 125))
print_fit(ws, "A1:N45", landscape=False)

# ================================================================= cleaning sheets
CS.C1(ctx); CS.C2(ctx); CS.C3(ctx); CS.C4(ctx); CS.C5(ctx); CS.C6(ctx); CS.C7(ctx)
CS.C8(ctx, "Six rules were tested. The OEE definition (A x P x Q) takes precedence over the stored OEE figure; the operator's description takes precedence over stored codes and categories; "
          "percentage inputs that are bounded and validated take precedence over raw time and unit counts that break physical limits. Conflicting fields are kept for audit and flagged where they would otherwise mislead.")
CS.C9(ctx); CS.C10(ctx, FEATURES); CS.C11(ctx)
c11 = wb["C11 Cleaning Audit Log"]
diff_row = [c.row for row in c11.iter_rows(min_col=3, max_col=3) for c in row if c.value == "Difference"][0]
wb["08 Validation & Assumptions"].cell(V7ROW, 5).value = f"='C11 Cleaning Audit Log'!F{diff_row}"

# ================================================================= Cover
ws = S["Cover"]
setup(ws, "Manufacturing OEE and Downtime Root Cause Command Workbook", "Excel Project 2 of 10  |  Manufacturing  |  Data Analyst Portfolio  |  Prepared by Anthony Chinedu Echem", cols=12, width_last=11, back=False)
ws.column_dimensions["B"].width = 30; ws.column_dimensions["C"].width = 60
r = section(ws, 5, 2, "Business problem", 11)
r = note(ws, r, 2, "A precision parts manufacturer running 3 shifts across 12 CNC lines has OEE stuck around 61 percent. Downtime is logged in incompatible formats, so root causes are invisible and improvement budgets are guessed.", span=11, color=INK, size=10)
r = section(ws, r + 1, 2, "Core objective", 11)
r = note(ws, r, 2, "Unify downtime into one governed taxonomy, calculate OEE correctly by line and shift, and surface the causes and shifts that drive losses, starting from a fully audited cleaning of the supplied 15,394 row extract.", span=11, color=INK, size=10)
r = section(ws, r + 1, 2, "Workbook map", 11)
header_row(ws, r, 2, ["Sheet", "Purpose", "Catalogue deliverable"]); ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
nav = [("01 Executive Summary", "One page KPIs, success metrics, findings and recommendations", "Executive summary"),
       ("02 OEE Command Center", "Interactive plant, line, shift and period dashboard", "Real time OEE command workbook"),
       ("03 Root Cause Pareto", "Reason, loss class and line hotspot Pareto with 80 percent test", "Root cause Pareto analysis tab"),
       ("04 Downtime Fingerprint", "Radar of each line's cause mix with chi square test", "Machine downtime fingerprint radar chart"),
       ("05 Shift Variance Detector", "13 week baseline control limits for all 36 line shifts", "Shift variance anomaly detector"),
       ("06 Drill Down Navigator", "Plant to line to shift to root cause", "Line level drill down navigation"),
       ("07 Weekly OEE Scorecard", "Print ready one page weekly scorecard", "Weekly automated OEE scorecard export"),
       ("08 Validation & Assumptions", "Three month validation, scale test, assumptions, limitations", "Validation (success metrics)"),
       ("M1 Downtime Taxonomy", "Governed reason list and proof of stored code failure", "Approach step 1: taxonomy"),
       ("M2 OEE Engine", "Time weighted A, P, Q and OEE by line, shift and plant", "Approach step 2: OEE engine")] + \
      [(n, d, "Data cleaning procedure") for n, d in [("C1 Data Profile", "Step 1: baseline profile"), ("C2 Structural Integrity", "Step 2: blank rows and duplicates"),
                                                      ("C3 Key Collision Repair", "Step 3: record_id repair"), ("C4 Text Standardization", "Step 4: label governance"),
                                                      ("C5 Boolean Normalization", "Step 5: flags to TRUE/FALSE"), ("C6 Date Standardization", "Step 6: four formats to dates"),
                                                      ("C7 Numeric Validation", "Step 7: limits and outlier fences"), ("C8 Cross Field Consistency", "Step 8: business rules"),
                                                      ("C9 Missing Value Treatment", "Step 9: imputation rules"), ("C10 Feature Engineering", "Step 10: governed features"),
                                                      ("C11 Cleaning Audit Log", "Step 11: audit trail and reconciliation")]] + \
      [("Clean Data", "Analysis ready table (named ranges cd_*)", "Output of cleaning"), ("Raw Data", "Supplied dataset, unchanged", "Source"), ("Data Dictionary", "Supplied dictionary, unchanged", "Source")]
for i, (s_, p_, d_) in enumerate(nav):
    rr = r + 1 + i
    c = ws.cell(rr, 2, s_); c.hyperlink = f"#'{s_}'!A1"; body_cell(c, None, i % 2 == 1); c.font = Font(name=FONT, size=9, color=TEAL, underline="single")
    body_cell(ws.cell(rr, 3, p_), None, i % 2 == 1)
    ws.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=7); body_cell(ws.cell(rr, 4, d_), None, i % 2 == 1)
r = r + len(nav) + 2
r = section(ws, r, 2, "Conventions", 11)
for i, (a, b, fc, bg) in enumerate([("Blue text on yellow", "Input or assumption you may change", "0000FF", INPUT_FILL), ("Black text", "Formula; do not overwrite", INK, WHITE),
                                     ("Green text", "Link to another sheet", "008000", WHITE), ("Teal bordered cell", "Drop down selector (choose from list)", NAVY, WHITE)]):
    c = ws.cell(r + i, 2, a); c.font = F(9, True, fc); c.fill = fill(bg); c.border = BORDER
    ws.cell(r + i, 3, b).font = F(9)
print_fit(ws, "A1:M60", landscape=False)

order = names + ["C1 Data Profile", "C2 Structural Integrity", "C3 Key Collision Repair", "C4 Text Standardization", "C5 Boolean Normalization",
                 "C6 Date Standardization", "C7 Numeric Validation", "C8 Cross Field Consistency", "C9 Missing Value Treatment",
                 "C10 Feature Engineering", "C11 Cleaning Audit Log", "Clean Data", "Raw Data", "Data Dictionary", "Lists"]
wb._sheets = [wb[n] for n in order]
for n in order:
    if n.startswith("C") and n[1].isdigit():
        wb[n].sheet_properties.tabColor = "7C8DA6"
    elif n.startswith("M") and n[1].isdigit():
        wb[n].sheet_properties.tabColor = "2F5D8A"
    elif n[:2].isdigit():
        wb[n].sheet_properties.tabColor = TEAL
wb["Cover"].sheet_properties.tabColor = NAVY
wb["Clean Data"].sheet_properties.tabColor = TEAL
wb.active = 0
unsmooth(wb)
wb.calculation.fullCalcOnLoad = True
os.makedirs("workbook", exist_ok=True)
wb.save(OUT)
print("saved", OUT)
