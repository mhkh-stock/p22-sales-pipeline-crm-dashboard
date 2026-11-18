#!/usr/bin/env python3
"""P22 — build the deliverable workbook for CitySoft Solutions (SAMPLE):

  Dashboard              — KPIs (deals, pipeline, open + weighted open,
                           closed / won / lost, win rate, revenue won,
                           value lost, average deal sizes, average sales
                           cycles, July + August forecast, run-rate,
                           coverage), top rep, lowest win rate via
                           _xlfn.MINIFS, biggest leak — all live
                           formulas over Pipeline
  Pipeline               — all 156 deals (values for the source
                           columns, live formulas for weighted value,
                           is_closed, is_won and the reached-stage
                           helpers)
  Stages                 — funnel (deals reached per stage, conversion
                           from previous, share of all deals) + current
                           stage snapshot — live formulas
  Reps                   — 6 reps: deals, pipeline, open, closed, won,
                           win rate, revenue won, average deal sizes,
                           average cycles — COUNTIFS / SUMIFS /
                           AVERAGEIFS over Pipeline
  Monthly Trend          — 6 created months, same shape, plus won
                           revenue by close month and the 3-month
                           run-rate + chart
  Forecast               — Jul / Aug 2026 from expected close dates
                           (SUMIFS on expected_close_date) + run-rate
                           and coverage
  Breakdown              — by product line and by source, same shape
  Method                 — how to read it, the stage probabilities,
                           the win-rate and forecast rules, their
                           limits, how to re-verify

    python3 src/build_workbook.py
"""
import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"
OUT = PROJECT / "outputs" / "pipeline_report_SAMPLE.xlsx"

REPS = ["Ayesha Khan", "Bilal Ahmed", "Sana Malik", "Usman Tariq",
        "Hira Shah", "Daniyal Raza"]
PRODUCTS = ["ERP Implementation", "Cloud Migration", "Mobile App",
            "Website Development", "Digital Marketing"]
SOURCES = ["Referral", "Website", "LinkedIn", "Cold Outreach"]
MONTHS = [f"2026-{m:02d}" for m in range(1, 7)]
FORECAST_MONTHS = ["2026-07", "2026-08"]
STAGE_ORDER = ["Lead", "Qualified", "Proposal", "Negotiation",
               "Won", "Lost"]
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
COUNT_FMT = "#,##0"
PKR_FMT = "#,##0"
PCT_FMT = "0.00%"
AVG_FMT = "#,##0.00"
DAYS_FMT = "0.0"


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT


def main():
    raw = list(csv.DictReader(
        open(RAW / "deals_SAMPLE.csv", encoding="utf-8")))
    n = len(raw)
    last = n + 1  # last data row on Pipeline (row 1 = header)

    wb = Workbook()

    # ── Pipeline ────────────────────────────────────────────────────
    pl = wb.active
    pl.title = "Pipeline"
    pl.append(["deal_id", "created_date", "month", "close_date",
               "expected_close_date", "rep", "product_line", "source",
               "stage", "lost_at_stage", "lost_reason",
               "deal_value_pkr", "probability_pct",
               "weighted_value_pkr", "sales_cycle_days", "is_closed",
               "is_won", "reached_qualified", "reached_proposal",
               "reached_negotiation"])
    for r in raw:
        pl.append([r["deal_id"], r["created_date"], r["month"],
                   r["close_date"] or None,
                   r["expected_close_date"] or None, r["rep"],
                   r["product_line"], r["source"], r["stage"],
                   r["lost_at_stage"] or None,
                   r["lost_reason"] or None,
                   int(r["deal_value_pkr"]),
                   int(r["probability_pct"]), None,
                   int(r["sales_cycle_days"])
                   if r["sales_cycle_days"] else None,
                   None, None, None, None, None])
    for i in range(2, last + 1):
        pl.cell(row=i, column=14).value = f"=ROUND(L{i}*M{i}/100,0)"
        pl.cell(row=i, column=16).value = (
            f'=IF(OR(I{i}="Won",I{i}="Lost"),1,0)')
        pl.cell(row=i, column=17).value = f'=IF(I{i}="Won",1,0)'
        # reached helpers: current stage at-or-beyond, or lost there
        # or later; Won deals reached every stage.
        pl.cell(row=i, column=18).value = (
            f'=IF(OR(I{i}="Won",I{i}="Qualified",I{i}="Proposal",'
            f'I{i}="Negotiation",J{i}="Qualified",J{i}="Proposal",'
            f'J{i}="Negotiation"),1,0)')
        pl.cell(row=i, column=19).value = (
            f'=IF(OR(I{i}="Won",I{i}="Proposal",I{i}="Negotiation",'
            f'J{i}="Proposal",J{i}="Negotiation"),1,0)')
        pl.cell(row=i, column=20).value = (
            f'=IF(OR(I{i}="Won",I{i}="Negotiation",'
            f'J{i}="Negotiation"),1,0)')
    style_header(pl, 20)
    pl.freeze_panes = "A2"
    widths = {"A": 9, "B": 12, "C": 9, "D": 12, "E": 17, "F": 14,
              "G": 20, "H": 14, "I": 12, "J": 14, "K": 24, "L": 15,
              "M": 13, "N": 16, "O": 15, "P": 10, "Q": 9, "R": 16,
              "S": 16, "T": 18}
    for col, w in widths.items():
        pl.column_dimensions[col].width = w
    for r in range(2, last + 1):
        for col in (12, 14):
            pl.cell(row=r, column=col).number_format = PKR_FMT

    # ── shared group sheet (reps / months / products / sources) ────
    def group_sheet(name, labels, key_col, title):
        ws = wb.create_sheet(name)
        ws.append(["name", "deals", "pipeline_value_pkr", "open_deals",
                   "open_value_pkr", "weighted_open_pkr", "closed_deals",
                   "won_deals", "lost_deals", "win_rate",
                   "revenue_won_pkr", "value_lost_pkr",
                   "avg_deal_size_pkr", "avg_won_deal_size_pkr",
                   "avg_cycle_won_days", "avg_cycle_closed_days"])
        for j, label in enumerate(labels, start=2):
            ws.cell(row=j, column=1).value = label
            # Pipeline cols: A deal_id, C month, F rep, G product,
            # H source, I stage, L value, N weighted, O cycle,
            # P is_closed, Q is_won
            ws.cell(row=j, column=2).value = (
                f'=COUNTIF(Pipeline!${key_col}$2:${key_col}${last},'
                f'$A{j})')
            ws.cell(row=j, column=3).value = (
                f'=SUMIF(Pipeline!${key_col}$2:${key_col}${last},'
                f'$A{j},Pipeline!$L$2:$L${last})')
            ws.cell(row=j, column=4).value = (
                f'=COUNTIFS(Pipeline!${key_col}$2:${key_col}${last},'
                f'$A{j},Pipeline!$P$2:$P${last},0)')
            ws.cell(row=j, column=5).value = (
                f'=SUMIFS(Pipeline!$L$2:$L${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{j},'
                f'Pipeline!$P$2:$P${last},0)')
            ws.cell(row=j, column=6).value = (
                f'=SUMIFS(Pipeline!$N$2:$N${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{j},'
                f'Pipeline!$P$2:$P${last},0)')
            ws.cell(row=j, column=7).value = (
                f'=SUMIFS(Pipeline!$P$2:$P${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{j})')
            ws.cell(row=j, column=8).value = (
                f'=SUMIFS(Pipeline!$Q$2:$Q${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{j})')
            ws.cell(row=j, column=9).value = f"=G{j}-H{j}"
            ws.cell(row=j, column=10).value = (
                f'=IF(G{j}=0,"—",H{j}/G{j})')
            ws.cell(row=j, column=11).value = (
                f'=SUMIFS(Pipeline!$L$2:$L${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{j},'
                f'Pipeline!$Q$2:$Q${last},1)')
            ws.cell(row=j, column=12).value = (
                f'=SUMIFS(Pipeline!$L$2:$L${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{j},'
                f'Pipeline!$I$2:$I${last},"Lost")')
            ws.cell(row=j, column=13).value = (
                f'=IF(B{j}=0,"—",C{j}/B{j})')
            ws.cell(row=j, column=14).value = (
                f'=IF(H{j}=0,"—",K{j}/H{j})')
            ws.cell(row=j, column=15).value = (
                f'=IFERROR(AVERAGEIFS(Pipeline!$O$2:$O${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{j},'
                f'Pipeline!$Q$2:$Q${last},1),"—")')
            ws.cell(row=j, column=16).value = (
                f'=IFERROR(AVERAGEIFS(Pipeline!$O$2:$O${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{j},'
                f'Pipeline!$P$2:$P${last},1),"—")')
        style_header(ws, 16)
        ws.freeze_panes = "A2"
        ws.column_dimensions["A"].width = 22
        for col in "BCDEFGHIKLMN":
            ws.column_dimensions[col].width = 14
        ws.column_dimensions["J"].width = 10
        ws.column_dimensions["O"].width = 16
        ws.column_dimensions["P"].width = 18
        for r in range(2, 2 + len(labels)):
            for col in (3, 5, 6, 11, 12):
                ws.cell(row=r, column=col).number_format = PKR_FMT
            ws.cell(row=r, column=10).number_format = PCT_FMT
            for col in (13, 14):
                ws.cell(row=r, column=col).number_format = AVG_FMT
            for col in (15, 16):
                ws.cell(row=r, column=col).number_format = DAYS_FMT
        return ws

    reps_ws = group_sheet("Reps", REPS, "F", "Reps")
    months_ws = group_sheet("Monthly Trend", MONTHS, "C",
                            "Monthly Trend")

    # Monthly Trend: second table — won revenue by close month
    # (rows 10-15) + run-rate (row 16)
    mt = months_ws
    mt.cell(row=10, column=1).value = "Won revenue by close month"
    mt.cell(row=10, column=1).font = Font(bold=True, size=12,
                                          color="1F4E79")
    mt.append([])  # row 11 header is written explicitly below
    mt.cell(row=11, column=1).value = "close_month"
    mt.cell(row=11, column=2).value = "deals_won"
    mt.cell(row=11, column=3).value = "revenue_won_pkr"
    style_header(mt, 3, row=11)
    close_months = ["2026-03", "2026-04", "2026-05", "2026-06"]
    for k, cm in enumerate(close_months):
        r = 12 + k
        mt.cell(row=r, column=1).value = cm
        mt.cell(row=r, column=2).value = (
            f'=COUNTIFS(Pipeline!$D$2:$D${last},"{cm}*",'
            f'Pipeline!$Q$2:$Q${last},1)')
        mt.cell(row=r, column=3).value = (
            f'=SUMIFS(Pipeline!$L$2:$L${last},'
            f'Pipeline!$D$2:$D${last},"{cm}*",'
            f'Pipeline!$Q$2:$Q${last},1)')
        mt.cell(row=r, column=3).number_format = PKR_FMT
    mt.cell(row=16, column=1).value = "Run-rate (avg won revenue, Apr–Jun)"
    mt.cell(row=16, column=1).font = Font(bold=True)
    mt.cell(row=16, column=3).value = "=AVERAGE(C13:C15)"
    mt.cell(row=16, column=3).number_format = PKR_FMT
    mt.cell(row=16, column=3).font = Font(bold=True)

    chart = BarChart()
    chart.type = "col"
    chart.title = "Revenue won by created month — CitySoft (SAMPLE)"
    chart.y_axis.title = "PKR"
    chart.x_axis.title = "Created month"
    data = Reference(months_ws, min_col=11, min_row=1, max_row=7)
    cats = Reference(months_ws, min_col=1, min_row=2, max_row=7)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 7.5
    chart.width = 14
    months_ws.add_chart(chart, "A18")

    # ── Stages ──────────────────────────────────────────────────────
    st = wb.create_sheet("Stages")
    st.cell(row=1, column=1).value = "Funnel — deals that reached each stage"
    st.cell(row=1, column=1).font = Font(bold=True, size=12,
                                         color="1F4E79")
    st.append(["stage_reached", "deals", "conversion_from_previous",
               "share_of_all_deals"])
    style_header(st, 4, row=2)
    funnel_rows = [("Lead", f"=COUNTA(Pipeline!$A$2:$A${last})"),
                   ("Qualified", f"=SUM(Pipeline!$R$2:$R${last})"),
                   ("Proposal", f"=SUM(Pipeline!$S$2:$S${last})"),
                   ("Negotiation", f"=SUM(Pipeline!$T$2:$T${last})"),
                   ("Won", f"=SUM(Pipeline!$Q$2:$Q${last})")]
    for k, (label, formula) in enumerate(funnel_rows):
        r = 3 + k
        st.cell(row=r, column=1).value = label
        st.cell(row=r, column=2).value = formula
        if k == 0:
            st.cell(row=r, column=3).value = 1.0
        else:
            st.cell(row=r, column=3).value = f"=IF(B{r-1}=0,\"—\",B{r}/B{r-1})"
        st.cell(row=r, column=4).value = f"=IF($B$3=0,\"—\",B{r}/$B$3)"
        st.cell(row=r, column=3).number_format = PCT_FMT
        st.cell(row=r, column=4).number_format = PCT_FMT
    st.cell(row=8, column=1).value = "Biggest leak (lowest conversion)"
    st.cell(row=8, column=1).font = Font(bold=True)
    st.cell(row=8, column=2).value = (
        "=INDEX(A3:A7,MATCH(MIN(C4:C7),C4:C7,0))&\" → \"&"
        "INDEX(A4:A7,MATCH(MIN(C4:C7),C4:C7,0))")
    # note: the row above is a label built from the funnel; the
    # conversion value itself sits in C9
    st.cell(row=9, column=1).value = "Biggest leak conversion"
    st.cell(row=9, column=2).value = "=MIN(C4:C7)"
    st.cell(row=9, column=2).number_format = PCT_FMT

    st.cell(row=11, column=1).value = "Current stage snapshot"
    st.cell(row=11, column=1).font = Font(bold=True, size=12,
                                          color="1F4E79")
    st.append(["current_stage", "deals", "share_of_deals",
               "pipeline_value_pkr", "weighted_value_pkr"])
    style_header(st, 5, row=12)
    for k, s in enumerate(STAGE_ORDER):
        r = 13 + k
        st.cell(row=r, column=1).value = s
        st.cell(row=r, column=2).value = (
            f'=COUNTIF(Pipeline!$I$2:$I${last},$A{r})')
        st.cell(row=r, column=3).value = (
            f'=IF($B$3=0,"—",B{r}/$B$3)')
        st.cell(row=r, column=4).value = (
            f'=SUMIF(Pipeline!$I$2:$I${last},$A{r},'
            f'Pipeline!$L$2:$L${last})')
        st.cell(row=r, column=5).value = (
            f'=SUMIF(Pipeline!$I$2:$I${last},$A{r},'
            f'Pipeline!$N$2:$N${last})')
        st.cell(row=r, column=3).number_format = PCT_FMT
        st.cell(row=r, column=4).number_format = PKR_FMT
        st.cell(row=r, column=5).number_format = PKR_FMT
    st.cell(row=19, column=1).value = "TOTAL"
    st.cell(row=19, column=1).font = Font(bold=True)
    for col, f in ((2, "=SUM(B13:B18)"), (4, "=SUM(D13:D18)"),
                   (5, "=SUM(E13:E18)")):
        st.cell(row=19, column=col).value = f
        st.cell(row=19, column=col).font = Font(bold=True)
    st.column_dimensions["A"].width = 26
    st.column_dimensions["B"].width = 12
    st.column_dimensions["C"].width = 24
    st.column_dimensions["D"].width = 20
    st.column_dimensions["E"].width = 20

    stage_chart = BarChart()
    stage_chart.type = "bar"
    stage_chart.title = "Deals by current stage — CitySoft (SAMPLE)"
    stage_chart.height = 7
    stage_chart.width = 14
    sdata = Reference(st, min_col=2, min_row=12, max_row=18)
    scats = Reference(st, min_col=1, min_row=13, max_row=18)
    stage_chart.add_data(sdata, titles_from_data=True)
    stage_chart.set_categories(scats)
    st.add_chart(stage_chart, "A21")

    # ── Breakdown (products + sources) ──────────────────────────────
    bd = wb.create_sheet("Breakdown")
    bd.cell(row=1, column=1).value = "By product line"
    bd.cell(row=1, column=1).font = Font(bold=True, size=12,
                                         color="1F4E79")
    hdr = ["name", "deals", "pipeline_value_pkr", "open_deals",
           "open_value_pkr", "weighted_open_pkr", "closed_deals",
           "won_deals", "lost_deals", "win_rate", "revenue_won_pkr",
           "value_lost_pkr", "avg_deal_size_pkr",
           "avg_won_deal_size_pkr", "avg_cycle_won_days",
           "avg_cycle_closed_days"]
    bd.append(hdr)
    style_header(bd, 16, row=2)
    # reuse the group formulas by copying the pattern via a helper
    # that writes into bd rows 3-7 (products, Pipeline col G) and
    # rows 10-13 (sources, Pipeline col H)
    def write_group_rows(ws, labels, key_col, start_row):
        for j, label in enumerate(labels):
            r = start_row + j
            ws.cell(row=r, column=1).value = label
            ws.cell(row=r, column=2).value = (
                f'=COUNTIF(Pipeline!${key_col}$2:${key_col}${last},'
                f'$A{r})')
            ws.cell(row=r, column=3).value = (
                f'=SUMIF(Pipeline!${key_col}$2:${key_col}${last},'
                f'$A{r},Pipeline!$L$2:$L${last})')
            ws.cell(row=r, column=4).value = (
                f'=COUNTIFS(Pipeline!${key_col}$2:${key_col}${last},'
                f'$A{r},Pipeline!$P$2:$P${last},0)')
            ws.cell(row=r, column=5).value = (
                f'=SUMIFS(Pipeline!$L$2:$L${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{r},'
                f'Pipeline!$P$2:$P${last},0)')
            ws.cell(row=r, column=6).value = (
                f'=SUMIFS(Pipeline!$N$2:$N${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{r},'
                f'Pipeline!$P$2:$P${last},0)')
            ws.cell(row=r, column=7).value = (
                f'=SUMIFS(Pipeline!$P$2:$P${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{r})')
            ws.cell(row=r, column=8).value = (
                f'=SUMIFS(Pipeline!$Q$2:$Q${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{r})')
            ws.cell(row=r, column=9).value = f"=G{r}-H{r}"
            ws.cell(row=r, column=10).value = (
                f'=IF(G{r}=0,"—",H{r}/G{r})')
            ws.cell(row=r, column=11).value = (
                f'=SUMIFS(Pipeline!$L$2:$L${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{r},'
                f'Pipeline!$Q$2:$Q${last},1)')
            ws.cell(row=r, column=12).value = (
                f'=SUMIFS(Pipeline!$L$2:$L${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{r},'
                f'Pipeline!$I$2:$I${last},"Lost")')
            ws.cell(row=r, column=13).value = (
                f'=IF(B{r}=0,"—",C{r}/B{r})')
            ws.cell(row=r, column=14).value = (
                f'=IF(H{r}=0,"—",K{r}/H{r})')
            ws.cell(row=r, column=15).value = (
                f'=IFERROR(AVERAGEIFS(Pipeline!$O$2:$O${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{r},'
                f'Pipeline!$Q$2:$Q${last},1),"—")')
            ws.cell(row=r, column=16).value = (
                f'=IFERROR(AVERAGEIFS(Pipeline!$O$2:$O${last},'
                f'Pipeline!${key_col}$2:${key_col}${last},$A{r},'
                f'Pipeline!$P$2:$P${last},1),"—")')
            for col in (3, 5, 6, 11, 12):
                ws.cell(row=r, column=col).number_format = PKR_FMT
            ws.cell(row=r, column=10).number_format = PCT_FMT
            for col in (13, 14):
                ws.cell(row=r, column=col).number_format = AVG_FMT
            for col in (15, 16):
                ws.cell(row=r, column=col).number_format = DAYS_FMT

    write_group_rows(bd, PRODUCTS, "G", 3)
    bd.cell(row=9, column=1).value = "By source"
    bd.cell(row=9, column=1).font = Font(bold=True, size=12,
                                         color="1F4E79")
    bd.append(hdr)
    style_header(bd, 16, row=10)
    write_group_rows(bd, SOURCES, "H", 11)
    bd.column_dimensions["A"].width = 22
    for col in "BCDEFGHIKLMN":
        bd.column_dimensions[col].width = 14
    bd.column_dimensions["J"].width = 10
    bd.column_dimensions["O"].width = 16
    bd.column_dimensions["P"].width = 18

    # ── Forecast ────────────────────────────────────────────────────
    fc = wb.create_sheet("Forecast")
    fc.cell(row=1, column=1).value = (
        "Forecast — open deals by expected close month (SAMPLE)")
    fc.cell(row=1, column=1).font = Font(bold=True, size=12,
                                         color="1F4E79")
    fc.append(["month", "open_deals_expected", "open_value_pkr",
               "forecast_weighted_pkr"])
    style_header(fc, 4, row=2)
    for k, m in enumerate(FORECAST_MONTHS):
        r = 3 + k
        fc.cell(row=r, column=1).value = m
        fc.cell(row=r, column=2).value = (
            f'=COUNTIFS(Pipeline!$E$2:$E${last},"{m}*",'
            f'Pipeline!$P$2:$P${last},0)')
        fc.cell(row=r, column=3).value = (
            f'=SUMIFS(Pipeline!$L$2:$L${last},'
            f'Pipeline!$E$2:$E${last},"{m}*",'
            f'Pipeline!$P$2:$P${last},0)')
        fc.cell(row=r, column=4).value = (
            f'=SUMIFS(Pipeline!$N$2:$N${last},'
            f'Pipeline!$E$2:$E${last},"{m}*",'
            f'Pipeline!$P$2:$P${last},0)')
        fc.cell(row=r, column=3).number_format = PKR_FMT
        fc.cell(row=r, column=4).number_format = PKR_FMT
    fc.cell(row=5, column=1).value = (
        "Monthly run-rate (avg won revenue, Apr–Jun 2026)")
    fc.cell(row=5, column=1).font = Font(bold=True)
    fc.cell(row=5, column=4).value = "='Monthly Trend'!C16"
    fc.cell(row=5, column=4).number_format = PKR_FMT
    fc.cell(row=5, column=4).font = Font(bold=True)
    fc.cell(row=6, column=1).value = (
        "Weighted open pipeline (all open deals)")
    fc.cell(row=6, column=1).font = Font(bold=True)
    fc.cell(row=6, column=4).value = (
        f'=SUMIFS(Pipeline!$N$2:$N${last},'
        f'Pipeline!$P$2:$P${last},0)')
    fc.cell(row=6, column=4).number_format = PKR_FMT
    fc.cell(row=6, column=4).font = Font(bold=True)
    fc.cell(row=7, column=1).value = (
        "Pipeline coverage (weighted open / run-rate)")
    fc.cell(row=7, column=1).font = Font(bold=True)
    fc.cell(row=7, column=4).value = "=IF(D5=0,\"—\",D6/D5)"
    fc.cell(row=7, column=4).number_format = "0.00"
    fc.cell(row=7, column=4).font = Font(bold=True)
    fc.column_dimensions["A"].width = 46
    fc.column_dimensions["B"].width = 20
    fc.column_dimensions["C"].width = 18
    fc.column_dimensions["D"].width = 22

    # ── Dashboard ───────────────────────────────────────────────────
    db = wb.create_sheet("Dashboard", 0)
    db.cell(row=1, column=1).value = (
        "CitySoft Solutions — Sales Pipeline / CRM Dashboard")
    db.cell(row=1, column=1).font = TITLE_FONT
    db.cell(row=2, column=1).value = (
        "SAMPLE delivery on DEMO data · Jan–Jun 2026 · "
        "156 deals · All numbers are live formulas over Pipeline")
    db.cell(row=2, column=1).font = Font(italic=True, size=11)
    db.append(["KPI", "Value", "How it is calculated"])
    style_header(db, 3, row=3)
    L = last
    kpis = [
        ("Total deals", f"=COUNTA(Pipeline!$A$2:$A${L})",
         "Count of deals on Pipeline", COUNT_FMT),
        ("Total pipeline value (PKR)",
         f"=SUM(Pipeline!$L$2:$L${L})", "Sum of deal values", PKR_FMT),
        ("Open deals", f"=COUNTIF(Pipeline!$P$2:$P${L},0)",
         "Deals not yet Won or Lost", COUNT_FMT),
        ("Open pipeline value (PKR)",
         f"=SUMIFS(Pipeline!$L$2:$L${L},Pipeline!$P$2:$P${L},0)",
         "Deal value of open deals", PKR_FMT),
        ("Weighted open pipeline (PKR)",
         f"=SUMIFS(Pipeline!$N$2:$N${L},Pipeline!$P$2:$P${L},0)",
         "Deal value x stage probability, open deals only", PKR_FMT),
        ("Closed deals", f"=SUM(Pipeline!$P$2:$P${L})",
         "Won + Lost", COUNT_FMT),
        ("Won deals", f"=SUM(Pipeline!$Q$2:$Q${L})", "Stage = Won",
         COUNT_FMT),
        ("Lost deals", "=B9-B10", "Closed − Won", COUNT_FMT),
        ("Win rate", "=IF(B9=0,\"—\",B10/B9)",
         "Won / Closed", PCT_FMT),
        ("Revenue won (PKR)",
         f"=SUMIFS(Pipeline!$L$2:$L${L},Pipeline!$Q$2:$Q${L},1)",
         "Deal value of Won deals", PKR_FMT),
        ("Value lost (PKR)",
         f"=SUMIF(Pipeline!$I$2:$I${L},\"Lost\","
         f"Pipeline!$L$2:$L${L})",
         "Deal value of Lost deals", PKR_FMT),
        ("Average deal size (PKR)", "=IF(B4=0,\"—\",B5/B4)",
         "Total pipeline / total deals", AVG_FMT),
        ("Average won deal size (PKR)", "=IF(B10=0,\"—\",B13/B10)",
         "Revenue won / won deals", AVG_FMT),
        ("Average sales cycle — won deals (days)",
         f"=AVERAGEIFS(Pipeline!$O$2:$O${L},"
         f"Pipeline!$Q$2:$Q${L},1)",
         "Close − created, Won deals only", DAYS_FMT),
        ("Average sales cycle — closed deals (days)",
         f"=AVERAGEIFS(Pipeline!$O$2:$O${L},"
         f"Pipeline!$P$2:$P${L},1)",
         "Close − created, all closed deals", DAYS_FMT),
        ("July 2026 forecast — weighted (PKR)", "=Forecast!D3",
         "Weighted value of open deals expected to close in July",
         PKR_FMT),
        ("August 2026 forecast — weighted (PKR)", "=Forecast!D4",
         "Weighted value of open deals expected to close in August",
         PKR_FMT),
        ("Monthly run-rate, Apr–Jun (PKR)", "=Forecast!D5",
         "Average won revenue per month, Apr–Jun 2026", PKR_FMT),
        ("Pipeline coverage (x)", "=Forecast!D7",
         "Weighted open pipeline / run-rate", "0.00"),
    ]
    # KPI rows start at row 4; note the cross-references above use
    # the fixed layout: B4 total deals, B5 pipeline, B9 closed,
    # B10 won, B13 revenue won.
    for k, (label, formula, note, fmt) in enumerate(kpis):
        r = 4 + k
        db.cell(row=r, column=1).value = label
        db.cell(row=r, column=2).value = formula
        db.cell(row=r, column=3).value = note
        db.cell(row=r, column=2).number_format = fmt
        db.cell(row=r, column=1).font = Font(bold=True)
    # verify the fixed-layout assumption the formulas above rely on
    assert kpis[0][0] == "Total deals" and kpis[5][0] == "Closed deals"
    assert kpis[6][0] == "Won deals" and kpis[9][0] == "Revenue won (PKR)"

    r0 = 4 + len(kpis) + 1  # first row of the top/bottom section
    db.cell(row=r0, column=1).value = "Leaders, laggards and leaks"
    db.cell(row=r0, column=1).font = Font(bold=True, size=12,
                                          color="1F4E79")
    db.cell(row=r0 + 1, column=1).value = "Top rep by revenue won"
    db.cell(row=r0 + 1, column=2).value = (
        "=INDEX(Reps!$A$2:$A$7,MATCH(MAX(Reps!$K$2:$K$7),"
        "Reps!$K$2:$K$7,0))")
    db.cell(row=r0 + 1, column=3).value = (
        "=MAX(Reps!$K$2:$K$7)")
    db.cell(row=r0 + 1, column=3).number_format = PKR_FMT
    db.cell(row=r0 + 2, column=1).value = (
        "Lowest win rate — reps with 10+ closed deals")
    db.cell(row=r0 + 2, column=2).value = (
        "=INDEX(Reps!$A$2:$A$7,MATCH("
        "_xlfn.MINIFS(Reps!$J$2:$J$7,Reps!$G$2:$G$7,\">=10\"),"
        "Reps!$J$2:$J$7,0))")
    db.cell(row=r0 + 2, column=3).value = (
        "=_xlfn.MINIFS(Reps!$J$2:$J$7,Reps!$G$2:$G$7,\">=10\")")
    db.cell(row=r0 + 2, column=3).number_format = PCT_FMT
    db.cell(row=r0 + 3, column=1).value = "Best source by win rate"
    db.cell(row=r0 + 3, column=2).value = (
        "=INDEX(Breakdown!$A$11:$A$14,MATCH("
        "_xlfn.MAXIFS(Breakdown!$J$11:$J$14,"
        "Breakdown!$G$11:$G$14,\">=10\"),"
        "Breakdown!$J$11:$J$14,0))")
    db.cell(row=r0 + 3, column=3).value = (
        "=_xlfn.MAXIFS(Breakdown!$J$11:$J$14,"
        "Breakdown!$G$11:$G$14,\">=10\")")
    db.cell(row=r0 + 3, column=3).number_format = PCT_FMT
    db.cell(row=r0 + 4, column=1).value = "Worst product line by win rate"
    db.cell(row=r0 + 4, column=2).value = (
        "=INDEX(Breakdown!$A$3:$A$7,MATCH("
        "_xlfn.MINIFS(Breakdown!$J$3:$J$7,"
        "Breakdown!$G$3:$G$7,\">=10\"),"
        "Breakdown!$J$3:$J$7,0))")
    db.cell(row=r0 + 4, column=3).value = (
        "=_xlfn.MINIFS(Breakdown!$J$3:$J$7,"
        "Breakdown!$G$3:$G$7,\">=10\")")
    db.cell(row=r0 + 4, column=3).number_format = PCT_FMT
    db.cell(row=r0 + 5, column=1).value = "Biggest funnel leak"
    db.cell(row=r0 + 5, column=2).value = "=Stages!B8"
    db.cell(row=r0 + 5, column=3).value = "=Stages!B9"
    db.cell(row=r0 + 5, column=3).number_format = PCT_FMT
    db.column_dimensions["A"].width = 46
    db.column_dimensions["B"].width = 22
    db.column_dimensions["C"].width = 52
    for r in range(r0 + 1, r0 + 6):
        db.cell(row=r, column=1).font = Font(bold=True)

    # ── Method ──────────────────────────────────────────────────────
    md = wb.create_sheet("Method")
    md.cell(row=1, column=1).value = "Method — how to read this workbook"
    md.cell(row=1, column=1).font = Font(bold=True, size=12,
                                         color="1F4E79")
    lines = [
        "SAMPLE delivery on DEMO data. CitySoft Solutions is a "
        "fictional company; the 156 deals were generated to show the "
        "method. No real company, rep or client is involved.",
        "",
        "Win rate = Won / Closed (Won + Lost). Open deals are not in "
        "the denominator — a deal still in the pipe has not been won "
        "or lost yet.",
        "",
        "Weighted pipeline = deal value x stage probability. "
        "Probabilities: Lead 10%, Qualified 25%, Proposal 50%, "
        "Negotiation 75%, Won 100%, Lost 0%. They are the team's "
        "planning weights, not measured close rates — the workbook "
        "shows the value next to every weight so the reader can see "
        "both.",
        "",
        "Funnel: a deal 'reached' a stage if its current stage is "
        "that stage or later, or it was lost at that stage or later. "
        "Won deals reached every stage. Conversion = reached this "
        "stage / reached the previous stage.",
        "",
        "Sales cycle = close date − created date, in days, for closed "
        "deals only. Open deals have no cycle yet.",
        "",
        "Forecast, two views. (1) Pipeline-based: the weighted value "
        "of open deals whose expected close date falls in the month — "
        "that is the July / August number on Dashboard and Forecast. "
        "(2) Run-rate: average monthly revenue won over Apr–Jun "
        "2026. Coverage = weighted open pipeline / run-rate: how "
        "many months of typical winning the open pipe already holds.",
        "",
        "Win rates on small bases are noisy: a rep with 9 closed "
        "deals can show a 55% win rate on five wins. The Dashboard's "
        "lowest-win-rate row only ranks reps with 10+ closed deals, "
        "and every rate in this workbook sits next to its deal count.",
        "",
        "Win rate by created month is only meaningful for mature "
        "cohorts: deals created in May–Jun 2026 mostly have not had "
        "time to close yet (June: 3 of 19 closed, 0 won so far), so "
        "their win rates read near zero. Judge the win rate on "
        "closed deals overall, not on the youngest months.",
        "",
        "Re-verify: run python3 src/checks.py — it recomputes every "
        "headline number from data/raw/deals_SAMPLE.csv, compares "
        "the workbook formulas after a LibreOffice recalc, and "
        "checks the PDF and charts exist. On a real client file, "
        "replace data/raw/deals_SAMPLE.csv with the CRM export "
        "(same columns) and skip src/generate_data.py — everything "
        "downstream runs unchanged.",
    ]
    for i, line in enumerate(lines, start=2):
        c = md.cell(row=i, column=1)
        c.value = line
        c.alignment = Alignment(wrap_text=True, vertical="top")
    md.column_dimensions["A"].width = 120

    wb.save(OUT)
    print(f"Wrote {OUT} — {wb.sheetnames}")


if __name__ == "__main__":
    main()
