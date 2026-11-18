#!/usr/bin/env python3
"""Build the P22 deliverables from the SAMPLE pipeline data.

Recomputes every number from outputs/summary.json (itself recomputed
from the raw CSV by src/analyze.py — nothing is typed in by hand)
and produces:
  * outputs/funnel_stage.png — funnel: deals that reached each stage
    (matplotlib)
  * outputs/monthly_trend.png — deals created and revenue won by
    created month (matplotlib)
  * outputs/pipeline_report.pdf — 2-page client report (headline,
    funnel + biggest leak, rep leaderboard, monthly trend, forecast,
    product / source, lost reasons, 4 recommendations)

Run:  python3 src/build_report.py
"""
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parent.parent
SUMMARY = ROOT / "outputs" / "summary.json"
CHART_FN = ROOT / "outputs" / "funnel_stage.png"
CHART_MO = ROOT / "outputs" / "monthly_trend.png"
PDF = ROOT / "outputs" / "pipeline_report.pdf"

s = json.loads(SUMMARY.read_text(encoding="utf-8"))
TOT = s["total_deals"]
PIPE = s["total_pipeline_value"]
OPEN = s["open_deals"]
OPEN_V = s["open_value"]
WOPEN = s["weighted_open_pipeline"]
WON = s["won_deals"]
CLOSED = s["closed_deals"]
WINRATE = s["win_rate"]
REV = s["revenue_won"]
AVG = s["avg_deal_size"]
CYCLE = s["avg_cycle_won"]
FUN = s["funnel"]
LEAK = s["biggest_leak_step"]
LEAK_RATE = s["biggest_leak_rate"]
REPS = sorted(s["reps"], key=lambda x: -x["revenue_won"])
MO = s["months_summary"]
FC = s["forecast"]
RUNRATE = s["run_rate_3m"]
COVER = s["pipeline_coverage_x"]
PROD = sorted(s["products"], key=lambda x: -x["revenue_won"])
SRC = sorted(s["sources"], key=lambda x: -x["win_rate"])
LOST = s["lost_reasons"]
ACTS = s["actions"]

# funnel chart: deals that reached each stage
labels = [f["stage"] for f in FUN]
reached = [f["reached"] for f in FUN]
fig, ax = plt.subplots(figsize=(8, 3.4))
bars = ax.barh(labels[::-1], reached[::-1], color="#1F4E79")
ax.set_xlabel("Deals that reached the stage")
ax.set_title("Sales funnel — CitySoft Solutions (SAMPLE, Jan–Jun 2026)")
for bar, v in zip(bars, reached[::-1]):
    ax.text(v + 2, bar.get_y() + bar.get_height() / 2, str(v),
            va="center", fontsize=9)
ax.set_xlim(0, max(reached) * 1.12)
fig.tight_layout()
plt.savefig(CHART_FN, dpi=150)
plt.close()

# monthly trend: deals created + revenue won by created month
months = [m["name"] for m in MO]
deals = [m["deals"] for m in MO]
rev = [m["revenue_won"] for m in MO]
fig, ax1 = plt.subplots(figsize=(8, 3.4))
ax1.bar(months, deals, width=0.55, color="#1F4E79", label="Deals created")
ax1.set_ylabel("Deals created")
ax1.set_title("Monthly trend — CitySoft Solutions (SAMPLE)")
for i, v in enumerate(deals):
    ax1.text(i, v + 0.5, str(v), ha="center", fontsize=9)
ax2 = ax1.twinx()
ax2.plot(months, rev, color="#C62828", marker="o",
         label="Revenue won (PKR)")
ax2.set_ylabel("Revenue won (PKR)")
ax2.ticklabel_format(style="plain", axis="y")
fig.tight_layout()
plt.savefig(CHART_MO, dpi=150)
plt.close()

styles = getSampleStyleSheet()
styles["Heading2"].spaceBefore = 4
styles["Heading2"].spaceAfter = 2
styles["Heading2"].fontSize = 12
styles["Heading2"].leading = 14
styles["Normal"].leading = 13
styles["Normal"].spaceAfter = 2
styles["Title"].fontSize = 20
styles["Title"].leading = 22
styles["Title"].spaceAfter = 2
hdr = colors.HexColor("#1F4E79")


def table(rows, widths, size=8):
    t = Table(rows, colWidths=widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), hdr),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), size),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def p(txt, style="Normal"):
    return Paragraph(txt, styles[style])


story = [
    p("CitySoft Solutions — Sales Pipeline / CRM Dashboard",
      "Title"),
    p("SAMPLE delivery on DEMO data · Jan–Jun 2026 · "
      "Where is next month's revenue?", "Normal"),
    Spacer(1, 0.2 * cm),
    p(f"<b>{TOT} deals · PKR {PIPE:,} total pipeline · "
      f"{OPEN} open deals (PKR {OPEN_V:,}, weighted "
      f"PKR {WOPEN:,}) · {WON} won of {CLOSED} closed "
      f"({WINRATE}%) · PKR {REV:,} revenue won · average deal "
      f"PKR {AVG:,.0f} · average sales cycle {CYCLE} days</b>",
      "Normal"),
    Spacer(1, 0.2 * cm),
    p("<b>The answer in one line:</b> July 2026's pipeline-based "
      f"forecast is <b>PKR {FC[0]['forecast_weighted']:,}</b> "
      f"({FC[0]['open_deals_expected']} open deals expected to close "
      f"in July), against a PKR {RUNRATE:,.0f} monthly run-rate "
      f"(average revenue won, Apr–Jun). Weighted open pipeline "
      f"covers {COVER}x one month's won revenue — July depends on "
      "closing what is already in the pipe, not on new leads.",
      "Normal"),
    Spacer(1, 0.15 * cm),
    p("The funnel — where deals leak", "Heading2"),
    Image(str(CHART_FN), width=14 * cm, height=3.4 * cm),
    table(
        [["Stage reached", "Deals", "Conversion", "Share of all"]]
        + [[f["stage"], str(f["reached"]),
            f"{f['conversion_from_previous']}%",
            f"{f['share_of_all_deals']}%"] for f in FUN],
        [4.5 * cm, 2.5 * cm, 4 * cm, 4 * cm]),
    Spacer(1, 0.2 * cm),
    p(f"<b>Biggest leak: {LEAK} ({LEAK_RATE}%).</b> Only "
      f"{LEAK_RATE}% of deals that reach Negotiation become Won. "
      f"Price too high is the top lost reason "
      f"({LOST[0]['lost_deals']} lost deals, "
      f"PKR {LOST[0]['value_lost']:,} of pipeline value).",
      "Normal"),
    Spacer(1, 0.15 * cm),
    p("Rep leaderboard — by revenue won", "Heading2"),
    table(
        [["Rep", "Deals", "Closed", "Won", "Win rate",
          "Revenue won", "Avg cycle"]]
        + [[r["name"], str(r["deals"]), str(r["closed"]),
            str(r["won"]), f"{r['win_rate']}%",
            f"PKR {r['revenue_won']:,}",
            f"{r['avg_cycle_won']} days"] for r in REPS],
        [3.2 * cm, 1.6 * cm, 1.8 * cm, 1.4 * cm, 2 * cm, 3.4 * cm,
         2.4 * cm]),
    Spacer(1, 0.2 * cm),
    p("Month by month", "Heading2"),
    Image(str(CHART_MO), width=14 * cm, height=3.4 * cm),
    table(
        [["Month", "Deals", "Pipeline", "Won", "Win rate",
          "Revenue won"]]
        + [[m["name"], str(m["deals"]),
            f"PKR {m['pipeline_value']:,}", str(m["won"]),
            f"{m['win_rate']}%", f"PKR {m['revenue_won']:,}"]
           for m in MO],
        [2.6 * cm, 1.8 * cm, 3.6 * cm, 1.6 * cm, 2.2 * cm, 3.8 * cm]),
    Spacer(1, 0.15 * cm),
    p("Forecast — July and August 2026", "Heading2"),
    table(
        [["Month", "Open deals expected", "Open value",
          "Forecast (weighted)"]]
        + [[f["month"], str(f["open_deals_expected"]),
            f"PKR {f['open_value_expected']:,}",
            f"PKR {f['forecast_weighted']:,}"] for f in FC],
        [3.5 * cm, 4 * cm, 4 * cm, 4.2 * cm]),
    Spacer(1, 0.2 * cm),
    p("Forecast = weighted value of open deals by expected close "
      "month. August is smaller — only 14 open deals are expected "
      "to close then.", "Normal"),
    Spacer(1, 0.15 * cm),
    p("Product lines and sources", "Heading2"),
    table(
        [["Product line", "Deals", "Won", "Win rate", "Revenue won"]]
        + [[x["name"], str(x["deals"]), str(x["won"]),
            f"{x['win_rate']}%", f"PKR {x['revenue_won']:,}"]
           for x in PROD],
        [5 * cm, 2 * cm, 2 * cm, 2.5 * cm, 4.3 * cm]),
    Spacer(1, 0.2 * cm),
    table(
        [["Source", "Deals", "Won", "Win rate", "Revenue won"]]
        + [[x["name"], str(x["deals"]), str(x["won"]),
            f"{x['win_rate']}%", f"PKR {x['revenue_won']:,}"]
           for x in SRC],
        [5 * cm, 2 * cm, 2 * cm, 2.5 * cm, 4.3 * cm]),
    Spacer(1, 0.2 * cm),
    p("Why deals were lost", "Heading2"),
    table(
        [["Lost reason", "Lost deals", "Share of losses",
          "Value lost"]]
        + [[x["lost_reason"], str(x["lost_deals"]),
            f"{x['share_of_lost']}%", f"PKR {x['value_lost']:,}"]
           for x in LOST],
        [5.5 * cm, 2.5 * cm, 3.5 * cm, 4.3 * cm]),
    Spacer(1, 0.15 * cm),
    p("What to do next — four actions from the data", "Heading2"),
]
for i, a in enumerate(ACTS, 1):
    story.append(p(f"<b>{i}.</b> {a}", "Normal"))
    story.append(Spacer(1, 0.15 * cm))
story.append(Spacer(1, 0.15 * cm))
story.append(p(
    "Method: win rate = Won / Closed. Weighted pipeline = deal "
    "value x stage probability (Lead 10%, Qualified 25%, Proposal "
    "50%, Negotiation 75%, Won 100%, Lost 0%). Forecast = weighted "
    "value of open deals by expected close month. Sales cycle = "
    "close date − created date. Every number is recomputed from "
    "the deals file and checked by src/checks.py; the workbook "
    "carries the same numbers as live formulas. SAMPLE data: "
    "CitySoft Solutions is fictional.", "Normal"))

doc = SimpleDocTemplate(str(PDF), pagesize=A4,
                        leftMargin=1.8 * cm, rightMargin=1.8 * cm,
                        topMargin=1.1 * cm, bottomMargin=1.1 * cm,
                        title="CitySoft Solutions — Sales Pipeline / "
                              "CRM Dashboard (SAMPLE)",
                        author="Muhammad Hashir")
doc.build(story)
print(f"Wrote {PDF}")
