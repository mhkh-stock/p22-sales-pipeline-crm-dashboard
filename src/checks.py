#!/usr/bin/env python3
"""Checks as code for P22. Recomputes every headline number from the
raw CSV and verifies the JSON, the processed CSVs, the markdown
report, the workbook formulas and the LibreOffice recalc against it
— three independent paths that must agree. Win rate = won / closed;
weighted value = deal value x stage probability (Lead 10, Qualified
25, Proposal 50, Negotiation 75, Won 100, Lost 0); forecast =
weighted value of open deals by expected close month; sales cycle =
close date - created date for closed deals; every group total (rep,
month, product, source, lost reason, stage) must sum to the same
headline numbers.

Runs the LibreOffice recalc itself on every run, so the check
count is stable run to run.

    python3 src/checks.py

Exit code 0 when every check passes, 1 otherwise.
"""
import csv
import json
import subprocess
import sys
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"
PROC = PROJECT / "data" / "processed"
OUT = PROJECT / "outputs"
SUMMARY = OUT / "summary.json"
WORKBOOK = OUT / "pipeline_report_SAMPLE.xlsx"
RECALC_DIR = OUT / ".recalc"
RECALC = RECALC_DIR / "pipeline_report_SAMPLE.xlsx"
PDF = OUT / "pipeline_report.pdf"
MD = OUT / "pipeline_report.md"
README = PROJECT / "README.md"
PACKAGE = PROJECT / "PACKAGE-NOTE.md"
CHART_FN = OUT / "funnel_stage.png"
CHART_MO = OUT / "monthly_trend.png"

STAGES = ["Lead", "Qualified", "Proposal", "Negotiation"]
STAGE_ORDER = ["Lead", "Qualified", "Proposal", "Negotiation",
               "Won", "Lost"]
PROB = {"Lead": 10, "Qualified": 25, "Proposal": 50,
        "Negotiation": 75, "Won": 100, "Lost": 0}
REPS = ["Ayesha Khan", "Bilal Ahmed", "Sana Malik", "Usman Tariq",
        "Hira Shah", "Daniyal Raza"]
PRODUCTS = ["ERP Implementation", "Cloud Migration", "Mobile App",
            "Website Development", "Digital Marketing"]
SOURCES = ["Referral", "Website", "LinkedIn", "Cold Outreach"]
MONTHS = [f"2026-{m:02d}" for m in range(1, 7)]
LOST_REASONS = ["Price too high", "Chose competitor",
                "Budget frozen", "No decision — went quiet",
                "Requirements changed", "Timing — not now"]
fails = []
ran = []


def check(name, cond, detail=""):
    ran.append(name)
    if not cond:
        print("FAIL " + name + (f" — {detail}" if detail else ""))
        fails.append(name)


def close(a, b, tol=0.01):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def pct(n, d):
    return round(100.0 * n / d, 2) if d else 0.0


def group_stats(rows):
    n = len(rows)
    won = [r for r in rows if r["is_won"]]
    closed = [r for r in rows if r["is_closed"]]
    lost = [r for r in rows if r["is_closed"] and not r["is_won"]]
    open_rows = [r for r in rows if not r["is_closed"]]
    cw = [r["sales_cycle_days"] for r in won
          if r["sales_cycle_days"] is not None]
    cc = [r["sales_cycle_days"] for r in closed
          if r["sales_cycle_days"] is not None]
    return {
        "deals": n,
        "pipeline_value": sum(r["deal_value_pkr"] for r in rows),
        "open_deals": len(open_rows),
        "open_value": sum(r["deal_value_pkr"] for r in open_rows),
        "weighted_open": sum(r["weighted_value_pkr"] for r in open_rows),
        "closed": len(closed),
        "won": len(won),
        "lost": len(lost),
        "win_rate": pct(len(won), len(closed)),
        "revenue_won": sum(r["deal_value_pkr"] for r in won),
        "value_lost": sum(r["deal_value_pkr"] for r in lost),
        "avg_deal_size": round(
            sum(r["deal_value_pkr"] for r in rows) / n, 2) if n else 0.0,
        "avg_won_deal_size": round(
            sum(r["deal_value_pkr"] for r in won) / len(won), 2)
        if won else 0.0,
        "avg_cycle_won": round(sum(cw) / len(cw), 1) if cw else 0.0,
        "avg_cycle_closed": round(sum(cc) / len(cc), 1) if cc else 0.0,
    }


# ── raw file ────────────────────────────────────────────────────────
raw = list(csv.DictReader(open(RAW / "deals_SAMPLE.csv",
                               encoding="utf-8")))
check("raw: 156 deals", len(raw) == 156, f"got {len(raw)}")
check("raw: columns exact",
      list(raw[0].keys()) == [
          "deal_id", "created_date", "month", "close_date",
          "expected_close_date", "rep", "product_line", "source",
          "stage", "lost_at_stage", "lost_reason", "deal_value_pkr",
          "probability_pct", "weighted_value_pkr",
          "sales_cycle_days", "is_closed", "is_won"],
      str(list(raw[0].keys())))
ids = [r["deal_id"] for r in raw]
check("raw: deal ids unique", len(set(ids)) == len(ids))
check("raw: deal ids DL-0001..DL-0156",
      ids == [f"DL-{i:04d}" for i in range(1, 157)])

rows = []
for r in raw:
    rows.append({
        "deal_id": r["deal_id"],
        "created_date": r["created_date"],
        "month": r["month"],
        "close_date": r["close_date"] or None,
        "expected_close_date": r["expected_close_date"] or None,
        "rep": r["rep"],
        "product_line": r["product_line"],
        "source": r["source"],
        "stage": r["stage"],
        "lost_at_stage": r["lost_at_stage"] or None,
        "lost_reason": r["lost_reason"] or None,
        "deal_value_pkr": int(r["deal_value_pkr"]),
        "probability_pct": int(r["probability_pct"]),
        "weighted_value_pkr": int(r["weighted_value_pkr"]),
        "sales_cycle_days": (int(r["sales_cycle_days"])
                             if r["sales_cycle_days"] else None),
        "is_closed": int(r["is_closed"]),
        "is_won": int(r["is_won"]),
    })

for r in rows:
    d = r["deal_id"]
    check(f"row {d}: stage valid", r["stage"] in STAGE_ORDER)
    check(f"row {d}: rep known", r["rep"] in REPS)
    check(f"row {d}: product known", r["product_line"] in PRODUCTS)
    check(f"row {d}: source known", r["source"] in SOURCES)
    check(f"row {d}: month matches created date",
          r["month"] == r["created_date"][:7])
    check(f"row {d}: created in Jan-Jun 2026",
          "2026-01-01" <= r["created_date"] <= "2026-06-30")
    check(f"row {d}: value positive", r["deal_value_pkr"] > 0)
    check(f"row {d}: probability matches stage",
          r["probability_pct"] == PROB[r["stage"]],
          f"{r['probability_pct']} vs {PROB[r['stage']]}")
    check(f"row {d}: weighted = round(value x prob / 100)",
          r["weighted_value_pkr"] == round(
              r["deal_value_pkr"] * r["probability_pct"] / 100))
    if r["stage"] == "Won":
        check(f"row {d}: won implies closed+won",
              r["is_closed"] == 1 and r["is_won"] == 1)
        check(f"row {d}: won has close date, no expected close",
              r["close_date"] and not r["expected_close_date"])
    elif r["stage"] == "Lost":
        check(f"row {d}: lost implies closed, not won",
              r["is_closed"] == 1 and r["is_won"] == 0)
        check(f"row {d}: lost has lost_at_stage in stages",
              r["lost_at_stage"] in STAGES)
        check(f"row {d}: lost has lost reason",
              r["lost_reason"] in LOST_REASONS)
        check(f"row {d}: lost has close date, no expected close",
              r["close_date"] and not r["expected_close_date"])
    else:
        check(f"row {d}: open deal not closed/won",
              r["is_closed"] == 0 and r["is_won"] == 0)
        check(f"row {d}: open has expected close, no close date",
              r["expected_close_date"] and not r["close_date"])
        check(f"row {d}: open has no lost fields",
              not r["lost_at_stage"] and not r["lost_reason"])
    if r["is_closed"]:
        cd = date.fromisoformat(r["close_date"])
        cr = date.fromisoformat(r["created_date"])
        check(f"row {d}: cycle = close - created",
              r["sales_cycle_days"] == (cd - cr).days)
        check(f"row {d}: close on/after created", cd >= cr)
    else:
        check(f"row {d}: open has no cycle",
              r["sales_cycle_days"] is None)

total = group_stats(rows)

# ── funnel (recomputed independently) ───────────────────────────────
def reached(stage):
    idx = STAGES.index(stage)
    out = []
    for r in rows:
        cur = r["stage"]
        if cur == "Won":
            out.append(r)
        elif cur in STAGES and STAGES.index(cur) >= idx:
            out.append(r)
        elif cur == "Lost" and r["lost_at_stage"] in STAGES \
                and STAGES.index(r["lost_at_stage"]) >= idx:
            out.append(r)
    return out


funnel = [("Lead", len(rows))]
for s_ in STAGES[1:]:
    funnel.append((s_, len(reached(s_))))
funnel.append(("Won", total["won"]))
convs = [pct(funnel[i][1], funnel[i - 1][1])
         for i in range(1, len(funnel))]
leak_i = convs.index(min(convs)) + 1
leak_step = f"{funnel[leak_i - 1][0]} → {funnel[leak_i][0]}"

# ── summary.json ────────────────────────────────────────────────────
s = json.loads(SUMMARY.read_text(encoding="utf-8"))
check("json: total deals", s["total_deals"] == total["deals"])
check("json: total pipeline",
      s["total_pipeline_value"] == total["pipeline_value"])
check("json: open deals", s["open_deals"] == total["open_deals"])
check("json: open value", s["open_value"] == total["open_value"])
check("json: weighted open",
      s["weighted_open_pipeline"] == total["weighted_open"])
check("json: closed", s["closed_deals"] == total["closed"])
check("json: won", s["won_deals"] == total["won"])
check("json: lost", s["lost_deals"] == total["lost"])
check("json: win rate", close(s["win_rate"], total["win_rate"], 0.01))
check("json: revenue won", s["revenue_won"] == total["revenue_won"])
check("json: value lost", s["value_lost"] == total["value_lost"])
check("json: avg deal size",
      close(s["avg_deal_size"], total["avg_deal_size"]))
check("json: avg won deal",
      close(s["avg_won_deal_size"], total["avg_won_deal_size"]))
check("json: avg cycle won",
      close(s["avg_cycle_won"], total["avg_cycle_won"], 0.05))
check("json: avg cycle closed",
      close(s["avg_cycle_closed"], total["avg_cycle_closed"], 0.05))
check("json: funnel stages",
      [f["stage"] for f in s["funnel"]] ==
      [f[0] for f in funnel])
check("json: funnel reached",
      [f["reached"] for f in s["funnel"]] ==
      [f[1] for f in funnel])
check("json: funnel conversions",
      [f["conversion_from_previous"] for f in s["funnel"]] ==
      [100.0] + convs)
check("json: biggest leak step", s["biggest_leak_step"] == leak_step,
      s["biggest_leak_step"])
check("json: biggest leak rate",
      close(s["biggest_leak_rate"], min(convs), 0.01))
check("json: funnel monotonic non-increasing",
      all(funnel[i][1] <= funnel[i - 1][1]
          for i in range(1, len(funnel))))
check("json: won + lost = closed",
      s["won_deals"] + s["lost_deals"] == s["closed_deals"])
check("json: open + closed = total",
      s["open_deals"] + s["closed_deals"] == s["total_deals"])
check("json: revenue won + value lost <= pipeline",
      s["revenue_won"] + s["value_lost"] <= s["total_pipeline_value"])
check("json: weighted open <= open value",
      s["weighted_open_pipeline"] <= s["open_value"])
check("json: 4 actions", len(s["actions"]) == 4)
check("json: actions name the leak",
      leak_step in s["actions"][0], s["actions"][0][:80])
check("json: actions name the top lost reason",
      s["lost_reasons"][0]["lost_reason"] in s["actions"][2])

# forecast
last3 = [x for x in s["won_by_close_month"]
         if x["month"] in ["2026-04", "2026-05", "2026-06"]]
run_rate = round(sum(x["revenue_won"] for x in last3) / 3.0, 2)
check("json: run-rate = avg Apr-Jun won revenue",
      close(s["run_rate_3m"], run_rate), f"{s['run_rate_3m']} vs {run_rate}")
check("json: coverage = weighted open / run-rate",
      close(s["pipeline_coverage_x"],
            round(total["weighted_open"] / run_rate, 2), 0.01))
for f in s["forecast"]:
    g = [r for r in rows if not r["is_closed"]
         and r["expected_close_date"]
         and r["expected_close_date"][:7] == f["month"]]
    check(f"json: forecast {f['month']} deals",
          f["open_deals_expected"] == len(g))
    check(f"json: forecast {f['month']} weighted",
          f["forecast_weighted"] ==
          sum(r["weighted_value_pkr"] for r in g))
    check(f"json: forecast {f['month']} open value",
          f["open_value_expected"] ==
          sum(r["deal_value_pkr"] for r in g))
check("json: July forecast = headline",
      s["july_forecast_weighted"] == s["forecast"][0]["forecast_weighted"])

# group sums reconcile to totals
def check_groups(key, values, label, json_key):
    by_json = {g["name"]: g for g in s[json_key]}
    check(f"json {label}: groups listed",
          set(by_json) == set(values))
    for v in values:
        g = [r for r in rows if r[key] == v]
        gs = group_stats(g)
        j = by_json[v]
        check(f"json {label} {v}: deals", j["deals"] == gs["deals"])
        check(f"json {label} {v}: pipeline",
              j["pipeline_value"] == gs["pipeline_value"])
        check(f"json {label} {v}: won", j["won"] == gs["won"])
        check(f"json {label} {v}: win rate",
              close(j["win_rate"], gs["win_rate"], 0.01))
        check(f"json {label} {v}: revenue won",
              j["revenue_won"] == gs["revenue_won"])
        check(f"json {label} {v}: weighted open",
              j["weighted_open"] == gs["weighted_open"])
    for field in ("deals", "pipeline_value", "closed", "won", "lost",
                  "revenue_won", "value_lost", "open_value",
                  "weighted_open"):
        tot = sum(by_json[v][field] for v in values)
        exp = total[{"deals": "deals", "pipeline_value": "pipeline_value",
                     "closed": "closed", "won": "won", "lost": "lost",
                     "revenue_won": "revenue_won",
                     "value_lost": "value_lost",
                     "open_value": "open_value",
                     "weighted_open": "weighted_open"}[field]]
        check(f"json {label}: sum {field} = total", tot == exp,
              f"{tot} vs {exp}")


check_groups("rep", REPS, "reps", "reps")
check_groups("month", MONTHS, "months", "months_summary")
check_groups("product_line", PRODUCTS, "products", "products")
check_groups("source", SOURCES, "sources", "sources")

# stage snapshot sums
st_sum = {x["stage"]: x for x in s["stage_summary"]}
check("json stages: 6 stages", set(st_sum) == set(STAGE_ORDER))
check("json stages: deals sum = total",
      sum(x["deals"] for x in s["stage_summary"]) == total["deals"])
check("json stages: pipeline sum = total",
      sum(x["pipeline_value"] for x in s["stage_summary"]) ==
      total["pipeline_value"])
check("json stages: weighted sum = total weighted",
      sum(x["weighted_value"] for x in s["stage_summary"]) ==
      sum(r["weighted_value_pkr"] for r in rows))

# lost reasons
lr_sum = sum(x["lost_deals"] for x in s["lost_reasons"])
check("json lost reasons: deals sum = lost",
      lr_sum == total["lost"], f"{lr_sum} vs {total['lost']}")
check("json lost reasons: value sum = value lost",
      sum(x["value_lost"] for x in s["lost_reasons"]) ==
      total["value_lost"])
check("json lost reasons: ranked desc",
      [x["lost_deals"] for x in s["lost_reasons"]] ==
      sorted([x["lost_deals"] for x in s["lost_reasons"]],
             reverse=True))
check("json lost_at_stage: sums to lost",
      sum(s["lost_at_stage"].values()) == total["lost"])

# ── processed CSVs ──────────────────────────────────────────────────
def read_csv(path):
    return list(csv.DictReader(open(path, encoding="utf-8")))


enr = read_csv(PROC / "deals_enriched.csv")
check("processed: enriched 156 rows", len(enr) == 156)
check("processed: enriched weighted sums",
      sum(int(r["weighted_value_pkr"]) for r in enr) ==
      sum(r["weighted_value_pkr"] for r in rows))
for name, key, values in (
        ("rep_summary.csv", "name", REPS),
        ("monthly_summary.csv", "name", MONTHS),
        ("product_summary.csv", "name", PRODUCTS),
        ("source_summary.csv", "name", SOURCES)):
    df = read_csv(PROC / name)
    check(f"processed: {name} rows", len(df) == len(values),
          f"{len(df)} vs {len(values)}")
    check(f"processed: {name} deals sum = total",
          sum(int(r["deals"]) for r in df) == total["deals"])
    check(f"processed: {name} revenue won sum = total",
          sum(int(r["revenue_won"]) for r in df) ==
          total["revenue_won"])
stage_df = read_csv(PROC / "stage_summary.csv")
check("processed: stage_summary 6 rows", len(stage_df) == 6)
lr_df = read_csv(PROC / "lost_reason_summary.csv")
check("processed: lost_reason_summary 6 rows", len(lr_df) == 6)
check("processed: lost_reason deals sum = lost",
      sum(int(r["lost_deals"]) for r in lr_df) == total["lost"])

# ── markdown report ─────────────────────────────────────────────────
md = MD.read_text(encoding="utf-8")
for needle in (f"{total['deals']} deals",
               f"PKR {total['pipeline_value']:,}",
               f"PKR {total['revenue_won']:,}",
               f"{total['win_rate']}%",
               leak_step,
               s["forecast"][0]["month"]):
    check(f"md: contains {needle!r}", needle in md)

# ── workbook recalc ─────────────────────────────────────────────────
check("workbook exists", WORKBOOK.exists())
# always re-run the recalc, so the check count is stable run to run
RECALC_DIR.mkdir(parents=True, exist_ok=True)
res = subprocess.run(
    ["soffice", "--headless", "--calc", "--convert-to", "xlsx",
     "--outdir", str(RECALC_DIR), str(WORKBOOK)],
    capture_output=True, text=True, timeout=180)
check("recalc: soffice ran", res.returncode == 0,
      res.stderr[-300:] if res.stderr else "")
check("recalc file exists", RECALC.exists())
wb = load_workbook(RECALC, data_only=True)
check("workbook: 8 sheets", wb.sheetnames == [
    "Dashboard", "Pipeline", "Reps", "Monthly Trend", "Stages",
    "Breakdown", "Forecast", "Method"], str(wb.sheetnames))

# no error values anywhere in the recalc
errs = []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("#"):
                errs.append(f"{ws.title}!{c.coordinate}={c.value}")
check("recalc: no error values", not errs, "; ".join(errs[:5]))

db = wb["Dashboard"]
dash_expected = {
    4: total["deals"], 5: total["pipeline_value"],
    6: total["open_deals"], 7: total["open_value"],
    8: total["weighted_open"], 9: total["closed"],
    10: total["won"], 11: total["lost"],
    13: total["revenue_won"], 14: total["value_lost"],
    19: s["forecast"][0]["forecast_weighted"],
    20: s["forecast"][1]["forecast_weighted"],
}
for r, exp in dash_expected.items():
    check(f"recalc Dashboard B{r}",
          db.cell(row=r, column=2).value == exp,
          f"{db.cell(row=r, column=2).value} vs {exp}")
check("recalc Dashboard win rate",
      close(db.cell(row=12, column=2).value,
            total["won"] / total["closed"], 0.0001))
check("recalc Dashboard avg deal size",
      close(db.cell(row=15, column=2).value,
            total["avg_deal_size"], 0.01))
check("recalc Dashboard avg won deal",
      close(db.cell(row=16, column=2).value,
            total["avg_won_deal_size"], 0.01))
check("recalc Dashboard avg cycle won",
      close(db.cell(row=17, column=2).value,
            total["avg_cycle_won"], 0.05))
check("recalc Dashboard avg cycle closed",
      close(db.cell(row=18, column=2).value,
            total["avg_cycle_closed"], 0.05))
check("recalc Dashboard run-rate",
      close(db.cell(row=21, column=2).value, run_rate, 0.01))
check("recalc Dashboard coverage",
      close(db.cell(row=22, column=2).value,
            round(total["weighted_open"] / run_rate, 2), 0.01))
check("recalc Dashboard top rep",
      db.cell(row=25, column=2).value == s["best_rep"],
      str(db.cell(row=25, column=2).value))
check("recalc Dashboard top rep value",
      db.cell(row=25, column=3).value ==
      max(r["revenue_won"] for r in
          [group_stats([x for x in rows if x["rep"] == rep])
           for rep in REPS]))
check("recalc Dashboard worst rep",
      db.cell(row=26, column=2).value == s["worst_rep"],
      str(db.cell(row=26, column=2).value))
check("recalc Dashboard worst rep rate",
      close(db.cell(row=26, column=3).value,
            min(r["win_rate"] for r in s["reps"]
                if r["closed"] >= 10) / 100, 0.001))
check("recalc Dashboard best source",
      db.cell(row=27, column=2).value ==
      max(s["sources"], key=lambda x: x["win_rate"])["name"])
check("recalc Dashboard worst product",
      db.cell(row=28, column=2).value ==
      min(s["products"], key=lambda x: x["win_rate"])["name"])
check("recalc Dashboard biggest leak",
      db.cell(row=29, column=2).value == leak_step,
      str(db.cell(row=29, column=2).value))
check("recalc Dashboard leak rate",
      close(db.cell(row=29, column=3).value, min(convs) / 100, 0.001))

# Pipeline sheet formulas recalculated
pl = wb["Pipeline"]
check("recalc Pipeline rows", pl.max_row == 157,
      f"max_row {pl.max_row}")
bad_weighted = 0
bad_closed = 0
bad_won = 0
for i, r in enumerate(rows, start=2):
    if pl.cell(row=i, column=14).value != r["weighted_value_pkr"]:
        bad_weighted += 1
    if pl.cell(row=i, column=16).value != r["is_closed"]:
        bad_closed += 1
    if pl.cell(row=i, column=17).value != r["is_won"]:
        bad_won += 1
check("recalc Pipeline weighted values all match", bad_weighted == 0,
      f"{bad_weighted} mismatches")
check("recalc Pipeline is_closed all match", bad_closed == 0)
check("recalc Pipeline is_won all match", bad_won == 0)
for j, s_ in enumerate(STAGES[1:], start=0):
    col = 18 + j
    exp = len(reached(s_))
    got = sum(pl.cell(row=i, column=col).value or 0
              for i in range(2, 158))
    check(f"recalc Pipeline reached_{s_.lower()} sum",
          got == exp, f"{got} vs {exp}")

# Stages funnel
st = wb["Stages"]
for k, (label, n_) in enumerate(funnel):
    r = 3 + k
    check(f"recalc Stages funnel {label} reached",
          st.cell(row=r, column=2).value == n_)
    if k:
        check(f"recalc Stages funnel {label} conversion",
              close(st.cell(row=r, column=3).value,
                    convs[k - 1] / 100, 0.001))
check("recalc Stages biggest leak label",
      st.cell(row=8, column=2).value == leak_step,
      str(st.cell(row=8, column=2).value))
check("recalc Stages snapshot total deals",
      st.cell(row=19, column=2).value == total["deals"])
check("recalc Stages snapshot total pipeline",
      st.cell(row=19, column=4).value == total["pipeline_value"])
for k, s_ in enumerate(STAGE_ORDER):
    r = 13 + k
    g = [x for x in rows if x["stage"] == s_]
    check(f"recalc Stages snapshot {s_} deals",
          st.cell(row=r, column=2).value == len(g))
    check(f"recalc Stages snapshot {s_} pipeline",
          st.cell(row=r, column=4).value ==
          sum(x["deal_value_pkr"] for x in g))

# Reps / Monthly / Breakdown / Forecast sheets
def check_group_sheet(ws, labels, json_key):
    by_json = {g["name"]: g for g in s[json_key]}
    for j, label in enumerate(labels):
        r = j + 2
        gj = by_json[label]
        checks_ = [
            ("deals", 2, gj["deals"]),
            ("pipeline", 3, gj["pipeline_value"]),
            ("open", 4, gj["open_deals"]),
            ("closed", 7, gj["closed"]),
            ("won", 8, gj["won"]),
            ("revenue_won", 11, gj["revenue_won"]),
            ("weighted_open", 6, gj["weighted_open"]),
        ]
        for name, col, exp in checks_:
            check(f"recalc {ws.title} {label} {name}",
                  ws.cell(row=r, column=col).value == exp,
                  f"{ws.cell(row=r, column=col).value} vs {exp}")
        check(f"recalc {ws.title} {label} win rate",
              close(ws.cell(row=r, column=10).value,
                    gj["win_rate"] / 100, 0.001))


check_group_sheet(wb["Reps"], REPS, "reps")
check_group_sheet(wb["Monthly Trend"], MONTHS, "months_summary")
bd = wb["Breakdown"]
for j, label in enumerate(PRODUCTS):
    r = 3 + j
    gj = {g["name"]: g for g in s["products"]}[label]
    check(f"recalc Breakdown product {label} deals",
          bd.cell(row=r, column=2).value == gj["deals"])
    check(f"recalc Breakdown product {label} revenue won",
          bd.cell(row=r, column=11).value == gj["revenue_won"])
for j, label in enumerate(SOURCES):
    r = 11 + j
    gj = {g["name"]: g for g in s["sources"]}[label]
    check(f"recalc Breakdown source {label} deals",
          bd.cell(row=r, column=2).value == gj["deals"])
    check(f"recalc Breakdown source {label} win rate",
          close(bd.cell(row=r, column=10).value,
                gj["win_rate"] / 100, 0.001))
fc = wb["Forecast"]
for k, f in enumerate(s["forecast"]):
    r = 3 + k
    check(f"recalc Forecast {f['month']} deals",
          fc.cell(row=r, column=2).value == f["open_deals_expected"])
    check(f"recalc Forecast {f['month']} weighted",
          fc.cell(row=r, column=4).value == f["forecast_weighted"])
check("recalc Forecast run-rate",
      close(fc.cell(row=5, column=4).value, run_rate, 0.01))
check("recalc Forecast weighted open",
      fc.cell(row=6, column=4).value == total["weighted_open"])
check("recalc Forecast coverage",
      close(fc.cell(row=7, column=4).value,
            round(total["weighted_open"] / run_rate, 2), 0.01))
mt = wb["Monthly Trend"]
check("recalc Monthly run-rate",
      close(mt.cell(row=16, column=3).value, run_rate, 0.01))
for k, m in enumerate(s["won_by_close_month"]):
    if m["month"] in ("2026-03", "2026-04", "2026-05", "2026-06"):
        r = 12 + ["2026-03", "2026-04", "2026-05",
                  "2026-06"].index(m["month"])
        check(f"recalc Monthly close {m['month']} won",
              mt.cell(row=r, column=2).value == m["won"])
        check(f"recalc Monthly close {m['month']} revenue",
              mt.cell(row=r, column=3).value == m["revenue_won"])

# ── PDF + charts ────────────────────────────────────────────────────
check("pdf exists", PDF.exists())
check("pdf size > 50KB", PDF.stat().st_size > 50_000,
      f"{PDF.stat().st_size} bytes")
res = subprocess.run(["pdfinfo", str(PDF)], capture_output=True,
                     text=True)
pages = None
for line in res.stdout.splitlines():
    if line.startswith("Pages:"):
        pages = int(line.split(":")[1].strip())
check("pdf is 2 pages", pages == 2, f"pages={pages}")
for chart_path, label in ((CHART_FN, "funnel"), (CHART_MO, "monthly")):
    check(f"chart {label} exists", chart_path.exists())
    data = chart_path.read_bytes() if chart_path.exists() else b""
    check(f"chart {label} is a PNG",
          data[:8] == b"\x89PNG\r\n\x1a\n")
    check(f"chart {label} size > 10KB", len(data) > 10_000,
          f"{len(data)} bytes")

# ── README + PACKAGE-NOTE carry the same headline numbers ───────────
for path, label in ((README, "README"), (PACKAGE, "PACKAGE-NOTE")):
    check(f"{label} exists", path.exists())
    if path.exists():
        txt = path.read_text(encoding="utf-8")
        for needle in (f"{total['deals']}", f"{total['win_rate']}%",
                       f"PKR {total['revenue_won']:,}",
                       f"PKR {s['forecast'][0]['forecast_weighted']:,}",
                       "SAMPLE"):
            check(f"{label} contains {needle!r}", needle in txt)
        check(f"{label} says SAMPLE, not real client work",
              "SAMPLE" in txt and "fictional" in txt.lower())

print(f"\n{len(ran) - len(fails)}/{len(ran)} checks passed, "
      f"{len(fails)} failed")
if fails:
    print("Failed checks:")
    for f in fails:
        print("  - " + f)
    sys.exit(1)
print("ALL CHECKS PASSED")
