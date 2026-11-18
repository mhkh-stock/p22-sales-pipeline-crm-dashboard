#!/usr/bin/env python3
"""P22 — analyse the SAMPLE sales-pipeline file and regenerate every
output number from the raw CSV (never typed): headline pipeline KPIs,
stage funnel + conversion rates + biggest leak, rep leaderboard,
monthly trend, product / source / lost-reason breakdowns, and the
July/August forecast from open deals' expected close dates.

Derived per deal: weighted_value_pkr = deal_value_pkr x
probability_pct / 100 (probability from stage: Lead 10, Qualified
25, Proposal 50, Negotiation 75, Won 100, Lost 0); is_closed,
is_won; sales_cycle_days = close_date - created_date for closed
deals. Win rate = 100 x won / closed. A deal "reached" a stage if
its current stage is that stage or later, or it was lost at that
stage or later (Won deals reached every stage).

Forecast (two honest views, both computed): (1) pipeline-based —
sum of weighted values of open deals whose expected_close_date
falls in the month; (2) run-rate — average monthly won revenue over
the last 3 closed months (Apr–Jun 2026). Coverage = weighted open
pipeline / run-rate.

Writes:
  data/processed/deals_enriched.csv
  data/processed/stage_summary.csv
  data/processed/rep_summary.csv
  data/processed/monthly_summary.csv
  data/processed/product_summary.csv
  data/processed/source_summary.csv
  data/processed/lost_reason_summary.csv
  outputs/summary.json
  outputs/pipeline_report.md

    python3 src/analyze.py
"""
import csv
import json
from collections import Counter
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"
PROC = PROJECT / "data" / "processed"
OUT = PROJECT / "outputs"

STAGES = ["Lead", "Qualified", "Proposal", "Negotiation"]
STAGE_ORDER = ["Lead", "Qualified", "Proposal", "Negotiation",
               "Won", "Lost"]
REPS = ["Ayesha Khan", "Bilal Ahmed", "Sana Malik", "Usman Tariq",
        "Hira Shah", "Daniyal Raza"]
PRODUCTS = ["ERP Implementation", "Cloud Migration", "Mobile App",
            "Website Development", "Digital Marketing"]
SOURCES = ["Referral", "Website", "LinkedIn", "Cold Outreach"]
MONTHS = [f"2026-{m:02d}" for m in range(1, 7)]
FORECAST_MONTHS = ["2026-07", "2026-08"]
LOST_REASONS = ["Price too high", "Chose competitor",
                "Budget frozen", "No decision — went quiet",
                "Requirements changed", "Timing — not now"]


def pct(n, d):
    return round(100.0 * n / d, 2) if d else 0.0


def group_stats(rows):
    n = len(rows)
    won = [r for r in rows if r["is_won"]]
    closed = [r for r in rows if r["is_closed"]]
    lost = [r for r in rows if r["is_closed"] and not r["is_won"]]
    open_rows = [r for r in rows if not r["is_closed"]]
    cycles_won = [r["sales_cycle_days"] for r in won
                  if r["sales_cycle_days"] is not None]
    cycles_closed = [r["sales_cycle_days"] for r in closed
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
        "avg_cycle_won": round(sum(cycles_won) / len(cycles_won), 1)
        if cycles_won else 0.0,
        "avg_cycle_closed": round(
            sum(cycles_closed) / len(cycles_closed), 1)
        if cycles_closed else 0.0,
    }


def main():
    raw = list(csv.DictReader(open(RAW / "deals_SAMPLE.csv",
                                   encoding="utf-8")))
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

    total = group_stats(rows)
    total["total_deals"] = total["deals"]

    # ── funnel: reached counts ──────────────────────────────────────
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

    reached_counts = {s: len(reached(s)) for s in STAGES}
    reached_counts["Close decision"] = total["closed"]
    reached_counts["Won"] = total["won"]
    funnel = []
    steps = [("Lead", len(rows)), ("Qualified", reached_counts["Qualified"]),
             ("Proposal", reached_counts["Proposal"]),
             ("Negotiation", reached_counts["Negotiation"]),
             ("Won", total["won"])]
    prev_n = None
    for label, n in steps:
        conv = pct(n, prev_n) if prev_n is not None else 100.0
        funnel.append({"stage": label, "reached": n,
                       "conversion_from_previous": conv,
                       "share_of_all_deals": pct(n, len(rows))})
        prev_n = n
    conv_steps = [(funnel[i]["stage"], funnel[i]["conversion_from_previous"],
                   funnel[i - 1]["stage"])
                  for i in range(1, len(funnel))]
    leak_stage, leak_rate, leak_from = min(conv_steps, key=lambda x: x[1])
    leak_step = f"{leak_from} → {leak_stage}"

    # ── stage snapshot (current stage) ──────────────────────────────
    stage_summary = []
    for s in STAGE_ORDER:
        g = [r for r in rows if r["stage"] == s]
        stage_summary.append({
            "stage": s, "deals": len(g),
            "share_of_deals": pct(len(g), len(rows)),
            "pipeline_value": sum(r["deal_value_pkr"] for r in g),
            "weighted_value": sum(r["weighted_value_pkr"] for r in g),
        })

    # ── reps / months / products / sources / lost reasons ──────────
    def by_group(key, values):
        out = []
        for v in values:
            g = [r for r in rows if r[key] == v]
            d = group_stats(g)
            d["name"] = v
            out.append(d)
        return out

    reps = by_group("rep", REPS)
    reps_by_revenue = sorted(reps, key=lambda x: -x["revenue_won"])
    reps_by_winrate = sorted(
        [r for r in reps if r["closed"] >= 10],
        key=lambda x: x["win_rate"])
    months = by_group("month", MONTHS)
    products = by_group("product_line", PRODUCTS)
    sources = by_group("source", SOURCES)

    lost_rows = [r for r in rows if r["is_closed"] and not r["is_won"]]
    lost_reasons = []
    for reason in LOST_REASONS:
        g = [r for r in lost_rows if r["lost_reason"] == reason]
        lost_reasons.append({
            "lost_reason": reason, "lost_deals": len(g),
            "share_of_lost": pct(len(g), len(lost_rows)),
            "value_lost": sum(r["deal_value_pkr"] for r in g),
        })
    lost_reasons.sort(key=lambda x: -x["lost_deals"])
    lost_at = Counter(r["lost_at_stage"] for r in lost_rows)

    # ── won revenue by close month (run-rate) ───────────────────────
    won_by_close_month = []
    for m in MONTHS:
        g = [r for r in rows if r["is_won"] and r["close_date"]
             and r["close_date"][:7] == m]
        won_by_close_month.append({
            "month": m, "won": len(g),
            "revenue_won": sum(r["deal_value_pkr"] for r in g),
        })
    last3 = [x for x in won_by_close_month if x["month"] in
             ["2026-04", "2026-05", "2026-06"]]
    run_rate = round(sum(x["revenue_won"] for x in last3) / 3.0, 2)

    # ── forecast from expected close dates ──────────────────────────
    forecast = []
    for m in FORECAST_MONTHS:
        g = [r for r in rows if not r["is_closed"]
             and r["expected_close_date"]
             and r["expected_close_date"][:7] == m]
        forecast.append({
            "month": m,
            "open_deals_expected": len(g),
            "open_value_expected": sum(r["deal_value_pkr"] for r in g),
            "forecast_weighted": sum(r["weighted_value_pkr"] for r in g),
        })
    july = forecast[0]
    coverage = round(total["weighted_open"] / run_rate, 2) if run_rate else 0.0

    worst_rep = reps_by_winrate[0] if reps_by_winrate else None
    worst_product = min(products, key=lambda x: x["win_rate"])
    worst_source = min(sources, key=lambda x: x["win_rate"])
    best_rep = reps_by_revenue[0]
    top_lost_reason = lost_reasons[0]

    leak_idx = [f["stage"] for f in funnel].index(leak_stage)
    leak_reached = funnel[leak_idx]["reached"]
    leak_prev_reached = funnel[leak_idx - 1]["reached"]
    actions = [
        f"Fix the {leak_step} step first — it is the biggest leak. "
        f"Only {leak_rate}% of deals that reach {leak_from} move on "
        f"to {leak_stage} ({leak_reached} of {leak_prev_reached}). "
        f"{lost_at.get('Lead', 0) + lost_at.get('Qualified', 0)} deals "
        f"were lost at Lead or Qualified — tighten qualification "
        f"before spending on later stages.",
        f"Coach {worst_rep['name']} first — lowest win rate among reps "
        f"with 10+ closed deals ({worst_rep['win_rate']}%, "
        f"{worst_rep['won']} of {worst_rep['closed']} closed), against "
        f"{best_rep['name']} at {best_rep['win_rate']}% with "
        f"PKR {best_rep['revenue_won']:,} won. Pair them on the next "
        f"five live deals and copy what the top rep does at "
        f"{leak_step}.",
        f"Attack the top lost reason: {top_lost_reason['lost_reason']} "
        f"— {top_lost_reason['lost_deals']} lost deals "
        f"({top_lost_reason['share_of_lost']}% of all losses), "
        f"PKR {top_lost_reason['value_lost']:,} of pipeline value. "
        f"Price objections dominate late losses, so put a written "
        f"price-objection script and a discount-approval rule in place "
        f"before the next Negotiation.",
        f"Push the July pipeline, don't just hope for it. "
        f"{july['open_deals_expected']} open deals are expected to "
        f"close in July 2026 worth PKR "
        f"{july['forecast_weighted']:,} weighted — against a "
        f"PKR {run_rate:,.0f} monthly run-rate. Weighted open pipeline "
        f"is {coverage}x one month's won revenue, so July depends on "
        f"closing what's already in the pipe, not on finding new "
        f"leads.",
    ]

    summary = {
        "client": "CitySoft Solutions (SAMPLE)",
        "period": "2026-01-01 to 2026-06-30",
        "months": MONTHS,
        "total_deals": total["deals"],
        "total_pipeline_value": total["pipeline_value"],
        "open_deals": total["open_deals"],
        "open_value": total["open_value"],
        "weighted_open_pipeline": total["weighted_open"],
        "closed_deals": total["closed"],
        "won_deals": total["won"],
        "lost_deals": total["lost"],
        "win_rate": total["win_rate"],
        "revenue_won": total["revenue_won"],
        "value_lost": total["value_lost"],
        "avg_deal_size": total["avg_deal_size"],
        "avg_won_deal_size": total["avg_won_deal_size"],
        "avg_cycle_won": total["avg_cycle_won"],
        "avg_cycle_closed": total["avg_cycle_closed"],
        "funnel": funnel,
        "biggest_leak_stage": leak_stage,
        "biggest_leak_step": leak_step,
        "biggest_leak_rate": leak_rate,
        "stage_summary": stage_summary,
        "reps": reps,
        "reps_by_revenue": [r["name"] for r in reps_by_revenue],
        "best_rep": best_rep["name"],
        "worst_rep": worst_rep["name"] if worst_rep else None,
        "months_summary": months,
        "products": products,
        "sources": sources,
        "lost_reasons": lost_reasons,
        "lost_at_stage": {s: lost_at.get(s, 0) for s in STAGES},
        "won_by_close_month": won_by_close_month,
        "run_rate_3m": run_rate,
        "forecast": forecast,
        "july_forecast_weighted": july["forecast_weighted"],
        "pipeline_coverage_x": coverage,
        "actions": actions,
    }

    PROC.mkdir(parents=True, exist_ok=True)
    OUT.mkdir(parents=True, exist_ok=True)

    def write_csv(path, fieldnames, rows_):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(rows_)

    write_csv(PROC / "deals_enriched.csv", list(rows[0].keys()), rows)
    write_csv(PROC / "stage_summary.csv",
              ["stage", "deals", "share_of_deals", "pipeline_value",
               "weighted_value"], stage_summary)
    write_csv(PROC / "rep_summary.csv",
              ["name", "deals", "pipeline_value", "open_deals",
               "open_value", "weighted_open", "closed", "won", "lost",
               "win_rate", "revenue_won", "value_lost", "avg_deal_size",
               "avg_won_deal_size", "avg_cycle_won",
               "avg_cycle_closed"], reps)
    write_csv(PROC / "monthly_summary.csv",
              ["name", "deals", "pipeline_value", "open_deals",
               "open_value", "weighted_open", "closed", "won", "lost",
               "win_rate", "revenue_won", "value_lost", "avg_deal_size",
               "avg_won_deal_size", "avg_cycle_won",
               "avg_cycle_closed"], months)
    write_csv(PROC / "product_summary.csv",
              ["name", "deals", "pipeline_value", "open_deals",
               "open_value", "weighted_open", "closed", "won", "lost",
               "win_rate", "revenue_won", "value_lost", "avg_deal_size",
               "avg_won_deal_size", "avg_cycle_won",
               "avg_cycle_closed"], products)
    write_csv(PROC / "source_summary.csv",
              ["name", "deals", "pipeline_value", "open_deals",
               "open_value", "weighted_open", "closed", "won", "lost",
               "win_rate", "revenue_won", "value_lost", "avg_deal_size",
               "avg_won_deal_size", "avg_cycle_won",
               "avg_cycle_closed"], sources)
    write_csv(PROC / "lost_reason_summary.csv",
              ["lost_reason", "lost_deals", "share_of_lost",
               "value_lost"], lost_reasons)

    (OUT / "summary.json").write_text(
        json.dumps(summary, indent=2, ensure_ascii=False),
        encoding="utf-8")

    lines = [
        "# P22 — Sales Pipeline / CRM Dashboard: CitySoft Solutions "
        "(SAMPLE)",
        "",
        f"**Period:** Thu 2026-01-01 → Tue 2026-06-30 (6 months, "
        f"181 days)",
        f"**Deals:** {total['deals']} · **Total pipeline:** "
        f"PKR {total['pipeline_value']:,} · **Open pipeline:** "
        f"PKR {total['open_value']:,} ({total['open_deals']} deals, "
        f"weighted PKR {total['weighted_open']:,}) · "
        f"**Won:** {total['won']} of {total['closed']} closed "
        f"({total['win_rate']}%) · **Revenue won:** "
        f"PKR {total['revenue_won']:,} · **Avg deal:** "
        f"PKR {total['avg_deal_size']:,.0f} · **Avg sales cycle (won):** "
        f"{total['avg_cycle_won']} days",
        "",
        "## 1. The headline — where is next month's revenue?",
        "",
        f"Of {total['deals']} deals worth "
        f"PKR {total['pipeline_value']:,}, {total['won']} were won "
        f"(PKR {total['revenue_won']:,} banked) and {total['lost']} "
        f"lost (PKR {total['value_lost']:,} of pipeline value gone). "
        f"{total['open_deals']} deals are still open, worth "
        f"PKR {total['open_value']:,} — weighted by stage "
        f"probability that is PKR {total['weighted_open']:,} of "
        f"expected revenue. July 2026's pipeline-based forecast is "
        f"PKR {july['forecast_weighted']:,} "
        f"({july['open_deals_expected']} deals expected to close in "
        f"July), against a PKR {run_rate:,.0f} monthly run-rate "
        f"(average won revenue Apr–Jun). Weighted open pipeline "
        f"covers {coverage}x one month's won revenue.",
        "",
        "## 2. The funnel — where deals leak",
        "",
        "| Stage reached | Deals | Conversion from previous | "
        "Share of all deals |",
        "|---|---|---|---|",
    ]
    for f in funnel:
        lines.append(f"| {f['stage']} | {f['reached']} | "
                     f"{f['conversion_from_previous']}% | "
                     f"{f['share_of_all_deals']}% |")
    lines += [
        "",
        f"**Biggest leak: {leak_step} ({leak_rate}%).** "
        f"Lost at stage: " + ", ".join(
            f"{s} {lost_at.get(s, 0)}" for s in STAGES) + ".",
        "",
        "## 3. Rep leaderboard (by revenue won)",
        "",
        "| Rep | Deals | Closed | Won | Win rate | Revenue won | "
        "Avg won deal | Avg cycle (won) | Weighted open |",
        "|---|---|---|---|---|---|---|---|",
    ]
    for r in reps_by_revenue:
        lines.append(f"| {r['name']} | {r['deals']} | {r['closed']} | "
                     f"{r['won']} | {r['win_rate']}% | "
                     f"PKR {r['revenue_won']:,} | "
                     f"PKR {r['avg_won_deal_size']:,.0f} | "
                     f"{r['avg_cycle_won']} days | "
                     f"PKR {r['weighted_open']:,} |")
    lines += [
        "",
        "## 4. Month by month (by created month)",
        "",
        "| Month | Deals | Pipeline value | Closed | Won | Win rate | "
        "Revenue won |",
        "|---|---|---|---|---|---|---|",
    ]
    for m in months:
        lines.append(f"| {m['name']} | {m['deals']} | "
                     f"PKR {m['pipeline_value']:,} | {m['closed']} | "
                     f"{m['won']} | {m['win_rate']}% | "
                     f"PKR {m['revenue_won']:,} |")
    lines += [
        "",
        "## 5. Won revenue by close month (the run-rate)",
        "",
        "| Close month | Deals won | Revenue won |",
        "|---|---|---|",
    ]
    for m in won_by_close_month:
        lines.append(f"| {m['month']} | {m['won']} | "
                     f"PKR {m['revenue_won']:,} |")
    lines += [
        "",
        f"3-month run-rate (Apr–Jun 2026): PKR {run_rate:,.0f} per "
        f"month.",
        "",
        "## 6. Forecast",
        "",
        "| Month | Open deals expected | Open value | Forecast "
        "(weighted) |",
        "|---|---|---|---|",
    ]
    for f in forecast:
        lines.append(f"| {f['month']} | {f['open_deals_expected']} | "
                     f"PKR {f['open_value_expected']:,} | "
                     f"PKR {f['forecast_weighted']:,} |")
    lines += [
        "",
        "## 7. By product line",
        "",
        "| Product line | Deals | Closed | Won | Win rate | Revenue won |",
        "|---|---|---|---|---|---|",
    ]
    for p in sorted(products, key=lambda x: -x["revenue_won"]):
        lines.append(f"| {p['name']} | {p['deals']} | {p['closed']} | "
                     f"{p['won']} | {p['win_rate']}% | "
                     f"PKR {p['revenue_won']:,} |")
    lines += [
        "",
        "## 8. By source",
        "",
        "| Source | Deals | Closed | Won | Win rate | Revenue won |",
        "|---|---|---|---|---|---|",
    ]
    for s_ in sorted(sources, key=lambda x: -x["revenue_won"]):
        lines.append(f"| {s_['name']} | {s_['deals']} | {s_['closed']} | "
                     f"{s_['won']} | {s_['win_rate']}% | "
                     f"PKR {s_['revenue_won']:,} |")
    lines += [
        "",
        "## 9. Why deals were lost",
        "",
        "| Lost reason | Lost deals | Share of losses | Value lost |",
        "|---|---|---|---|",
    ]
    for lr in lost_reasons:
        lines.append(f"| {lr['lost_reason']} | {lr['lost_deals']} | "
                     f"{lr['share_of_lost']}% | "
                     f"PKR {lr['value_lost']:,} |")
    lines += ["", "## 10. The four recommendations", ""]
    for i, a in enumerate(actions, 1):
        lines.append(f"{i}. **{a}**")
    lines.append("")
    (OUT / "pipeline_report.md").write_text(
        "\n".join(lines), encoding="utf-8")
    print(f"Analysed {len(rows)} deals: {total['won']} won "
          f"({total['win_rate']}%), revenue won "
          f"PKR {total['revenue_won']:,}, weighted open "
          f"PKR {total['weighted_open']:,}, July forecast "
          f"PKR {july['forecast_weighted']:,}, biggest leak "
          f"{leak_step} ({leak_rate}%)")


if __name__ == "__main__":
    main()
