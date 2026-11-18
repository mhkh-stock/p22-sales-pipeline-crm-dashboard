# P22 — Sales Pipeline / CRM Dashboard

**Client (SAMPLE):** CitySoft Solutions — fictional Pakistani
software company, 6 months (Thu 2026-01-01 → Tue 2026-06-30,
181 days) · **Status: SAMPLE delivery on DEMO data**
**Repo:** https://github.com/mhkh-stock/p22-sales-pipeline-crm-dashboard
**Tools:** Python · Excel (live formulas) · reportlab
**Deliverable:** `outputs/pipeline_report.pdf` (2 pages) +
`outputs/pipeline_report_SAMPLE.xlsx`

> ⚠️ **What this is and is not.** The deals file is a fictional
> company's demo data, generated to show the method. No real
> company, rep or client is involved. What is real is the analysis:
> the same funnel, win-rate, rep-leaderboard and forecast
> breakdowns and workbook run on a real client's CRM export the
> same way, and every number in the report is recomputed from the
> data, never typed.

## Business question it answers

> "Where is my next month's revenue?"

A team that only watches total pipeline cannot tell whether its
revenue is stuck at qualification, dying in negotiation, carried
by one rep, or coming from one source — or how much of next
month is already in the pipe. This delivery answers it from the
deals: the funnel with conversion at every step and the biggest
leak named, win rate overall and per rep / product / source,
average deal size and sales cycle, the rep leaderboard, and a
July forecast built deal by deal from expected close dates —
with a run-rate beside it so the forecast can be checked, not
just believed.

## The method (plainly stated)

**Win rate = 100 × won / closed.** Open deals are not in the
denominator — a deal still in the pipe has not been won or lost
yet.

**Weighted pipeline = deal value × stage probability.**
Probabilities are the team's planning weights: Lead 10%,
Qualified 25%, Proposal 50%, Negotiation 75%, Won 100%, Lost 0%.
They are not measured close rates, and the workbook shows the
unweighted value next to every weighted value so the reader can
see both.

**Funnel:** a deal "reached" a stage if its current stage is
that stage or later, or it was lost at that stage or later. Won
deals reached every stage. Conversion = reached this stage /
reached the previous stage. The biggest leak is the step with
the lowest conversion: Negotiation → Won at 51.95%.

**Sales cycle = close date − created date, in days,** for closed
deals only. Open deals have no cycle yet.

**Forecast, two views.** (1) Pipeline-based: the weighted value
of open deals whose expected close date falls in the month —
July 2026 is PKR 13,080,000 from 27 open deals. (2) Run-rate:
average monthly revenue won over Apr–Jun 2026, PKR 8,496,667.
Coverage = weighted open pipeline / run-rate = 1.94x: the open
pipe already holds almost two months of typical winning.

**Limits, stated upfront:** win rate by created month is only
meaningful for mature cohorts — deals created in May–Jun 2026
mostly have not had time to close yet (June: 3 of 19 closed, 0
won so far), so their win rates read near zero. Small bases are
noisy: Daniyal Raza shows the highest win rate (55.56%) on only
9 closed deals, so the Dashboard's lowest-win-rate row only
ranks reps with 10+ closed deals. The July forecast is
pipeline-based, not a promise — if those 27 deals slip, July
lands closer to the run-rate instead.

## What is inside

| File | What it does |
|---|---|
| `outputs/pipeline_report.pdf` | The deliverable: 2-page client report — headline, funnel + biggest leak, rep leaderboard, monthly trend, forecast, product / source, lost reasons, 4 recommendations |
| `outputs/pipeline_report_SAMPLE.xlsx` | The working report: Dashboard (KPIs, live formulas) + Pipeline + Stages + Reps + Monthly Trend + Forecast + Breakdown + Method sheets |
| `outputs/pipeline_report.md` | Findings report: headline numbers, all tables, the reasoning |
| `outputs/summary.json` | Machine-readable summary (headline KPIs + funnel + reps + months + products + sources + lost reasons + forecast + actions) |
| `outputs/funnel_stage.png` | Funnel chart — deals that reached each stage (render of the same data) |
| `outputs/monthly_trend.png` | Monthly trend — deals created and revenue won by created month (render of the same data) |
| `data/raw/deals_SAMPLE.csv` | The deals file: 156 deals, rep + product line + source + stage + value + dates + lost reason |
| `data/processed/deals_enriched.csv` | All 156 deals as read by `src/analyze.py` |
| `data/processed/stage_summary.csv` | 6-row current-stage summary produced by `src/analyze.py` |
| `data/processed/rep_summary.csv` | 6-row rep summary produced by `src/analyze.py` |
| `data/processed/monthly_summary.csv` | 6-month summary produced by `src/analyze.py` |
| `data/processed/product_summary.csv` | 5-row product-line summary produced by `src/analyze.py` |
| `data/processed/source_summary.csv` | 4-row source summary produced by `src/analyze.py` |
| `data/processed/lost_reason_summary.csv` | 6-row lost-reason summary produced by `src/analyze.py` |
| `data/README.md` | Data dictionary + known quirks |
| `src/generate_data.py` | Makes the demo file (seeded; NOT run on a real client file) |
| `src/analyze.py` | Recomputes every headline number from the raw CSV |
| `src/build_workbook.py` | Builds the workbook: 8 sheets, live formulas |
| `src/build_report.py` | Builds the 2-page PDF + charts from `outputs/summary.json` |
| `src/checks.py` | Data-quality + reconciliation checks as code — 2,632 checks, all passing |
| `requirements.txt` | Pinned: `openpyxl==3.1.2`, `reportlab==5.0.1`, `matplotlib==3.6.3` (Python 3.12) |

## How it fits together

```
Client CRM export (same columns) ──▶ src/analyze.py ──▶ summary.json
        (156 rows in the demo)                + pipeline_report.md
                                                 ▼
                                       src/build_workbook.py ──▶ pipeline_report_SAMPLE.xlsx
                                                 │              (8 sheets, live formulas)
                                                 ▼
                                       src/build_report.py ──▶ pipeline_report.pdf (2 pages)
                                                 ▲ src/checks.py verifies all of it (2,632 checks)
```

## Headline numbers (SAMPLE — CitySoft Solutions, 6 months)

- **156 deals** · **total pipeline PKR 117,175,000** · **41 open
  deals (PKR 31,925,000, weighted PKR 16,441,250)** · **115
  closed** · **40 won (34.78%)** · **75 lost** · **revenue won
  PKR 26,135,000** · **value lost PKR 59,115,000** ·
  **average deal PKR 751,122** · **average won deal
  PKR 653,375** · **average sales cycle 58.7 days (won),
  44.6 days (all closed)**
- **July 2026 forecast (pipeline-based): PKR 13,080,000** (27 open
  deals expected to close in July) · **August forecast:
  PKR 3,361,250** (14 deals) · **run-rate PKR 8,496,667/month**
  (average revenue won, Apr–Jun) · **pipeline coverage 1.94x**
- **Biggest leak: Negotiation → Won (51.95%)** — only 40 of the
  77 deals that reached Negotiation were won
- **Top rep by revenue won: Bilal Ahmed (PKR 6,625,000, 12 won)**;
  lowest win rate among reps with 10+ closed: **Usman Tariq
  (29.41%)**
- **Best source: Website (43.59% win rate, PKR 12,285,000 won)**;
  worst: **LinkedIn (20.00%)**
- **Hardest product line: ERP Implementation (21.05% win rate)**,
  but biggest average deal (PKR 1,638,200); easiest: **Website
  Development (48.39%)**
- **Top lost reason: Price too high — 19 lost deals (25.33%),
  PKR 11,930,000 of pipeline value**

## The funnel — where deals leak

| Stage reached | Deals | Conversion from previous | Share of all deals |
|---|---|---|---|
| Lead | 156 | 100.00% | 100.00% |
| Qualified | 141 | 90.38% | 90.38% |
| Proposal | 112 | 79.43% | 71.79% |
| Negotiation | 77 | 68.75% | 49.36% |
| Won | 40 | 51.95% | 25.64% |

The early funnel is healthy — 90.38% of leads get qualified.
The leak is at the end: only 51.95% of deals that reach
Negotiation become Won. Lost at stage: Lead 13, Qualified 20,
Proposal 17, Negotiation 25.

## Rep leaderboard (by revenue won)

| Rep | Deals | Closed | Won | Win rate | Revenue won | Avg won deal | Avg cycle (won) | Weighted open |
|---|---|---|---|---|---|---|---|---|
| Bilal Ahmed | 44 | 36 | 12 | 33.33% | PKR 6,625,000 | PKR 552,083 | 57.8 days | PKR 3,790,500 |
| Ayesha Khan | 29 | 24 | 8 | 33.33% | PKR 5,360,000 | PKR 670,000 | 60.8 days | PKR 2,310,000 |
| Sana Malik | 31 | 19 | 7 | 36.84% | PKR 5,155,000 | PKR 736,429 | 61.7 days | PKR 5,150,000 |
| Daniyal Raza | 16 | 9 | 5 | 55.56% | PKR 4,515,000 | PKR 903,000 | 56.2 days | PKR 2,616,250 |
| Usman Tariq | 22 | 17 | 5 | 29.41% | PKR 2,505,000 | PKR 501,000 | 56.4 days | PKR 1,863,250 |
| Hira Shah | 14 | 10 | 3 | 30.00% | PKR 1,975,000 | PKR 658,333 | 57.3 days | PKR 711,250 |

Bilal Ahmed leads on revenue won because he carries the most
deals (44). Sana Malik has the best win rate among reps with
10+ closed (36.84%). Daniyal Raza's 55.56% sits on only 9 closed
deals — five wins — so it is noise, not a lead to copy. Sana
Malik also holds the biggest weighted open pipeline
(PKR 5,150,000), so the next month's revenue runs through her
deals as much as Bilal's.

## Month by month (by created month)

| Month | Deals | Pipeline value | Closed | Won | Win rate | Revenue won |
|---|---|---|---|---|---|---|
| 2026-01 | 5 | PKR 2,990,000 | 5 | 2 | 40.00% | PKR 645,000 |
| 2026-02 | 14 | PKR 10,020,000 | 14 | 6 | 42.86% | PKR 4,975,000 |
| 2026-03 | 31 | PKR 24,040,000 | 31 | 14 | 45.16% | PKR 9,460,000 |
| 2026-04 | 43 | PKR 30,525,000 | 42 | 17 | 40.48% | PKR 10,275,000 |
| 2026-05 | 44 | PKR 34,540,000 | 20 | 1 | 5.00% | PKR 780,000 |
| 2026-06 | 19 | PKR 15,060,000 | 3 | 0 | 0.00% | PKR 0 |

May and June read near zero because those deals have not had
time to close yet — a deal needs roughly 40–80 days to walk the
stages. The mature cohorts (Jan–Apr) all win at 40–45%. Won
revenue by close month tells the real trend: March PKR 645,000,
April PKR 4,620,000, May PKR 10,440,000, June PKR 10,430,000 —
the run-rate of PKR 8,496,667/month comes from Apr–Jun.

## Forecast — July and August 2026

| Month | Open deals expected | Open value | Forecast (weighted) |
|---|---|---|---|
| 2026-07 | 27 | PKR 20,570,000 | PKR 13,080,000 |
| 2026-08 | 14 | PKR 11,355,000 | PKR 3,361,250 |

July's forecast (PKR 13,080,000) sits above the PKR 8,496,667
run-rate because 27 open deals are expected to close in July —
12 of them already in Negotiation. August is smaller
(PKR 3,361,250) because only 14 open deals are expected then.
Weighted open pipeline overall is PKR 16,441,250, or 1.94x one
month's won revenue.

## By product line

| Product line | Deals | Closed | Won | Win rate | Revenue won | Avg deal size |
|---|---|---|---|---|---|---|
| ERP Implementation | 25 | 19 | 4 | 21.05% | PKR 6,960,000 | PKR 1,638,200 |
| Cloud Migration | 30 | 22 | 7 | 31.82% | PKR 6,540,000 | PKR 961,167 |
| Website Development | 40 | 31 | 15 | 48.39% | PKR 5,395,000 | PKR 368,000 |
| Mobile App | 33 | 18 | 7 | 38.89% | PKR 5,170,000 | PKR 767,727 |
| Digital Marketing | 28 | 25 | 7 | 28.00% | PKR 2,070,000 | PKR 261,786 |

ERP is the hardest close (21.05%) and the biggest ticket
(PKR 1,638,200 average) — it still brings in the most revenue
won (PKR 6,960,000) because four wins are worth more than
fifteen small ones. Website Development is the volume engine:
40 deals, 48.39% win rate.

## By source

| Source | Deals | Closed | Won | Win rate | Revenue won |
|---|---|---|---|---|---|
| Website | 54 | 39 | 17 | 43.59% | PKR 12,285,000 |
| Referral | 36 | 26 | 10 | 38.46% | PKR 4,055,000 |
| Cold Outreach | 25 | 20 | 7 | 35.00% | PKR 6,340,000 |
| LinkedIn | 41 | 30 | 6 | 20.00% | PKR 3,455,000 |

Website is the best source on both rate and revenue — 54 deals,
PKR 12,285,000 won. LinkedIn brings volume (41 deals) but the
worst win rate (20.00%): it fills the pipe with deals that do
not close.

## Why deals were lost

| Lost reason | Lost deals | Share of losses | Value lost |
|---|---|---|---|
| Price too high | 19 | 25.33% | PKR 11,930,000 |
| Timing — not now | 16 | 21.33% | PKR 13,100,000 |
| No decision — went quiet | 14 | 18.67% | PKR 9,025,000 |
| Chose competitor | 10 | 13.33% | PKR 8,760,000 |
| Budget frozen | 8 | 10.67% | PKR 10,120,000 |
| Requirements changed | 8 | 10.67% | PKR 6,180,000 |

Price too high leads on count (19 deals). Timing — not now
leads on money (PKR 13,100,000) — those are deals that may come
back, not deals that are dead.

## The four recommendations (from the data, not from a template)

1. **Fix the Negotiation → Won step first — it is the biggest
   leak.** Only 51.95% of deals that reach Negotiation move on to
   Won (40 of 77). 33 deals were lost at Lead or Qualified —
   tighten qualification before spending on later stages.
2. **Coach Usman Tariq first — lowest win rate among reps with
   10+ closed deals (29.41%, 5 of 17 closed), against Bilal Ahmed
   at 33.33% with PKR 6,625,000 won.** Pair them on the next five
   live deals and copy what the top rep does at Negotiation → Won.
3. **Attack the top lost reason: Price too high — 19 lost deals
   (25.33% of all losses), PKR 11,930,000 of pipeline value.**
   Price objections dominate late losses, so put a written
   price-objection script and a discount-approval rule in place
   before the next Negotiation.
4. **Push the July pipeline, don't just hope for it.** 27 open
   deals are expected to close in July 2026 worth
   PKR 13,080,000 weighted — against a PKR 8,496,667 monthly
   run-rate. Weighted open pipeline is 1.94x one month's won
   revenue, so July depends on closing what's already in the pipe,
   not on finding new leads.

## Same logic, two ways — and they must agree

The headline KPIs, the funnel, the rep, month, product and
source totals exist in two forms: the Python in `src/analyze.py`
and this workbook's formulas. `src/checks.py` runs both against
the raw file and compares them — a mistake in either one shows up
as a failure. The workbook's formulas were also recalculated in
LibreOffice (including the `_xlfn.MINIFS` / `_xlfn.MAXIFS`
lowest-win-rate and best-source cells): every total matches the
recompute exactly, with zero error values.

## Reproduce it (2 minutes)

```bash
python3 src/generate_data.py   # rebuild the SAMPLE file (seed 20261118)
python3 src/analyze.py         # summary.json + pipeline_report.md
python3 src/build_workbook.py  # outputs/pipeline_report_SAMPLE.xlsx
python3 src/build_report.py    # outputs/pipeline_report.pdf (2 pages)
python3 src/checks.py          # 2,632 checks — expect all passing
```

On a real client file, replace `data/raw/deals_SAMPLE.csv`
with the CRM export (same columns) and skip
`generate_data.py` — everything downstream runs unchanged.
