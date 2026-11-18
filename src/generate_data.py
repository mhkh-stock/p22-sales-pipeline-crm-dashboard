#!/usr/bin/env python3
"""P22 — generate the SAMPLE sales-pipeline file for CitySoft Solutions
(SAMPLE).

Fictional Pakistani software company, 6 months of deals:
Thu 2026-01-01 -> Tue 2026-06-30 (181 days). Seeded, so the file is
reproducible. One file, one row per deal:

  data/raw/deals_SAMPLE.csv — grain: one deal
      (deal_id, created_date, month, close_date, expected_close_date,
       rep, product_line, source, stage, lost_at_stage, lost_reason,
       deal_value_pkr, probability_pct, weighted_value_pkr,
       sales_cycle_days, is_closed, is_won)

Pipeline stages: Lead -> Qualified -> Proposal -> Negotiation ->
Won / Lost. Stage probabilities (the weighting the forecast uses):
Lead 10%, Qualified 25%, Proposal 50%, Negotiation 75%, Won 100%,
Lost 0%. Weighted value = deal value x probability.

How the demo data is built (and why the report can still find real
things in it): each deal walks the stages with dwell times drawn per
stage (Lead ~7-14 days, Qualified ~10-20, Proposal ~12-25,
Negotiation ~10-20). At every stage boundary the deal either
advances or is lost there (lost_at_stage + lost_reason recorded);
a deal that survives Negotiation closes Won with a win probability
that depends on rep, product line and source, or Lost. A deal whose
walk runs past the 2026-06-30 snapshot date is still Open at the
stage it reached, with an expected_close_date drawn from the
remaining dwell times — so July/August forecast comes from real
expected close dates, not a typed number. Reps, products and sources
each have their own win propensity (Ayesha Khan wins most, Daniyal
Raza least; Referral closes best, Cold Outreach worst; ERP deals are
biggest and hardest to win), so the rep leaderboard, the stage
funnel and the forecast all have a real story to surface.

    python3 src/generate_data.py
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"

SEED = 20261118
START = date(2026, 1, 1)
END = date(2026, 6, 30)
SNAPSHOT = date(2026, 6, 30)
N_DEALS = 156

STAGES = ["Lead", "Qualified", "Proposal", "Negotiation"]
# stage: (probability %, dwell days low-high, advance probability)
STAGE_DEF = {
    "Lead": (10, (7, 14), 0.90),
    "Qualified": (25, (10, 20), 0.88),
    "Proposal": (50, (12, 25), 0.85),
    "Negotiation": (75, (10, 20), 0.82),
}
# rep: (win propensity, volume weight) — Ayesha wins most
REPS = {
    "Ayesha Khan": (0.58, 0.22),
    "Bilal Ahmed": (0.48, 0.20),
    "Sana Malik": (0.42, 0.18),
    "Usman Tariq": (0.36, 0.16),
    "Hira Shah": (0.30, 0.13),
    "Daniyal Raza": (0.24, 0.11),
}
# product line: (value low-high PKR, win propensity, volume weight)
PRODUCTS = {
    "ERP Implementation": ((800_000, 2_400_000), 0.30, 0.14),
    "Cloud Migration": ((500_000, 1_500_000), 0.36, 0.18),
    "Mobile App": ((400_000, 1_200_000), 0.42, 0.22),
    "Website Development": ((150_000, 600_000), 0.46, 0.26),
    "Digital Marketing": ((100_000, 400_000), 0.50, 0.20),
}
# source: (win propensity, volume weight)
SOURCES = {
    "Referral": (0.52, 0.24),
    "Website": (0.42, 0.30),
    "LinkedIn": (0.36, 0.26),
    "Cold Outreach": (0.26, 0.20),
}
LOST_REASONS = [
    "Price too high",
    "Chose competitor",
    "Budget frozen",
    "No decision — went quiet",
    "Requirements changed",
    "Timing — not now",
]
# lost reason weights per stage: early losses are mostly silence /
# timing, late losses are mostly price / competitor.
LOST_REASON_BY_STAGE = {
    "Lead": [0.10, 0.08, 0.12, 0.40, 0.10, 0.20],
    "Qualified": [0.14, 0.12, 0.16, 0.28, 0.12, 0.18],
    "Proposal": [0.24, 0.20, 0.16, 0.12, 0.14, 0.14],
    "Negotiation": [0.34, 0.30, 0.10, 0.06, 0.12, 0.08],
}


def pick_weighted(rng, mapping):
    items = list(mapping.items())
    weights = [v[-1] for _, v in items]
    return rng.choices([k for k, _ in items], weights=weights, k=1)[0]


def main():
    rng = random.Random(SEED)
    RAW.mkdir(parents=True, exist_ok=True)
    rows = []
    for i in range(1, N_DEALS + 1):
        # created date: uniform over the half year, slight upward tilt
        day_offset = int(rng.triangular(0, 181, 130))
        created = START + timedelta(days=min(day_offset, 180))
        rep = pick_weighted(rng, REPS)
        product = pick_weighted(rng, PRODUCTS)
        source = pick_weighted(rng, SOURCES)
        (lo, hi), prod_win, _ = PRODUCTS[product]
        rep_win = REPS[rep][0]
        src_win = SOURCES[source][0]
        win_p = max(0.15, min(0.88, (rep_win + prod_win + src_win) / 3.0
                               * 1.40))
        value = int(round(rng.uniform(lo, hi) / 5000.0) * 5000)

        # walk the stages from the created date; the walk stops at the
        # snapshot date (deal still Open) or when the deal is lost /
        # closed.
        age_days = (SNAPSHOT - created).days
        elapsed = 0
        stage = "Lead"
        lost_at = ""
        lost_reason = ""
        close_date = ""
        expected_close = ""
        is_closed = 0
        is_won = 0
        cycle = ""
        for idx, st in enumerate(STAGES):
            prob, (dlo, dhi), adv_p = STAGE_DEF[st]
            dwell = rng.randint(dlo, dhi)
            if elapsed + dwell > age_days:
                # still sitting in this stage at snapshot: Open
                stage = st
                remaining = dwell - (age_days - elapsed)
                for later in STAGES[idx + 1:]:
                    remaining += rng.randint(*STAGE_DEF[later][1])
                expected_close = (
                    SNAPSHOT + timedelta(days=remaining)).isoformat()
                break
            elapsed += dwell
            if rng.random() > adv_p:
                # lost at this stage
                stage = "Lost"
                lost_at = st
                lost_reason = rng.choices(
                    LOST_REASONS,
                    weights=LOST_REASON_BY_STAGE[st], k=1)[0]
                close_date = (created + timedelta(days=elapsed)).isoformat()
                is_closed = 1
                cycle = elapsed
                break
            if idx == len(STAGES) - 1:
                # survived Negotiation: close Won or Lost
                is_closed = 1
                cycle = elapsed
                close_date = (created + timedelta(days=elapsed)).isoformat()
                if rng.random() < win_p:
                    stage = "Won"
                    is_won = 1
                else:
                    stage = "Lost"
                    lost_at = "Negotiation"
                    lost_reason = rng.choices(
                        LOST_REASONS,
                        weights=LOST_REASON_BY_STAGE["Negotiation"],
                        k=1)[0]
                break
        else:  # pragma: no cover — loop always breaks
            stage = "Lead"
        prob_pct = 100 if stage == "Won" else (
            0 if stage == "Lost" else STAGE_DEF[stage][0])
        weighted = int(round(value * prob_pct / 100.0))
        rows.append({
            "deal_id": f"DL-{i:04d}",
            "created_date": created.isoformat(),
            "month": created.strftime("%Y-%m"),
            "close_date": close_date,
            "expected_close_date": expected_close,
            "rep": rep,
            "product_line": product,
            "source": source,
            "stage": stage,
            "lost_at_stage": lost_at,
            "lost_reason": lost_reason,
            "deal_value_pkr": value,
            "probability_pct": prob_pct,
            "weighted_value_pkr": weighted,
            "sales_cycle_days": cycle,
            "is_closed": is_closed,
            "is_won": is_won,
        })

    out = RAW / "deals_SAMPLE.csv"
    fields = list(rows[0].keys())
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)
    won = sum(r["is_won"] for r in rows)
    closed = sum(r["is_closed"] for r in rows)
    print(f"Wrote {out} — {len(rows)} deals, {closed} closed, "
          f"{won} won, {len(rows) - closed} open, "
          f"total pipeline PKR "
          f"{sum(r['deal_value_pkr'] for r in rows):,}")


if __name__ == "__main__":
    main()
