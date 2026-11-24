"""Generate simulated CRM data (customers, orders) for three retail personas and
aggregate it into the JSON payload consumed by the training dashboard.

All data is synthetic. Persona names only indicate the *style* of retailer
(apparel / drugstore-beauty / home-lifestyle); no real brand data is used.

Usage:
    python tools/generate_crm_data.py
Outputs:
    output/data/<brand>_customers.csv, <brand>_orders.csv
    output/data/dashboard_data.json
"""
import json
import random
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "data"
CONFIG = ROOT / "src" / "main" / "resources" / "config" / "brands.json"

SNAPSHOT = date(2026, 10, 31)
START = date(2024, 11, 1)
SEED = 42


def load_brands():
    return json.loads(CONFIG.read_text(encoding="utf-8"))


def simulate_brand(key, cfg, rng):
    n = cfg["customers"]
    days = (SNAPSHOT - START).days
    channels = list(cfg["channels"])
    cust_rows, order_rows = [], []
    oid = 0
    for i in range(n):
        cid = f"{key[:2].upper()}{i:05d}"
        # join date skewed toward recent so new/old mix exists
        join_off = int(days * (1 - rng.random() ** 1.6))
        join = START + timedelta(days=join_off)
        # latent loyalty drives purchase frequency; ~20% churn early
        loyal = rng.betavariate(1.2, 3.0)
        gap = max(12, int(cfg["base_gap_days"] * (1.6 - loyal * 1.4)))
        churn_after = rng.random() < 0.35
        stop = join + timedelta(days=rng.randint(30, 300)) if churn_after else SNAPSHOT
        channel = rng.choices(channels, weights=list(cfg["channels"].values()))[0]
        d = join
        first = True
        while d <= min(stop, SNAPSHOT):
            aov = max(
                cfg["aov_floor"],
                rng.lognormvariate(np.log(cfg["aov_median"]), 0.55) * (1 + loyal * 0.6),
            )
            if first:
                aov *= cfg["first_order_factor"]
            cat = rng.choices(
                list(cfg["categories"]), weights=list(cfg["categories"].values())
            )[0]
            oid += 1
            order_rows.append(
                (f"{key[:2].upper()}O{oid:07d}", cid, d.isoformat(), round(aov), cat, channel)
            )
            first = False
            d += timedelta(days=max(3, int(rng.gauss(gap, gap * 0.35))))
        cust_rows.append((cid, join.isoformat(), channel, round(loyal, 3)))
    customers = pd.DataFrame(cust_rows, columns=["customer_id", "join_date", "channel", "loyalty_latent"])
    orders = pd.DataFrame(
        order_rows, columns=["order_id", "customer_id", "order_date", "amount", "category", "channel"]
    )
    return customers, orders


def rfm_table(customers, orders, tiers):
    o = orders.copy()
    o["order_date"] = pd.to_datetime(o["order_date"])
    snap = pd.Timestamp(SNAPSHOT)
    g = o.groupby("customer_id")
    t = pd.DataFrame(
        {
            "recency": (snap - g["order_date"].max()).dt.days,
            "frequency": g.size(),
            "monetary": g["amount"].sum(),
            "first_order": g["order_date"].min(),
        }
    ).reset_index()
    t = customers.merge(t, on="customer_id", how="left")
    t["frequency"] = t["frequency"].fillna(0).astype(int)
    t["monetary"] = t["monetary"].fillna(0)
    t["recency"] = t["recency"].fillna(999).astype(int)
    t["aov"] = np.where(t["frequency"] > 0, t["monetary"] / t["frequency"].clip(lower=1), 0).round()
    # quintile scores (5 = best)
    def score(s, reverse=False):
        r = s.rank(method="first", ascending=not reverse)
        return pd.qcut(r, 5, labels=[1, 2, 3, 4, 5]).astype(int)

    t["R"] = score(t["recency"], reverse=True)
    t["F"] = score(t["frequency"])
    t["M"] = score(t["monetary"])
    t["segment"] = t.apply(segment_of, axis=1)
    # lifecycle status used in new / old / dormant tracking
    t["lifecycle"] = t.apply(lifecycle_of, axis=1)
    # member tier by trailing-12-month spend
    cutoff = snap - pd.Timedelta(days=365)
    ttm = o[o["order_date"] > cutoff].groupby("customer_id")["amount"].sum()
    t["ttm_spend"] = t["customer_id"].map(ttm).fillna(0)
    t["tier"] = t["ttm_spend"].apply(lambda v: tier_of(v, tiers))
    return t


def segment_of(r):
    R, F, M = r["R"], r["F"], r["M"]
    if r["frequency"] == 0:
        return "Prospect"
    if R >= 4 and F >= 4 and M >= 4:
        return "Champions"
    if R >= 3 and F >= 4:
        return "Loyal"
    if R >= 4 and F <= 2:
        return "New / Promising"
    if R == 3 and F <= 3:
        return "Need Attention"
    if R <= 2 and F >= 4:
        return "At Risk (VIP)"
    if R <= 2 and F >= 2:
        return "Hibernating"
    return "Lost"


def lifecycle_of(r):
    if r["frequency"] == 0:
        return "Never bought"
    if r["recency"] <= 90 and r["frequency"] == 1:
        return "New (1st order, <=90d)"
    if r["recency"] <= 90:
        return "Active"
    if r["recency"] <= 180:
        return "At risk (91-180d)"
    if r["recency"] <= 365:
        return "Dormant (181-365d)"
    return "Lost (>365d)"


def tier_of(v, tiers):
    name = tiers[0]["name"]
    for t in tiers:
        if v >= t["min"]:
            name = t["name"]
    return name


def monthly_cohort(orders, customers, max_m=12):
    o = orders.copy()
    o["order_date"] = pd.to_datetime(o["order_date"])
    first = o.groupby("customer_id")["order_date"].min().rename("first")
    o = o.join(first, on="customer_id")
    o["cohort"] = o["first"].dt.to_period("M")
    o["idx"] = (o["order_date"].dt.year - o["first"].dt.year) * 12 + (
        o["order_date"].dt.month - o["first"].dt.month
    )
    res = {}
    for coh, sub in o.groupby("cohort"):
        size = sub["customer_id"].nunique()
        row = []
        for m in range(max_m + 1):
            # skip cells not yet observable
            if (pd.Period(SNAPSHOT, "M") - coh).n < m:
                row.append(None)
            else:
                row.append(round(sub[sub["idx"] == m]["customer_id"].nunique() / size * 100, 1))
        res[str(coh)] = {"size": int(size), "ret": row}
    return res


def monthly_kpis(orders):
    o = orders.copy()
    o["m"] = pd.to_datetime(o["order_date"]).dt.to_period("M").astype(str)
    first = o.groupby("customer_id")["order_date"].transform("min")
    o["is_new"] = o["order_date"] == first
    g = o.groupby("m")
    out = pd.DataFrame(
        {
            "revenue": g["amount"].sum(),
            "orders": g.size(),
            "buyers": g["customer_id"].nunique(),
            "new_buyers": o[o["is_new"]].groupby("m")["customer_id"].nunique(),
        }
    ).fillna(0)
    out["aov"] = (out["revenue"] / out["orders"]).round()
    out["repeat_buyers"] = out["buyers"] - out["new_buyers"]
    return out.reset_index().to_dict("records")


def category_stats(orders):
    o = orders.copy()
    o["n"] = o.groupby("customer_id").cumcount()
    # repurchase rate = share of buyers of a category who bought the same category again
    rows = []
    for cat, sub in o.groupby("category"):
        buyers = sub.groupby("customer_id").size()
        rows.append(
            {
                "category": cat,
                "revenue": int(sub["amount"].sum()),
                "orders": int(len(sub)),
                "aov": int(sub["amount"].mean()),
                "repeat_rate": round(float((buyers >= 2).mean() * 100), 1),
            }
        )
    return rows


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    cfgs = load_brands()
    rng = random.Random(SEED)
    payload = {"snapshot": SNAPSHOT.isoformat(), "brands": {}}
    for key, cfg in cfgs.items():
        customers, orders = simulate_brand(key, cfg, rng)
        t = rfm_table(customers, orders, cfg["tiers"])
        # loyalty_latent is a hidden generator variable; keep it out of student files
        customers.drop(columns=["loyalty_latent"]).to_csv(OUT / f"{key}_customers.csv", index=False)
        orders.to_csv(OUT / f"{key}_orders.csv", index=False)
        t.drop(columns=["loyalty_latent"]).to_csv(OUT / f"{key}_rfm.csv", index=False)
        slim = t[
            ["customer_id", "channel", "recency", "frequency", "monetary", "aov",
             "R", "F", "M", "segment", "lifecycle", "tier"]
        ]
        payload["brands"][key] = {
            "label": cfg["label"],
            "style": cfg["style"],
            "currency": cfg["currency"],
            "tiers": cfg["tiers"],
            "customers": slim.to_dict("records"),
            "monthly": monthly_kpis(orders),
            "cohort": monthly_cohort(orders, customers),
            "categories": category_stats(orders),
        }
        print(f"{key}: {len(customers)} customers, {len(orders)} orders")
    blob = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    (OUT / "dashboard_data.json").write_text(blob, encoding="utf-8")
    template = (ROOT / "src" / "main" / "html" / "dashboard.template.html").read_text(encoding="utf-8")
    page = OUT.parent / "crm_dashboard.html"
    page.write_text(template.replace("/*__DATA__*/null", blob), encoding="utf-8")
    print(f"dashboard -> {page}")


if __name__ == "__main__":
    main()
