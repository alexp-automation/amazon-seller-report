"""One-page HTML summary of the same report, for email or a phone screen.

    python html_summary.py --as-of 2026-10-06    # writes Amazon_Seller_Summary.html
"""

import argparse
import html
import math
from datetime import date, timedelta
from pathlib import Path

from report import SETTINGS, attention_list, load

CSS = """
body{margin:0;background:#f4f5f7;color:#111827;font:14px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",Inter,sans-serif}
.wrap{max-width:1180px;margin:0 auto;padding:28px 20px}
h1{margin:0;font-size:24px;color:#1F3A5F} .sub{color:#6b7280;margin:2px 0 20px}
.kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:22px}
.kpi{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:12px 14px}
.kpi span{display:block;color:#6b7280;font-size:12px}.kpi b{font-size:22px;color:#1F3A5F}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.card{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:14px 16px}
h2{font-size:15px;margin:0 0 10px}
.item{display:grid;grid-template-columns:110px 105px 1fr;gap:8px;padding:6px 0;border-top:1px solid #f0f1f3;font-size:13px}
.tag{border-radius:6px;padding:1px 7px;font-size:11px;font-weight:600;text-transform:uppercase;align-self:start;text-align:center}
.red{background:#fde2e2;color:#9b1c1c}.amber{background:#fdecc8;color:#8a4b08}.blue{background:#dbeafe;color:#1e40af}.green{background:#d4edda;color:#1e5b2f}
.bar{display:grid;grid-template-columns:105px 1fr 70px;gap:8px;align-items:center;font-size:12px;margin:5px 0}
.track{height:14px;position:relative}.fill{position:absolute;top:0;height:14px;border-radius:3px}
.pos{background:#3b6fd8}.neg{background:#d64545}
.num{text-align:right;font-variant-numeric:tabular-nums}
table{width:100%;border-collapse:collapse;font-size:13px}
th{background:#1F3A5F;color:#fff;font-weight:600;padding:7px 8px;text-align:left;font-size:12px}
td{padding:7px 8px;border-bottom:1px solid #eef0f3}
"""


def money(v):
    return f"-${-v:,.0f}" if v < 0 else f"${v:,.0f}"


def status(r, s):
    safety, over = s["Safety stock, days"], s["Overstock above, days of cover"]
    days_left = r["cover"] - r["lead"] - safety          # same rule as the Excel formula
    if days_left <= 0:
        return ("ORDER NOW: stockout risk", "red") if r["fba_cover"] < r["lead"] else ("ORDER NOW", "red")
    if days_left <= 7:
        return "Order this week", "amber"
    if r["cover"] > over:
        return "Overstock", "blue"
    return ("Low FBA, inbound covers", "amber") if r["fba_cover"] < r["lead"] else ("OK", "green")


def render(rows, s):
    e = html.escape
    as_of = s["As-of date"]
    rev = sum(r["revenue"] for r in rows)
    net = sum(r["net"] for r in rows)
    spend = sum(r["ppc_spend"] for r in rows)
    to_order = sum(1 for r in rows if status(r, s)[0].startswith("ORDER NOW"))
    kpis = [("Revenue", money(rev)), ("Net profit", money(net)), ("Margin", f"{net / rev:.1%}"),
            ("Ad spend", money(spend)), ("TACoS", f"{spend / rev:.1%}"), ("SKUs to reorder now", str(to_order))]
    tag = {"Stockout risk": "red", "Losing money": "red", "Reorder soon": "amber",
           "Low conversion": "amber", "Overstock": "blue"}
    items = "".join(f"<div class='item'><span class='tag {tag[k]}'>{e(k)}</span><b>{e(sku)}</b><span>{e(t)}</span></div>"
                    for k, sku, t in attention_list(rows, s))
    top = max(abs(r["net"]) for r in rows)
    zero = 100 * min(r["net"] for r in rows) / -top if min(r["net"] for r in rows) < 0 else 0
    zero = max(zero, 0) / (1 + max(zero, 0) / 100)
    bars = ""
    for r in rows:
        w = abs(r["net"]) / top * (100 - zero)
        left = zero if r["net"] >= 0 else zero - w
        bars += (f"<div class='bar'><span>{e(r['sku'])}</span><div class='track'><div class='fill "
                 f"{'pos' if r['net'] >= 0 else 'neg'}' style='left:{left:.1f}%;width:{w:.1f}%'></div></div>"
                 f"<span class='num'>{money(r['net'])}</span></div>")
    restock = ""
    for r in sorted(rows, key=lambda r: r["fba_cover"]):
        label, color = status(r, s)
        v = r["velocity"]
        runs_out = as_of + timedelta(days=r["fba_cover"]) if v else None
        order_by = max(as_of, as_of + timedelta(days=min(r["cover"], 999) - r["lead"] - s["Safety stock, days"]))
        qty = 0
        if label.startswith("ORDER NOW") or label == "Order this week":
            qty = max(0, math.ceil(((r["lead"] + s["Safety stock, days"] + s["Target cover after arrival, days"]) * v
                                    - r["fulfillable"] - r["inbound"]) / 10) * 10)
        restock += (f"<tr><td><b>{e(r['sku'])}</b></td><td class='num'>{v:.1f}</td><td class='num'>{r['fulfillable']:,.0f}</td>"
                    f"<td class='num'>{r['inbound']:,.0f}</td><td class='num'>{min(r['fba_cover'], 999):.0f}</td>"
                    f"<td class='num'>{min(r['cover'], 999):.0f}</td>"
                    f"<td class='num'>{r['lead']:.0f}</td><td>{runs_out:%b %d}</td>"
                    f"<td>{'today' if order_by == as_of else format(order_by, '%b %d')}</td>"
                    f"<td class='num'>{qty:,}</td><td><span class='tag {color}'>{e(label)}</span></td></tr>")
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Amazon seller summary</title>
<style>{CSS}</style></head><body><div class="wrap">
<h1>Amazon seller report</h1><p class="sub">Last {s['Report period, days']} days · as of {as_of:%b %d, %Y} · {len(rows)} SKUs</p>
<div class="kpis">{''.join(f"<div class='kpi'><span>{k}</span><b>{v}</b></div>" for k, v in kpis)}</div>
<div class="grid"><div class="card"><h2>Needs attention</h2>{items}</div>
<div class="card"><h2>Net profit by SKU (after Amazon fees, COGS and ads)</h2>{bars}</div></div>
<div class="card" style="margin-top:16px"><h2>Restock plan</h2><table>
<tr><th>SKU</th><th>Units/day</th><th>FBA stock</th><th>Inbound</th><th>FBA days</th><th>Days incl. inbound</th><th>Lead time</th>
<th>FBA runs out</th><th>Order by</th><th>Order qty</th><th>Status</th></tr>{restock}</table></div>
</div></body></html>"""


def main():
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("--data", default="sample_data")
    p.add_argument("--days", type=int, default=30)
    p.add_argument("--as-of", default=None)
    p.add_argument("--out", default="Amazon_Seller_Summary.html")
    args = p.parse_args()
    s = dict(SETTINGS)
    s["Report period, days"] = args.days
    s["As-of date"] = date.fromisoformat(args.as_of) if args.as_of else date.today()
    rows, _ = load(Path(args.data), args.days)
    Path(args.out).write_text(render(rows, s), encoding="utf-8")
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
