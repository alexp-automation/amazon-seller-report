"""Amazon seller report: true profit per SKU, restock plan and listing red flags.

Inputs are the standard Seller Central downloads (column names are mapped at the top,
so a renamed column is a one-line fix) plus one small file the seller keeps up to date:
  BusinessReport-ByChildASIN.csv          Reports > Business Reports > By ASIN > Detail Page Sales and Traffic by Child Item
  FBA_Inventory.csv                       Reports > Fulfillment > Inventory
  SponsoredProducts_AdvertisedProduct.csv Advertising > Reports > Sponsored Products > Advertised product
  product_costs.csv                       sku, unit_cost, inbound_cost_per_unit, fba_fee, referral_rate, lead_time_days

    python report.py --days 30 --as-of 2026-10-06
"""

import argparse
import csv
import math
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BUSINESS = {"sku": "SKU", "title": "Title", "sessions": "Sessions - Total",
            "units": "Units Ordered", "sales": "Ordered Product Sales"}
INVENTORY = {"sku": "sku", "fulfillable": "afn-fulfillable-quantity",
             "inbound_shipped": "afn-inbound-shipped-quantity", "inbound_receiving": "afn-inbound-receiving-quantity"}
ADS = {"sku": "Advertised SKU", "spend": "Spend", "sales": "7 Day Total Sales"}

SETTINGS = [("Report period, days", 30), ("As-of date", None), ("Safety stock, days", 14),
            ("Target cover after arrival, days", 60), ("Overstock above, days of cover", 120),
            ("Low conversion below", 0.08), ("…with sessions above", 2000)]

NAVY, YELLOW = "1F3A5F", "FFF7D6"
HEAD_FONT, HEAD_FILL = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=NAVY)
MONEY, PCT = '"$"#,##0', "0.0%"


def num(value):
    """'$1,234.56' / '12.5%' / '1,024' / '' -> float"""
    s = str(value or "").replace("$", "").replace(",", "").replace("%", "").strip()
    if not s:
        return 0.0
    return float(s) / (100 if "%" in str(value) else 1)


def read(path, mapping):
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        missing = [col for col in mapping.values() if col not in reader.fieldnames]
        if missing:
            raise SystemExit(f"{path.name}: missing columns {missing}. Update the mapping at the top of report.py.")
        return {row[mapping["sku"]].strip(): {k: row[v] for k, v in mapping.items()} for row in reader}


def load(data: Path, days: int):
    business = read(data / "BusinessReport-ByChildASIN.csv", BUSINESS)
    inventory = read(data / "FBA_Inventory.csv", INVENTORY)
    ads = read(data / "SponsoredProducts_AdvertisedProduct.csv", ADS)
    costs = read(data / "product_costs.csv", {"sku": "sku", "unit_cost": "unit_cost",
                                              "inbound": "inbound_cost_per_unit", "fba_fee": "fba_fee",
                                              "referral": "referral_rate", "lead": "lead_time_days"})
    rows, warnings = [], []
    for sku in sorted(set(business) | set(inventory)):
        if sku not in costs:
            warnings.append(f"{sku}: no row in product_costs.csv, skipped")
            continue
        b, inv, ad, c = business.get(sku, {}), inventory.get(sku, {}), ads.get(sku, {}), costs[sku]
        units, revenue = num(b.get("units")), num(b.get("sales"))
        r = {
            "sku": sku, "title": b.get("title", ""), "sessions": num(b.get("sessions")),
            "units": units, "revenue": revenue, "referral": num(c["referral"]), "fba_fee": num(c["fba_fee"]),
            "cogs_unit": num(c["unit_cost"]) + num(c["inbound"]), "ppc_spend": num(ad.get("spend")),
            "ppc_sales": num(ad.get("sales")), "fulfillable": num(inv.get("fulfillable")),
            "inbound": num(inv.get("inbound_shipped")) + num(inv.get("inbound_receiving")),
            "lead": num(c["lead"]),
        }
        r["net"] = revenue - revenue * r["referral"] - units * r["fba_fee"] - units * r["cogs_unit"] - r["ppc_spend"]
        r["velocity"] = units / days if days else 0
        r["cover"] = (r["fulfillable"] + r["inbound"]) / r["velocity"] if r["velocity"] else math.inf
        r["fba_cover"] = r["fulfillable"] / r["velocity"] if r["velocity"] else math.inf
        rows.append(r)
    rows.sort(key=lambda r: -r["net"])
    return rows, warnings


def header(ws, cols, widths):
    ws.append(cols)
    for i, cell in enumerate(ws[1]):
        cell.font, cell.fill = HEAD_FONT, HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i + 1)].width = widths[i]
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "C2"


def attention_list(rows, s):
    """Plain-language to-do list, same thresholds as the sheet formulas."""
    out = []
    safety = s["Safety stock, days"]
    for r in sorted(rows, key=lambda r: r["fba_cover"]):
        days_left = r["cover"] - r["lead"] - safety      # days until the order must be placed
        if days_left <= 0 and r["fba_cover"] < r["lead"]:
            out.append(("Stockout risk", r["sku"], f"FBA stock lasts ~{r['fba_cover']:.0f} days, "
                                                    f"reorder lead time is {r['lead']:.0f} days. Order now."))
        elif days_left <= 7:
            when = "today" if days_left <= 0 else f"within {days_left:.0f} days"
            out.append(("Reorder soon", r["sku"], f"~{r['cover']:.0f} days of stock incl. inbound; "
                                                  f"place the next order {when}."))
    for r in rows:
        if r["net"] < 0:
            out.append(("Losing money", r["sku"], f"Net -${-r['net']:,.0f} this period; ad spend "
                                                  f"${r['ppc_spend']:,.0f} (TACoS {r['ppc_spend'] / r['revenue']:.0%})."))
    for r in rows:
        cvr = r["units"] / r["sessions"] if r["sessions"] else 0
        if cvr < s["Low conversion below"] and r["sessions"] > s["…with sessions above"]:
            out.append(("Low conversion", r["sku"], f"{cvr:.1%} of {r['sessions']:,.0f} sessions buy. "
                                                    f"Check price, images, reviews."))
    for r in rows:
        if r["cover"] > s["Overstock above, days of cover"]:
            out.append(("Overstock", r["sku"], f"{r['cover']:.0f} days of stock: storage fees add up. "
                                               f"Consider a promotion."))
    return out


def build(rows, warnings, settings, path):
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    profit = wb.create_sheet("Profit by SKU")
    restock = wb.create_sheet("Restock plan")
    st = wb.create_sheet("Settings")

    # Settings (editable)
    st.column_dimensions["A"].width, st.column_dimensions["B"].width = 34, 14
    for i, (label, value) in enumerate(SETTINGS, start=1):
        st.cell(row=i, column=1, value=label)
        cell = st.cell(row=i, column=2, value=settings[label])
        cell.fill = PatternFill("solid", fgColor=YELLOW)
    st["B2"].number_format = "mm/dd/yyyy"
    st["B6"].number_format = PCT
    st["A9"] = "Yellow cells on every sheet are inputs. Change them and all numbers recalculate."
    st["A9"].font = Font(italic=True, color="6B7280")

    # Profit by SKU: inputs as values, results as formulas
    header(profit, ["SKU", "Product", "Sessions", "Units", "Revenue", "Referral %", "FBA fee / unit",
                    "COGS / unit", "PPC spend", "PPC sales", "Amazon fees", "COGS", "Net profit", "Margin",
                    "Profit / unit", "ACoS", "TACoS", "Conversion"],
           [13, 34, 9, 7, 11, 9, 9, 9, 10, 10, 11, 10, 11, 8, 9, 8, 8, 10])
    for i, r in enumerate(rows, start=2):
        profit.append([r["sku"], r["title"], r["sessions"], r["units"], r["revenue"], r["referral"],
                       r["fba_fee"], r["cogs_unit"], r["ppc_spend"], r["ppc_sales"],
                       f"=E{i}*F{i}+D{i}*G{i}", f"=D{i}*H{i}", f"=E{i}-K{i}-L{i}-I{i}",
                       f"=IFERROR(M{i}/E{i},0)", f"=IFERROR(M{i}/D{i},0)", f"=IFERROR(I{i}/J{i},0)",
                       f"=IFERROR(I{i}/E{i},0)", f"=IFERROR(D{i}/C{i},0)"])
        for col in "FGH":
            profit[f"{col}{i}"].fill = PatternFill("solid", fgColor=YELLOW)
    last = len(rows) + 1
    t = last + 1
    profit.append(["Total", "", f"=SUM(C2:C{last})", f"=SUM(D2:D{last})", f"=SUM(E2:E{last})", "", "", "",
                   f"=SUM(I2:I{last})", f"=SUM(J2:J{last})", f"=SUM(K2:K{last})", f"=SUM(L2:L{last})",
                   f"=SUM(M2:M{last})", f"=IFERROR(M{t}/E{t},0)", f"=IFERROR(M{t}/D{t},0)",
                   f"=IFERROR(I{t}/J{t},0)", f"=IFERROR(I{t}/E{t},0)", f"=IFERROR(D{t}/C{t},0)"])
    for cell in profit[t]:
        cell.font = Font(bold=True)
    for row in profit.iter_rows(min_row=2, max_row=t):
        for c in ("E", "I", "J", "K", "L", "M"):
            row[ord(c) - 65].number_format = MONEY
        for c in ("F", "N", "P", "Q", "R"):
            row[ord(c) - 65].number_format = PCT
        for c in ("G", "H", "O"):
            row[ord(c) - 65].number_format = '"$"0.00'
        row[2].number_format = row[3].number_format = "#,##0"
    profit.conditional_formatting.add(f"M2:M{last}", CellIsRule(
        operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="F8D7DA"), font=Font(color="9B1C1C", bold=True)))
    profit.conditional_formatting.add(f"R2:R{last}", FormulaRule(
        formula=[f"AND(R2<Settings!$B$6,C2>Settings!$B$7)"], fill=PatternFill("solid", fgColor="FDE8C8")))

    # Restock plan
    header(restock, ["SKU", "Product", "Units / day", "FBA fulfillable", "Inbound", "Days of cover (FBA)",
                     "Days of cover (incl. inbound)", "Lead time, days", "Runs out (FBA)", "Order by (incl. inbound)",
                     "Suggested order qty", "Status"],
           [13, 34, 9, 10, 9, 11, 12, 9, 12, 12, 11, 22])
    for i, r in enumerate(sorted(rows, key=lambda r: r["fba_cover"]), start=2):
        restock.append([r["sku"], r["title"], f"='Profit by SKU'!D{rows.index(r) + 2}/Settings!$B$1",
                        r["fulfillable"], r["inbound"], f"=IFERROR(D{i}/C{i},999)", f"=IFERROR((D{i}+E{i})/C{i},999)",
                        r["lead"], f"=IFERROR(Settings!$B$2+D{i}/C{i},\"\")",
                        f"=IFERROR(MAX(Settings!$B$2,Settings!$B$2+G{i}-H{i}-Settings!$B$3),\"\")",
                        f"=IF(OR(LEFT(L{i},9)=\"ORDER NOW\",L{i}=\"Order this week\"),"
                        f"MAX(0,CEILING((H{i}+Settings!$B$3+Settings!$B$4)*C{i}-D{i}-E{i},10)),0)",
                        # days until the order must be placed = G - H - safety
                        f"=IF(G{i}-H{i}-Settings!$B$3<=0,IF(F{i}<H{i},\"ORDER NOW: stockout risk\",\"ORDER NOW\"),"
                        f"IF(G{i}-H{i}-Settings!$B$3<=7,\"Order this week\",IF(G{i}>Settings!$B$5,\"Overstock\","
                        f"IF(F{i}<H{i},\"Low FBA, inbound covers\",\"OK\"))))"])
        restock[f"H{i}"].fill = PatternFill("solid", fgColor=YELLOW)
        restock[f"C{i}"].number_format = "0.0"
        restock[f"F{i}"].number_format = restock[f"G{i}"].number_format = "0"
        restock[f"I{i}"].number_format = restock[f"J{i}"].number_format = "mm/dd/yyyy"
        restock[f"K{i}"].number_format = "#,##0"
    rl = len(rows) + 1
    for text, color in [("ORDER NOW", "F8D7DA"), ("Order this week", "FDE8C8"), ("Low FBA", "FFF3C4"),
                        ("Overstock", "DBEAFE"), ("OK", "D4EDDA")]:
        restock.conditional_formatting.add(f"L2:L{rl}", FormulaRule(
            formula=[f'LEFT($L2,{len(text)})="{text}"'], fill=PatternFill("solid", fgColor=color)))

    # Summary
    summary.sheet_view.showGridLines = False
    summary["B2"] = "Amazon seller report"
    summary["B2"].font = Font(size=20, bold=True, color=NAVY)
    summary["B3"] = f"=\"Last \"&Settings!B1&\" days, as of \"&TEXT(Settings!B2,\"mmm d, yyyy\")"
    summary["B3"].font = Font(color="6B7280")
    kpis = [("Revenue", f"='Profit by SKU'!E{t}", MONEY), ("Net profit", f"='Profit by SKU'!M{t}", MONEY),
            ("Margin", f"='Profit by SKU'!N{t}", PCT), ("Ad spend", f"='Profit by SKU'!I{t}", MONEY),
            ("TACoS", f"='Profit by SKU'!Q{t}", PCT),
            ("SKUs to reorder now", f"=COUNTIF('Restock plan'!L2:L{rl},\"ORDER NOW*\")", "0")]
    for k, (label, formula, fmt) in enumerate(kpis):
        col = 2 + k * 2
        summary.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)
        summary.merge_cells(start_row=6, start_column=col, end_row=6, end_column=col + 1)
        a = summary.cell(row=5, column=col, value=label)
        v = summary.cell(row=6, column=col, value=formula)
        a.font, v.font = Font(color="6B7280", size=10), Font(bold=True, size=17, color=NAVY)
        a.alignment = v.alignment = Alignment(horizontal="center")
        v.number_format = fmt
        for rr in (5, 6):
            for cc in (col, col + 1):
                summary.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor="EEF2F7")
    summary.row_dimensions[6].height = 28

    summary["B9"] = "Needs attention"
    summary["B9"].font = Font(bold=True, size=13, color=NAVY)
    colors = {"Stockout risk": "F8D7DA", "Reorder soon": "FDE8C8", "Losing money": "F8D7DA", "Low conversion": "FDE8C8", "Overstock": "DBEAFE"}
    items = attention_list(rows, settings)
    for j, (kind, sku, text) in enumerate(items or [("All good", "", "Nothing urgent this period.")]):
        r = 10 + j
        summary.cell(row=r, column=2, value=kind).fill = PatternFill("solid", fgColor=colors.get(kind, "D4EDDA"))
        summary.cell(row=r, column=3, value=sku).font = Font(bold=True)
        summary.cell(row=r, column=5, value=text)
    for col, w in zip("ABCDEFGHIJKLM", [2, 15, 13, 3, 12, 12, 12, 12, 12, 12, 12, 12, 12]):
        summary.column_dimensions[col].width = w

    chart = BarChart()
    chart.type = "bar"
    chart.title = "Net profit by SKU, this period"
    chart.add_data(Reference(profit, min_col=13, min_row=1, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(profit, min_col=1, min_row=2, max_row=last))
    chart.y_axis.numFmt = '"$"#,##0'
    chart.x_axis.scaling.orientation = "maxMin"
    chart.legend = None
    chart.x_axis.delete = chart.y_axis.delete = False
    chart.height, chart.width = 9, 17
    summary.add_chart(chart, f"B{12 + len(items)}")

    if warnings:
        st["A11"] = "Warnings"
        st["A11"].font = Font(bold=True)
        for j, w in enumerate(warnings):
            st.cell(row=12 + j, column=1, value=w)
    wb.save(path)


def main():
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("--data", default="sample_data")
    p.add_argument("--days", type=int, default=30, help="length of the Business Report period")
    p.add_argument("--as-of", default=None, help="report date YYYY-MM-DD (default today)")
    p.add_argument("--out", default="Amazon_Seller_Report.xlsx")
    args = p.parse_args()

    settings = dict(SETTINGS)
    settings["Report period, days"] = args.days
    settings["As-of date"] = date.fromisoformat(args.as_of) if args.as_of else date.today()
    rows, warnings = load(Path(args.data), args.days)
    build(rows, warnings, settings, args.out)
    total = sum(r["net"] for r in rows)
    print(f"{len(rows)} SKUs, net profit ${total:,.0f}. "
          f"{len(attention_list(rows, settings))} items need attention -> {args.out}")
    for w in warnings:
        print("warning:", w)


if __name__ == "__main__":
    main()
