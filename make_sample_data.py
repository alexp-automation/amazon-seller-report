"""Generate fictional Seller Central exports for the demo (a small kitchen & home brand).

    python make_sample_data.py   # writes sample_data/*.csv
"""

import csv
import random
from pathlib import Path

OUT = Path("sample_data")
PRODUCTS = [  # sku, title, price, unit_cost, inbound/unit, fba_fee, monthly units, sessions, lead time
    ("HP-BAM-CB3", "Bamboo Cutting Board Set, 3-Piece with Juice Groove", 32.99, 7.10, 1.20, 6.45, 410, 5200, 45),
    ("HP-SIL-LID6", "Silicone Stretch Lids, 6-Pack Assorted Sizes", 14.99, 2.05, 0.45, 3.86, 690, 9800, 40),
    ("HP-SS-STRW8", "Stainless Steel Straws, 8-Pack with Brushes", 9.99, 1.30, 0.30, 3.22, 520, 8300, 40),
    ("HP-GLS-JAR4", "Glass Mason Jars 32oz, Set of 4 with Lids", 27.49, 8.40, 2.60, 7.92, 240, 4100, 50),
    ("HP-OIL-SPR", "Olive Oil Sprayer for Cooking, 7oz Glass", 13.99, 2.90, 0.40, 3.86, 380, 7200, 35),
    ("HP-HRB-KEEP", "Herb Keeper Fresh Saver, Clear", 19.99, 4.60, 0.90, 5.10, 95, 3900, 45),
    ("HP-SPC-RCK", "Rotating Spice Rack Organizer, 2-Tier", 24.99, 6.80, 1.90, 6.98, 160, 3300, 50),
    ("HP-DSH-RCK", "Over-the-Sink Dish Drying Rack, Adjustable", 39.99, 12.50, 3.80, 9.41, 130, 4600, 55),
    ("HP-MSR-SET", "Magnetic Measuring Spoons Set of 8", 16.99, 3.10, 0.50, 3.86, 300, 4100, 40),
    ("HP-BEE-WRP", "Beeswax Food Wraps, 9-Pack", 18.99, 3.70, 0.60, 3.86, 210, 6900, 40),
    ("HP-CRM-MUG2", "Ceramic Coffee Mugs 16oz, Set of 2", 22.99, 6.20, 2.40, 7.92, 70, 2800, 60),
    ("HP-KNF-MAG", "Magnetic Knife Strip 16 inch, Walnut", 29.99, 7.80, 1.70, 6.45, 115, 2400, 50),
]
STOCK = [180, 2400, 1900, 70, 1100, 720, 410, 120, 950, 300, 900, 230]   # fulfillable
INBOUND = [600, 0, 0, 300, 0, 0, 0, 200, 0, 400, 0, 0]
PPC = [(980, 3900), (1150, 6200), (420, 2900), (610, 1900), (520, 2700), (690, 980),
       (300, 1500), (440, 1800), (310, 2100), (520, 1400), (260, 520), (210, 1100)]


def write(name, header, rows):
    with open(OUT / name, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def main():
    OUT.mkdir(exist_ok=True)
    rnd = random.Random(5)
    business, inventory, ads, costs = [], [], [], []
    for i, (sku, title, price, cost, inbound, fba, units, sessions, lead) in enumerate(PRODUCTS):
        asin = f"B0{rnd.randint(10**7, 10**8 - 1)}"
        units = int(units * rnd.uniform(0.9, 1.1))
        weak = sku in ("HP-DSH-RCK", "HP-HRB-KEEP", "HP-CRM-MUG2")  # listings with problems
        sessions = int(sessions * (1.0 if weak else 0.55) * rnd.uniform(0.9, 1.1))
        sales = round(units * price * rnd.uniform(0.93, 0.99), 2)
        business.append([asin, asin, title, sku, f"{sessions:,}", f"{int(sessions * 1.4):,}",
                         f"{units / sessions:.2%}", units, f"${sales:,.2f}", units])
        inventory.append([sku, f"X00{rnd.randint(10**6, 10**7 - 1)}", asin, title, "New", f"{price:.2f}",
                          STOCK[i], rnd.randint(5, 40), INBOUND[i], 0, rnd.randint(0, 6)])
        spend, ad_sales = PPC[i]
        ads.append([sku, asin, f"{spend * 12:,}", int(spend * 1.1), f"${spend:,.2f}", f"${ad_sales:,.2f}"])
        costs.append([sku, f"{cost:.2f}", f"{inbound:.2f}", f"{fba:.2f}", "0.15", lead])

    write("BusinessReport-ByChildASIN.csv",
          ["(Parent) ASIN", "(Child) ASIN", "Title", "SKU", "Sessions - Total", "Page Views - Total",
           "Unit Session Percentage", "Units Ordered", "Ordered Product Sales", "Total Order Items"], business)
    write("FBA_Inventory.csv",
          ["sku", "fnsku", "asin", "product-name", "condition", "your-price", "afn-fulfillable-quantity",
           "afn-reserved-quantity", "afn-inbound-shipped-quantity", "afn-inbound-receiving-quantity",
           "afn-unsellable-quantity"], inventory)
    write("SponsoredProducts_AdvertisedProduct.csv",
          ["Advertised SKU", "Advertised ASIN", "Impressions", "Clicks", "Spend", "7 Day Total Sales"], ads)
    write("product_costs.csv",
          ["sku", "unit_cost", "inbound_cost_per_unit", "fba_fee", "referral_rate", "lead_time_days"], costs)
    print(f"Wrote 4 files to {OUT}/")


if __name__ == "__main__":
    main()
