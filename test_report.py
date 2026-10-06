import csv

import pytest

from report import SETTINGS, attention_list, load, num


def write(path, header, rows):
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


@pytest.fixture
def data(tmp_path):
    write(tmp_path / "BusinessReport-ByChildASIN.csv",
          ["SKU", "Title", "Sessions - Total", "Units Ordered", "Ordered Product Sales"],
          [["A", "Board", "1,000", "100", "$2,000.00"], ["B", "Jar", "5,000", "60", "$900.00"]])
    write(tmp_path / "FBA_Inventory.csv",
          ["sku", "afn-fulfillable-quantity", "afn-inbound-shipped-quantity", "afn-inbound-receiving-quantity"],
          [["A", "50", "100", "0"], ["B", "900", "0", "0"]])
    write(tmp_path / "SponsoredProducts_AdvertisedProduct.csv", ["Advertised SKU", "Spend", "7 Day Total Sales"],
          [["A", "$200.00", "$800.00"], ["B", "$500.00", "$300.00"]])
    write(tmp_path / "product_costs.csv",
          ["sku", "unit_cost", "inbound_cost_per_unit", "fba_fee", "referral_rate", "lead_time_days"],
          [["A", "5", "1", "4", "0.15", "40"], ["B", "6", "2", "5", "0.15", "40"]])
    return tmp_path


def test_num_parses_seller_central_formats():
    assert num("$1,234.56") == 1234.56
    assert num("12.5%") == pytest.approx(0.125)
    assert num("") == 0


def test_profit_and_cover(data):
    rows, warnings = load(data, 30)
    a = next(r for r in rows if r["sku"] == "A")
    # 2000 - 15% referral 300 - FBA 400 - COGS 600 - PPC 200 = 500
    assert a["net"] == pytest.approx(500)
    assert a["velocity"] == pytest.approx(100 / 30)
    assert a["fba_cover"] == pytest.approx(15)         # 50 units at 3.33/day
    assert rows[0]["sku"] == "A"                        # sorted by profit
    assert not warnings


def test_attention_list_flags(data):
    rows, _ = load(data, 30)
    s = dict(SETTINGS)
    kinds = {(k, sku) for k, sku, _ in attention_list(rows, s)}
    assert ("Stockout risk", "A") in kinds              # 15 days of FBA stock < 40 day lead time
    assert ("Losing money", "B") in kinds               # 900 - 135 - 300 - 480 - 500 < 0
    assert ("Low conversion", "B") in kinds             # 1.2% of 5,000 sessions
    assert ("Overstock", "B") in kinds                  # 900 units at 2/day = 450 days


def test_missing_cost_row_is_a_warning(data):
    with open(data / "product_costs.csv") as f:
        lines = f.readlines()
    (data / "product_costs.csv").write_text("".join(lines[:2]))
    rows, warnings = load(data, 30)
    assert [r["sku"] for r in rows] == ["A"] and "B" in warnings[0]


def test_restock_status_rule():
    from html_summary import status
    s = dict(SETTINGS)  # safety 14, overstock 120

    def row(fba, total, lead=40):
        return {"fba_cover": fba, "cover": total, "lead": lead}
    assert status(row(10, 50), s)[0] == "ORDER NOW: stockout risk"   # 50-40-14 < 0, FBA runs dry first
    assert status(row(45, 50), s)[0] == "ORDER NOW"                  # late, but FBA lasts the lead time
    assert status(row(60, 60), s)[0] == "Order this week"            # 6 days left
    assert status(row(20, 90), s)[0] == "Low FBA, inbound covers"
    assert status(row(200, 200), s)[0] == "Overstock"
    assert status(row(80, 80), s)[0] == "OK"
