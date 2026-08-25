from pathlib import Path

import pandas as pd

from src.audit_online_retail import audit_workbook


def test_synthetic_workbook_covers_rules_and_aggregation(tmp_path: Path) -> None:
    rows = [
        ["100", "10000", "normal", 1, "2020-01-01", 2.0, 1, "UK"],
        ["101", "10000", "repeat item", 2, "2020-01-01", 2.0, 1, "UK"],
        ["102", "10001A", "suffix", 1, "2020-01-01", 1.0, 1, "UK"],
        ["103", "10002", "second user", 1, "2020-01-01", 1.0, 2, "UK"],
        ["C104", "10003", "cancelled", 1, "2020-01-01", 1.0, 3, "UK"],
        ["105", "10004", "return", -1, "2020-01-01", 1.0, 3, "UK"],
        ["106", "10005", "free", 1, "2020-01-01", 0.0, 4, "UK"],
        ["107", "POST", "service", 1, "2020-01-01", 1.0, 5, "UK"],
        ["108", None, "blank stock", 1, "2020-01-01", 1.0, 6, "UK"],
        ["109", "10006", "missing customer", 1, "2020-01-01", 1.0, None, "UK"],
        ["110", "10007", "fractional customer", 1, "2020-01-01", 1.0, 7.5, "UK"],
        ["111", " 10008b ", "normalized suffix", 1, "2020-01-01", 1.0, 8, "UK"],
        ["112", "10009", "priority customer first", -1, "2020-01-01", 1.0, None, "UK"],
    ]
    columns = ["InvoiceNo", "StockCode", "Description", "Quantity", "InvoiceDate", "UnitPrice", "CustomerID", "Country"]
    path = tmp_path / "synthetic.xlsx"
    pd.DataFrame(rows, columns=columns).to_excel(path, index=False, engine="openpyxl")

    first = audit_workbook(path)
    second = audit_workbook(path)
    assert first["file"]["sha256"] == second["file"]["sha256"]
    assert first["file"]["unchanged_during_audit"]
    assert first["filtering"]["all_rows_accounted_for"]
    assert first["filtering"]["accounting_total"] == len(rows)
    assert first["filtering"]["unique_primary_reason_counts"] == {
        "invalid_customer_id": 3,
        "cancelled_invoice": 1,
        "nonpositive_quantity": 1,
        "nonpositive_unit_price": 1,
        "invalid_stock_code": 2,
        "kept": 5,
    }
    agg = first["aggregation_preview"]
    assert agg["users_after_row_filters"] == 3
    assert agg["items_after_row_filters"] == 4
    assert agg["duplicate_user_item_rows_removed_by_basket_deduplication"] == 1
    assert agg["basket_length"]["min"] == 1
    assert agg["basket_length"]["max"] == 2
    assert agg["all_users_unique"] and agg["all_baskets_nonempty"] and agg["all_baskets_deduplicated"]
    assert first["mapping_preview"]["first_10_forward_entries"][-1]["stock_code"] == "10008B"
    assert not any(first["prohibited_outputs"].values())

