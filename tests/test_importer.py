import json

import pytest
from openpyxl import load_workbook

from receipt_importer import (
    build_workbook,
    compute_shares,
    is_alcohol_name,
    load_receipts,
    parse_receipt,
)
from receipt_importer.__main__ import main

GROCERY = {
    "store": "Corner Market",
    "date": "2026-10-30",
    "paid_by": "Alice",
    "items": [
        {"name": "Chips", "price": 3.00, "quantity": 2},
        {"name": "IPA 6-pack", "price": 12.00},
        {"name": "Soda", "price": 4.00, "alcohol": False},
    ],
    "tax": 2.20,
}

DINNER = {
    "store": "Taco Place",
    "date": "2026-10-31",
    "paid_by": "Bob",
    "items": [
        {"name": "Tacos", "price": 30.00},
        {"name": "Margarita", "price": 9.00, "quantity": 2, "alcohol": True},
    ],
}


@pytest.fixture
def receipt_dir(tmp_path):
    (tmp_path / "a_grocery.json").write_text(json.dumps(GROCERY))
    (tmp_path / "b_dinner.json").write_text(json.dumps([DINNER]))
    (tmp_path / "notes.txt").write_text("ignored")
    return tmp_path


@pytest.mark.parametrize(
    "name, expected",
    [
        ("IPA 6-pack", True),
        ("Red Wine", True),
        ("Craft Beers", True),
        ("Root Beer", False),
        ("Ginger Beer", False),
        ("Non-alcoholic Beer", False),
        ("Red Wine Vinegar", False),
        ("Ginger Snaps", False),
        ("Bread", False),
    ],
)
def test_is_alcohol_name(name, expected):
    assert is_alcohol_name(name) is expected


def test_parse_receipt_flags_and_tax_split():
    receipt = parse_receipt(GROCERY)
    assert [i.alcohol for i in receipt.items] == [False, True, False]
    assert receipt.subtotal(False) == 10.00
    assert receipt.subtotal(True) == 12.00
    assert receipt.tax_split() == (1.00, 1.20)
    lines = receipt.lines()
    assert [(l.name, l.total, l.alcohol) for l in lines[-2:]] == [
        ("Tax (non-alcohol share)", 1.00, False),
        ("Tax (alcohol share)", 1.20, True),
    ]


@pytest.mark.parametrize(
    "data",
    [
        {"items": []},
        {"items": [{"price": 1}]},
        {"items": [{"name": "x", "price": "1"}]},
        {"items": [{"name": "x", "price": 1, "alcohol": "yes"}]},
        {"items": [{"name": "x", "price": 1}], "tax": "1"},
        [],
    ],
)
def test_parse_receipt_rejects_invalid(data):
    with pytest.raises(ValueError):
        parse_receipt(data)


def test_load_receipts_from_directory_and_lists(receipt_dir):
    receipts = load_receipts([receipt_dir])
    assert [r.store for r in receipts] == ["Corner Market", "Taco Place"]
    assert [r.source for r in receipts] == ["a_grocery.json", "b_dinner.json"]


def test_load_receipts_invalid_json(tmp_path):
    bad = tmp_path / "bad.json"
    bad.write_text("{not json")
    with pytest.raises(ValueError, match="Invalid JSON"):
        load_receipts([bad])


def test_compute_shares_splits_alcohol_among_drinkers(receipt_dir):
    receipts = load_receipts([receipt_dir])
    shares = compute_shares(receipts, ["Alice", "Bob", "Carol"], ["Alice", "Bob"])
    # Non-alcohol: 10 + 1 tax + 30 = 41 -> 13.67 each.
    # Alcohol: 12 + 1.2 tax + 18 = 31.2 -> 15.60 per drinker.
    assert shares["Carol"] == {
        "non_alcohol": 13.67,
        "alcohol": 0.0,
        "owed": 13.67,
        "paid": 0.0,
        "balance": -13.67,
    }
    assert shares["Alice"]["alcohol"] == 15.60
    assert shares["Alice"]["paid"] == 24.20
    assert shares["Bob"]["paid"] == 48.00
    # Balances net to zero, up to per-person rounding to the cent.
    assert sum(s["balance"] for s in shares.values()) == pytest.approx(0, abs=0.015)


def test_build_workbook_sheets_and_formulas(receipt_dir):
    receipts = load_receipts([receipt_dir])
    wb = build_workbook(receipts, ["Alice", "Bob", "Carol"], ["Alice", "Bob"])
    assert wb.sheetnames == ["Items", "Receipts", "Summary"]

    items = wb["Items"]
    rows = list(items.iter_rows(min_row=2, values_only=True))
    assert len(rows) == 3 + 2 + 2  # grocery items + tax lines + dinner items
    assert rows[0][:7] == (1, "2026-10-30", "Corner Market", "Alice", "Chips", 2, 3.0)
    assert rows[0][7] == "=ROUND(F2*G2,2)"
    assert [r[8] for r in rows] == ["No", "Yes", "No", "No", "Yes", "No", "Yes"]

    receipts_ws = wb["Receipts"]
    assert receipts_ws["E2"].value == '=SUMIFS(Items!$H:$H,Items!$A:$A,A2,Items!$I:$I,"No")'
    assert receipts_ws["F2"].value == '=SUMIFS(Items!$H:$H,Items!$A:$A,A2,Items!$I:$I,"Yes")'

    summary = wb["Summary"]
    assert summary["B2"].value == '=SUMIFS(Items!$H:$H,Items!$I:$I,"No")'
    assert summary["B3"].value == '=SUMIFS(Items!$H:$H,Items!$I:$I,"Yes")'
    people = [(summary[f"A{r}"].value, summary[f"B{r}"].value) for r in range(7, 10)]
    assert people == [("Alice", "Yes"), ("Bob", "Yes"), ("Carol", "No")]
    assert summary["C7"].value == "=$B$2/COUNTA($A$7:$A$9)"
    assert summary["D9"].value == '=IF(B9="Yes",$B$3/COUNTIF($B$7:$B$9,"Yes"),0)'
    assert summary["F7"].value == "=SUMIFS(Items!$H:$H,Items!$D:$D,A7)"
    assert summary["G7"].value == "=F7-E7"
    assert summary["A10"].value == "Total"


def test_build_workbook_defaults_everyone_to_drinker(receipt_dir):
    wb = build_workbook(load_receipts([receipt_dir]), ["Alice", "Bob"])
    assert [wb["Summary"][f"B{r}"].value for r in (7, 8)] == ["Yes", "Yes"]


@pytest.mark.parametrize(
    "people, drinkers, message",
    [
        ([], [], "At least one person"),
        (["Alice", "Alice", "Bob"], None, "unique"),
        (["Alice", "Bob"], ["Zed"], "Drinkers not in people"),
        (["Alice"], ["Alice"], "paid by someone not in people: Bob"),
        (["Alice", "Bob"], [], "no drinkers"),
    ],
)
def test_build_workbook_validates_people(receipt_dir, people, drinkers, message):
    with pytest.raises(ValueError, match=message):
        build_workbook(load_receipts([receipt_dir]), people, drinkers)


def test_cli_writes_workbook(receipt_dir, tmp_path, capsys):
    output = tmp_path / "out.xlsx"
    code = main(
        [str(receipt_dir), "-o", str(output), "--people", "Alice", "Bob", "Carol",
         "--drinkers", "Alice", "Bob"]
    )
    assert code == 0
    assert load_workbook(output).sheetnames == ["Items", "Receipts", "Summary"]
    out = capsys.readouterr().out
    assert "Imported 2 receipt(s)" in out
    assert "Carol: owes 13.67" in out


def test_cli_reports_errors(tmp_path, capsys):
    code = main([str(tmp_path / "missing.json"), "--people", "Alice"])
    assert code == 1
    assert "error:" in capsys.readouterr().err
