"""Core logic for importing receipt JSON files into an Excel workbook.

Each receipt JSON file contains a single receipt object or a list of them::

    {
        "store": "Corner Market",
        "date": "2026-10-30",
        "paid_by": "Alice",
        "items": [
            {"name": "Chips", "price": 3.50, "quantity": 2},
            {"name": "IPA 6-pack", "price": 11.99, "alcohol": true}
        ],
        "tax": 1.25
    }

Items are split into alcohol and non-alcohol costs. Non-alcohol costs are
shared by every person, alcohol costs only by the people who drink. Items
without an explicit ``alcohol`` flag are classified by keyword.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Sequence

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

MONEY_FORMAT = "#,##0.00"

ALCOHOL_KEYWORDS = (
    "alcohol",
    "beer",
    "bourbon",
    "brandy",
    "champagne",
    "cider",
    "cognac",
    "gin",
    "ipa",
    "lager",
    "liquor",
    "merlot",
    "pinot",
    "prosecco",
    "rum",
    "sake",
    "scotch",
    "stout",
    "tequila",
    "vodka",
    "whiskey",
    "whisky",
    "wine",
)
_ALCOHOL_RE = re.compile(r"\b(" + "|".join(ALCOHOL_KEYWORDS) + r")s?\b", re.IGNORECASE)
_NON_ALCOHOL_RE = re.compile(
    r"\b(root|ginger) beers?\b|\bnon[- ]?alcoholic\b|\balcohol[- ]free\b|\bvinegar\b",
    re.IGNORECASE,
)

ITEM_HEADERS = (
    "Receipt #",
    "Date",
    "Store",
    "Paid By",
    "Item",
    "Quantity",
    "Unit Price",
    "Total",
    "Alcohol",
    "Source File",
)


@dataclass
class Item:
    name: str
    price: float
    quantity: float = 1
    alcohol: bool = False

    @property
    def total(self) -> float:
        return round(self.price * self.quantity, 2)


@dataclass
class Receipt:
    store: str
    date: str
    paid_by: str | None
    items: list[Item] = field(default_factory=list)
    tax: float = 0.0
    source: str = ""

    def subtotal(self, alcohol: bool | None = None) -> float:
        return round(
            sum(i.total for i in self.items if alcohol is None or i.alcohol == alcohol),
            2,
        )

    def tax_split(self) -> tuple[float, float]:
        """Allocate tax proportionally, returning ``(non_alcohol, alcohol)``."""
        subtotal = self.subtotal()
        alcohol_tax = round(self.tax * self.subtotal(True) / subtotal, 2) if subtotal else 0.0
        return round(self.tax - alcohol_tax, 2), alcohol_tax

    def lines(self) -> list[Item]:
        """Items plus tax lines, so every cost is tagged alcohol or not."""
        lines = list(self.items)
        non_alcohol_tax, alcohol_tax = self.tax_split()
        if non_alcohol_tax:
            lines.append(Item("Tax (non-alcohol share)", non_alcohol_tax))
        if alcohol_tax:
            lines.append(Item("Tax (alcohol share)", alcohol_tax, alcohol=True))
        return lines


def is_alcohol_name(name: str) -> bool:
    """Guess whether an item is alcoholic from its name."""
    return bool(_ALCOHOL_RE.search(name)) and not _NON_ALCOHOL_RE.search(name)


def _number(value, what: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{what} must be a number, got {value!r}")
    return float(value)


def parse_receipt(data: dict, source: str = "") -> Receipt:
    where = f" in {source}" if source else ""
    if not isinstance(data, dict):
        raise ValueError(f"Receipt must be a JSON object{where}")
    raw_items = data.get("items")
    if not isinstance(raw_items, list) or not raw_items:
        raise ValueError(f"Receipt must have a non-empty 'items' list{where}")

    items = []
    for index, raw in enumerate(raw_items, start=1):
        if not isinstance(raw, dict) or not raw.get("name"):
            raise ValueError(f"Item {index} must be an object with a 'name'{where}")
        name = str(raw["name"])
        alcohol = raw.get("alcohol")
        if alcohol is None:
            alcohol = is_alcohol_name(name)
        elif not isinstance(alcohol, bool):
            raise ValueError(f"Item {name!r} 'alcohol' must be true or false{where}")
        items.append(
            Item(
                name=name,
                price=_number(raw.get("price"), f"Item {name!r} price{where}"),
                quantity=_number(raw.get("quantity", 1), f"Item {name!r} quantity{where}"),
                alcohol=alcohol,
            )
        )

    paid_by = data.get("paid_by")
    return Receipt(
        store=str(data.get("store", "")),
        date=str(data.get("date", "")),
        paid_by=str(paid_by) if paid_by else None,
        items=items,
        tax=_number(data.get("tax", 0), f"Receipt tax{where}"),
        source=source,
    )


def _json_files(paths: Iterable[str | Path]) -> list[Path]:
    files: list[Path] = []
    for path in map(Path, paths):
        if path.is_dir():
            files.extend(sorted(path.glob("*.json")))
        elif path.is_file():
            files.append(path)
        else:
            raise FileNotFoundError(f"No such file or directory: {path}")
    return files


def load_receipts(paths: Iterable[str | Path]) -> list[Receipt]:
    """Load receipts from JSON files and/or directories of ``*.json`` files."""
    receipts = []
    for path in _json_files(paths):
        with open(path, encoding="utf-8") as handle:
            try:
                data = json.load(handle)
            except json.JSONDecodeError as exc:
                raise ValueError(f"Invalid JSON in {path}: {exc}") from exc
        for entry in data if isinstance(data, list) else [data]:
            receipts.append(parse_receipt(entry, source=path.name))
    return receipts


def compute_shares(
    receipts: Sequence[Receipt], people: Sequence[str], drinkers: Sequence[str]
) -> dict[str, dict[str, float]]:
    """Compute what each person owes, paid and their balance.

    Mirrors the formulas written to the Summary sheet.
    """
    non_alcohol = sum(r.subtotal(False) + r.tax_split()[0] for r in receipts)
    alcohol = sum(r.subtotal(True) + r.tax_split()[1] for r in receipts)
    result = {}
    for person in people:
        non_alcohol_share = non_alcohol / len(people)
        alcohol_share = alcohol / len(drinkers) if person in drinkers else 0.0
        paid = sum(r.subtotal() + r.tax for r in receipts if r.paid_by == person)
        owed = non_alcohol_share + alcohol_share
        result[person] = {
            "non_alcohol": round(non_alcohol_share, 2),
            "alcohol": round(alcohol_share, 2),
            "owed": round(owed, 2),
            "paid": round(paid, 2),
            "balance": round(paid - owed, 2),
        }
    return result


def _validate_people(
    receipts: Sequence[Receipt], people: Sequence[str], drinkers: Sequence[str]
) -> None:
    if not people:
        raise ValueError("At least one person is required")
    if len(set(people)) != len(people):
        raise ValueError("People must be unique")
    unknown = [d for d in drinkers if d not in people]
    if unknown:
        raise ValueError(f"Drinkers not in people: {', '.join(unknown)}")
    payers = sorted({r.paid_by for r in receipts if r.paid_by and r.paid_by not in people})
    if payers:
        raise ValueError(f"Receipts paid by someone not in people: {', '.join(payers)}")
    if not drinkers and any(r.subtotal(True) for r in receipts):
        raise ValueError("Receipts contain alcohol but no drinkers were given")


def _bold_row(ws, row: int) -> None:
    for cell in ws[row]:
        cell.font = Font(bold=True)


def _autosize(ws) -> None:
    for index, column in enumerate(ws.columns, start=1):
        width = max(len(str(c.value)) if c.value is not None and not str(c.value).startswith("=") else 10
                    for c in column)
        ws.column_dimensions[get_column_letter(index)].width = min(max(width + 2, 10), 40)


def _yes_no_validation(ws, cell_range: str) -> None:
    validation = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    ws.add_data_validation(validation)
    validation.add(cell_range)


def build_workbook(
    receipts: Sequence[Receipt], people: Sequence[str], drinkers: Sequence[str] | None = None
) -> Workbook:
    """Build a workbook with Items, Receipts and Summary sheets.

    Totals and shares are Excel formulas, so editing an item's price or
    toggling its ``Alcohol`` / a person's ``Drinks Alcohol`` cell recalculates
    the split.
    """
    people = list(people)
    drinkers = list(people if drinkers is None else drinkers)
    _validate_people(receipts, people, drinkers)

    wb = Workbook()

    items_ws = wb.active
    items_ws.title = "Items"
    items_ws.append(ITEM_HEADERS)
    _bold_row(items_ws, 1)
    for number, receipt in enumerate(receipts, start=1):
        for item in receipt.lines():
            row = items_ws.max_row + 1
            items_ws.append(
                [
                    number,
                    receipt.date,
                    receipt.store,
                    receipt.paid_by or "",
                    item.name,
                    item.quantity,
                    item.price,
                    f"=ROUND(F{row}*G{row},2)",
                    "Yes" if item.alcohol else "No",
                    receipt.source,
                ]
            )
            items_ws[f"G{row}"].number_format = MONEY_FORMAT
            items_ws[f"H{row}"].number_format = MONEY_FORMAT
    if items_ws.max_row > 1:
        _yes_no_validation(items_ws, f"I2:I{items_ws.max_row}")
    items_ws.freeze_panes = "A2"
    _autosize(items_ws)

    receipts_ws = wb.create_sheet("Receipts")
    receipts_ws.append(
        ["Receipt #", "Date", "Store", "Paid By", "Non-Alcohol", "Alcohol", "Total", "Source File"]
    )
    _bold_row(receipts_ws, 1)
    for number, receipt in enumerate(receipts, start=1):
        row = number + 1
        receipts_ws.append(
            [
                number,
                receipt.date,
                receipt.store,
                receipt.paid_by or "",
                f'=SUMIFS(Items!$H:$H,Items!$A:$A,A{row},Items!$I:$I,"No")',
                f'=SUMIFS(Items!$H:$H,Items!$A:$A,A{row},Items!$I:$I,"Yes")',
                f"=E{row}+F{row}",
                receipt.source,
            ]
        )
        for col in "EFG":
            receipts_ws[f"{col}{row}"].number_format = MONEY_FORMAT
    receipts_ws.freeze_panes = "A2"
    _autosize(receipts_ws)

    summary_ws = wb.create_sheet("Summary")
    summary_ws["A1"] = "Shared Cost Summary"
    summary_ws["A1"].font = Font(bold=True, size=14)
    summary_ws["A2"] = "Non-alcohol total"
    summary_ws["B2"] = '=SUMIFS(Items!$H:$H,Items!$I:$I,"No")'
    summary_ws["A3"] = "Alcohol total"
    summary_ws["B3"] = '=SUMIFS(Items!$H:$H,Items!$I:$I,"Yes")'
    summary_ws["A4"] = "Grand total"
    summary_ws["B4"] = "=B2+B3"
    for row in (2, 3, 4):
        summary_ws[f"A{row}"].font = Font(bold=True)
        summary_ws[f"B{row}"].number_format = MONEY_FORMAT

    header_row = 6
    first, last = header_row + 1, header_row + len(people)
    summary_ws.append([])
    summary_ws.append(
        [
            "Person",
            "Drinks Alcohol",
            "Non-Alcohol Share",
            "Alcohol Share",
            "Total Owed",
            "Paid",
            "Balance (+ means is owed)",
        ]
    )
    _bold_row(summary_ws, header_row)
    drinker_range = f"$B${first}:$B${last}"
    for row, person in enumerate(people, start=first):
        summary_ws.append(
            [
                person,
                "Yes" if person in drinkers else "No",
                f"=$B$2/COUNTA($A${first}:$A${last})",
                f'=IF(B{row}="Yes",$B$3/COUNTIF({drinker_range},"Yes"),0)',
                f"=C{row}+D{row}",
                f"=SUMIFS(Items!$H:$H,Items!$D:$D,A{row})",
                f"=F{row}-E{row}",
            ]
        )
    total_row = last + 1
    summary_ws.append(
        ["Total", ""] + [f"=SUM({c}{first}:{c}{last})" for c in "CDEFG"]
    )
    _bold_row(summary_ws, total_row)
    for row in range(first, total_row + 1):
        for col in "CDEFG":
            summary_ws[f"{col}{row}"].number_format = MONEY_FORMAT
    _yes_no_validation(summary_ws, f"B{first}:B{last}")
    _autosize(summary_ws)

    wb.active = wb.index(summary_ws)
    return wb
