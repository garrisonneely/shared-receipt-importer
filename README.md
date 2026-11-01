# shared-receipt-importer

Loads receipts from JSON files into an Excel workbook that splits shared costs
between a group of people. **Non-alcohol** costs are split evenly between
everyone. **Alcohol** costs are split only between the people who drink.

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
python -m receipt_importer examples/receipts -o shared_costs.xlsx \
    --people Alice Bob Carol --drinkers Alice Bob
```

- `inputs`: one or more receipt JSON files or directories. For a directory,
  every `*.json` file in it is loaded.
- `--people`: everyone sharing the costs. Each receipt's `paid_by` must be one
  of them.
- `--drinkers`: the people who share alcohol costs. If you leave it out,
  everyone in `--people` shares them.
- `-o/--output`: the workbook to write. Defaults to `shared_costs.xlsx`.

The command also prints a summary of what each person owes and has paid.

## Receipt JSON format

A file contains either one receipt object or a list of them:

```json
{
  "store": "Corner Market",
  "date": "2026-10-30",
  "paid_by": "Alice",
  "items": [
    {"name": "Tortilla Chips", "price": 3.50, "quantity": 2},
    {"name": "IPA 6-pack", "price": 11.99},
    {"name": "House Red", "price": 14.00, "alcohol": true}
  ],
  "tax": 2.50
}
```

| Field | Required | Notes |
| --- | --- | --- |
| `items[].name` | yes | |
| `items[].price` | yes | Unit price |
| `items[].quantity` | no | Defaults to `1` |
| `items[].alcohol` | no | `true`/`false`. If missing, it is guessed from the item name: keywords such as beer, wine, IPA and vodka count as alcohol, while root beer, ginger beer and non-alcoholic items don't. |
| `tax` | no | Split proportionally into an alcohol line and a non-alcohol line |
| `store`, `date`, `paid_by` | no | Used for display. `paid_by` is also used for the "Paid" column. |

## Workbook layout

- **Items**: one row per item (plus the tax lines). The **Alcohol** column
  holds Yes/No.
- **Receipts**: alcohol, non-alcohol and total amounts for each receipt.
- **Summary**:
  - Non-alcohol, alcohol and grand totals.
  - A row per person: whether they drink, their non-alcohol share, their
    alcohol share, total owed, amount paid, and balance. A positive balance
    means the others owe that person money.

All totals and shares are Excel formulas, so they recalculate in the
spreadsheet when you:

- change a price or quantity;
- switch an item's **Alcohol** value;
- switch a person's **Drinks Alcohol** value.

## Tests

```bash
pip install pytest
python -m pytest
```
