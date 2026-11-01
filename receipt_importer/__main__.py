"""Command line entry point: ``python -m receipt_importer``."""

from __future__ import annotations

import argparse
import sys

from .importer import build_workbook, compute_shares, load_receipts


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="receipt_importer",
        description=(
            "Load receipts from JSON files into an Excel workbook that splits "
            "shared costs, with alcohol shared only among drinkers."
        ),
    )
    parser.add_argument("inputs", nargs="+", help="Receipt JSON files or directories of *.json files")
    parser.add_argument("-o", "--output", default="shared_costs.xlsx", help="Output .xlsx file")
    parser.add_argument("--people", nargs="+", required=True, help="Everyone sharing the costs")
    parser.add_argument(
        "--drinkers",
        nargs="*",
        help="People who share alcohol costs (default: everyone in --people)",
    )
    args = parser.parse_args(argv)

    try:
        receipts = load_receipts(args.inputs)
        workbook = build_workbook(receipts, args.people, args.drinkers)
    except (OSError, ValueError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    workbook.save(args.output)

    drinkers = args.people if args.drinkers is None else args.drinkers
    print(f"Imported {len(receipts)} receipt(s) into {args.output}")
    for person, share in compute_shares(receipts, args.people, drinkers).items():
        print(
            f"  {person}: owes {share['owed']:.2f} "
            f"(non-alcohol {share['non_alcohol']:.2f}, alcohol {share['alcohol']:.2f}), "
            f"paid {share['paid']:.2f}, balance {share['balance']:+.2f}"
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())
