"""Validate the raw campaign CSV against the CampaignRecipient schema.

Reads the output of generate_data.py, checks every row against the pydantic
model in ab_common.py (types, allowed group/segment labels, funnel
consistency) plus a duplicate-customer check, and writes:
  - a cleaned CSV containing only the rows that passed validation
  - a plain-text report listing every rejected row and why it was rejected

analyze_ab_test.py is expected to run only on the cleaned CSV. A customer who
appears twice would otherwise be counted twice - and if they sit in both
groups, they contaminate the comparison.
"""

import argparse
import csv
import os

from pydantic import ValidationError

from ab_common import (
    CSV_FIELDS,
    RAW_DATA_PATH,
    VALIDATED_DATA_PATH,
    VALIDATION_REPORT_PATH,
    CampaignRecipient,
)


def read_rows(path: str) -> list[dict]:
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def validate_rows(rows: list[dict]) -> tuple[list[dict], list[str]]:
    valid_rows: list[dict] = []
    errors: list[str] = []
    seen_customers: set[str] = set()

    for line_no, row in enumerate(rows, start=2):  # header is line 1
        customer = row.get("customer_id") or f"<row {line_no}>"

        if row.get("customer_id") in seen_customers:
            errors.append(f"Line {line_no} (customer_id={customer}): duplicate customer, dropped")
            continue

        try:
            item = CampaignRecipient(**row)
        except ValidationError as exc:
            reasons = "; ".join(
                f"{e['loc'][0] if e['loc'] else 'row'}: {e['msg']}" for e in exc.errors()
            )
            errors.append(f"Line {line_no} (customer_id={customer}): {reasons}")
            continue

        seen_customers.add(item.customer_id)
        valid_rows.append(row)

    return valid_rows, errors


def write_csv(rows: list[dict], path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def write_report(errors: list[str], total: int, valid: int, path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("Campaign data validation report\n")
        f.write("================================\n")
        f.write(f"Total rows read:   {total}\n")
        f.write(f"Rows passed:       {valid}\n")
        f.write(f"Rows rejected:     {len(errors)}\n\n")
        if errors:
            f.write("Rejected rows:\n")
            for err in errors:
                f.write(f"  - {err}\n")
        else:
            f.write("No validation errors found.\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate the raw campaign CSV.")
    parser.add_argument("--input", type=str, default=RAW_DATA_PATH, help="Path to the raw CSV to validate")
    parser.add_argument("--output", type=str, default=VALIDATED_DATA_PATH, help="Path to write the cleaned CSV to")
    parser.add_argument(
        "--report", type=str, default=VALIDATION_REPORT_PATH, help="Path to write the validation report to"
    )
    args = parser.parse_args()

    rows = read_rows(args.input)
    valid_rows, errors = validate_rows(rows)

    write_csv(valid_rows, args.output)
    write_report(errors, len(rows), len(valid_rows), args.report)

    print(f"Validated {len(rows)} rows: {len(valid_rows)} passed, {len(errors)} rejected")
    print(f"Cleaned data -> {args.output}")
    print(f"Report       -> {args.report}")

    if not valid_rows:
        raise SystemExit("No valid rows survived validation; nothing to analyze.")


if __name__ == "__main__":
    main()
