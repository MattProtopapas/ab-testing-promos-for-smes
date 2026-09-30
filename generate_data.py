"""Generate a fake email-campaign A/B test dataset with Faker.

Simulates a promotional email sent in daily batches over a campaign window,
with recipients randomly split into a control group (standard email) and a
treatment group (email with a discount code). Each recipient row records the
funnel - opened, clicked, converted - and the order value if they bought.

Baked-in behaviour worth knowing about, because the analysis is designed to
surface it:
  - the treatment converts better but has a LOWER average basket (people use
    the discount), so "conversion went up" is not the whole story
  - day-of-week seasonality (and a Black Friday spike) hits BOTH groups equally
  - an optional novelty effect makes the treatment look best in the first days
  - returning customers convert more than new ones in BOTH groups

A small percentage of rows are deliberately corrupted (invalid group labels,
impossible funnels, negative order values, duplicate customers) so
validate_data.py has real work to do - pass --error-rate 0 for clean data.
"""

import argparse
import csv
import math
import os
import random
from datetime import date, timedelta

from faker import Faker

from ab_common import CSV_FIELDS, GROUPS, RAW_DATA_PATH, SEGMENTS

# Campaign window: 14 daily batches leading up to and including Black Friday weekend.
DEFAULT_START_DATE = date(2025, 11, 17)
BLACK_FRIDAY = date(2025, 11, 28)

# Funnel assumptions (open rate, click rate given open) per group. Conversion
# given click is back-solved so that the overall conversion rate hits the
# --control-rate / --treatment-rate targets.
OPEN_RATE = {"control": 0.34, "treatment": 0.41}
CLICK_GIVEN_OPEN = {"control": 0.22, "treatment": 0.24}

# Seasonality multiplier by weekday (Mon=0 .. Sun=6): quieter weekends, plus a
# Black Friday bump applied on top. Applied identically to both groups.
WEEKDAY_MULTIPLIER = {0: 1.05, 1: 1.05, 2: 1.0, 3: 1.0, 4: 1.1, 5: 0.8, 6: 0.75}
BLACK_FRIDAY_MULTIPLIER = 1.6

SEGMENT_SHARE_RETURNING = 0.4
SEGMENT_CONVERSION_MULTIPLIER = {"new": 0.8, "returning": 1.3}

BASKET_SIGMA = 0.45  # log-normal spread of order values
MIN_ORDER_VALUE = 5.0


def novelty_multipliers(days: int, novelty: float) -> list[float]:
    """Treatment-only multiplier per campaign day that starts high and decays.

    Normalised to average 1.0 over the window so the overall treatment rate
    still matches --treatment-rate; only the *shape* over time changes.
    """
    raw = [1 + novelty * math.exp(-d / 3.0) for d in range(days)]
    avg = sum(raw) / len(raw)
    return [r / avg for r in raw]


def lognormal_order_value(target_mean: float) -> float:
    mu = math.log(target_mean) - BASKET_SIGMA**2 / 2
    return max(MIN_ORDER_VALUE, round(random.lognormvariate(mu, BASKET_SIGMA), 2))


def generate_row(
    fake: Faker,
    index: int,
    send_date: date,
    group: str,
    conv_rate: dict[str, float],
    basket_mean: dict[str, float],
    treatment_day_multiplier: float,
) -> dict:
    segment = "returning" if random.random() < SEGMENT_SHARE_RETURNING else "new"

    season = WEEKDAY_MULTIPLIER[send_date.weekday()]
    if send_date == BLACK_FRIDAY:
        season *= BLACK_FRIDAY_MULTIPLIER
    group_multiplier = treatment_day_multiplier if group == "treatment" else 1.0
    target_conversion = conv_rate[group] * season * group_multiplier * SEGMENT_CONVERSION_MULTIPLIER[segment]

    p_open = OPEN_RATE[group]
    p_click = CLICK_GIVEN_OPEN[group]
    p_convert_given_click = min(0.95, target_conversion / (p_open * p_click))

    opened = random.random() < p_open
    clicked = opened and random.random() < p_click
    converted = clicked and random.random() < p_convert_given_click
    order_value = lognormal_order_value(basket_mean[group]) if converted else 0.0

    return {
        "customer_id": f"CUST-{index:06d}",
        "email": fake.unique.email(),
        "group": group,
        "segment": segment,
        "send_date": send_date.isoformat(),
        "opened": opened,
        "clicked": clicked,
        "converted": converted,
        "order_value": order_value,
    }


def corrupt_row(row: dict) -> dict:
    """Introduce one realistic data-quality issue into a row, in place."""
    issue = random.choice(
        [
            "invalid_group",
            "blank_group",
            "negative_order_value",
            "order_without_conversion",
            "conversion_without_order",
            "click_without_open",
            "blank_customer_id",
            "bad_date",
            "unknown_segment",
        ]
    )
    if issue == "invalid_group":
        row["group"] = "variant_c"
    elif issue == "blank_group":
        row["group"] = ""
    elif issue == "negative_order_value":
        row["order_value"] = -abs(row["order_value"] or 25.0)
    elif issue == "order_without_conversion":
        row["converted"] = False
        row["order_value"] = 49.99
    elif issue == "conversion_without_order":
        row["opened"] = row["clicked"] = row["converted"] = True
        row["order_value"] = 0.0
    elif issue == "click_without_open":
        row["opened"] = False
        row["clicked"] = True
    elif issue == "blank_customer_id":
        row["customer_id"] = ""
    elif issue == "bad_date":
        row["send_date"] = "not-a-date"
    elif issue == "unknown_segment":
        row["segment"] = "vip"
    return row


def generate_dataset(
    count: int,
    error_rate: float,
    seed: int | None,
    control_rate: float,
    treatment_rate: float,
    control_basket: float,
    treatment_basket: float,
    days: int,
    treatment_share: float,
    novelty: float,
    start_date: date,
) -> list[dict]:
    if seed is not None:
        random.seed(seed)
        Faker.seed(seed)
    fake = Faker()

    conv_rate = {"control": control_rate, "treatment": treatment_rate}
    basket_mean = {"control": control_basket, "treatment": treatment_basket}
    day_multipliers = novelty_multipliers(days, novelty)

    rows = []
    for i in range(1, count + 1):
        day_index = (i - 1) * days // count  # spread sends evenly across the window
        send_date = start_date + timedelta(days=day_index)
        group = "treatment" if random.random() < treatment_share else "control"
        rows.append(generate_row(fake, i, send_date, group, conv_rate, basket_mean, day_multipliers[day_index]))

    n_errors = round(count * error_rate)
    for row in random.sample(rows, k=min(n_errors, count)):
        corrupt_row(row)

    # Inject a duplicate customer (same person exported twice) to exercise the
    # validator's duplicate check, as long as there's more than one row.
    if error_rate > 0 and count > 1:
        dup = dict(rows[0])
        dup["customer_id"] = rows[-1]["customer_id"]
        rows.insert(len(rows) // 2, dup)

    return rows


def write_csv(rows: list[dict], path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a fake email-campaign A/B test dataset.")
    parser.add_argument("--count", type=int, default=10000, help="Number of recipients to generate (default: 10000)")
    parser.add_argument(
        "--error-rate",
        type=float,
        default=0.02,
        help="Fraction of rows (0-1) to deliberately corrupt, to exercise validate_data.py (default: 0.02)",
    )
    parser.add_argument("--seed", type=int, default=None, help="Random seed for reproducible output")
    parser.add_argument("--output", type=str, default=RAW_DATA_PATH, help="Path to write the raw CSV to")
    parser.add_argument("--control-rate", type=float, default=0.030, help="True conversion rate of the control email (default: 0.03)")
    parser.add_argument("--treatment-rate", type=float, default=0.039, help="True conversion rate of the treatment email (default: 0.039)")
    parser.add_argument("--control-basket", type=float, default=62.0, help="Mean order value, control group (default: 62)")
    parser.add_argument("--treatment-basket", type=float, default=55.0, help="Mean order value, treatment group (default: 55)")
    parser.add_argument("--days", type=int, default=14, help="Number of daily send batches (default: 14)")
    parser.add_argument("--treatment-share", type=float, default=0.5, help="Probability a recipient lands in treatment (default: 0.5)")
    parser.add_argument("--novelty", type=float, default=0.6, help="Size of the decaying first-days treatment boost; 0 disables (default: 0.6)")
    parser.add_argument("--start-date", type=date.fromisoformat, default=DEFAULT_START_DATE, help="First send date, YYYY-MM-DD (default: 2025-11-17)")
    args = parser.parse_args()

    if args.days < 1:
        parser.error("--days must be at least 1")
    if not 0 < args.treatment_share < 1:
        parser.error("--treatment-share must be strictly between 0 and 1")

    rows = generate_dataset(
        args.count,
        args.error_rate,
        args.seed,
        args.control_rate,
        args.treatment_rate,
        args.control_basket,
        args.treatment_basket,
        args.days,
        args.treatment_share,
        args.novelty,
        args.start_date,
    )
    write_csv(rows, args.output)
    print(f"Generated {len(rows)} rows ({round(args.error_rate * 100)}% target error rate) -> {args.output}")


if __name__ == "__main__":
    main()
