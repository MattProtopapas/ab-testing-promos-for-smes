"""Analyze the A/B test and write everything to one Excel file.

Reads the validated CSV produced by validate_data.py, compares the control and
treatment groups, and writes a brand-new workbook (no template) with six sheets:
  - Input Data               the validated records the analysis was run on
  - Results                  group comparison, every test with p-value + CI,
                             power of the test you ran, plain-English verdict
  - Daily Breakdown          per-day funnel per group and the cumulative
                             p-value you would have seen if you had "peeked"
  - Segments                 the same z-test per customer segment, with the
                             multiple-comparison correction shown
  - Sample Size Calculator   a live Excel calculator (edit the yellow cells)
                             plus a pre-computed table: how many recipients and
                             how many days the NEXT test needs
  - Pitfalls                 the checklist of ways this analysis goes wrong

Methods (see ab_stats.py for the formulas):
  - Conversion rate:            two-proportion z-test + chi-square (2x2) + 95% CI
  - Average order value:        Welch's t-test on converters only
  - Revenue per recipient:      Welch's t-test on all recipients (zeros included)
  - Sample ratio mismatch:      chi-square goodness-of-fit on the group counts
  - Sample size / duration:     standard two-sided power formula for proportions
"""

import argparse
import csv
import math
import os
from collections import defaultdict
from datetime import date
from statistics import mean, stdev

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import DataPoint
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ab_common import CSV_FIELDS, GROUPS, OUTPUT_WORKBOOK_PATH, SEGMENTS, VALIDATED_DATA_PATH, CampaignRecipient
from ab_stats import (
    achieved_power_proportions,
    chi_square_2x2,
    days_to_run,
    minimum_detectable_lift,
    required_sample_size_means,
    required_sample_size_proportions,
    sample_ratio_mismatch,
    two_proportion_ztest,
    welch_ttest,
)

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SECTION_FONT = Font(bold=True, size=12, color="1F4E78")
SIGNIFICANT_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
NOT_SIGNIFICANT_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
WARNING_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
INPUT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
INPUT_BORDER = Border(*(Side(style="thin", color="BF9000"),) * 4)

# Chart colours: validated categorical slots (control = blue, treatment = orange).
COLOR_CONTROL = "2A78D6"
COLOR_TREATMENT = "EB6834"
COLOR_THRESHOLD = "8A8A8A"

PCT = "0.00%"
MONEY = "#,##0.00"
PVAL = "0.0000"
NUM3 = "0.000"

DAILY_FIELDS = [
    "send_date",
    "weekday",
    "control_sent",
    "control_converted",
    "control_rate",
    "treatment_sent",
    "treatment_converted",
    "treatment_rate",
    "daily_relative_lift",
    "cum_control_sent",
    "cum_control_converted",
    "cum_treatment_sent",
    "cum_treatment_converted",
    "cum_relative_lift",
    "cum_p_value",
    "looks_significant_today",
]

SEGMENT_FIELDS = [
    "segment",
    "control_sent",
    "control_converted",
    "control_rate",
    "treatment_sent",
    "treatment_converted",
    "treatment_rate",
    "relative_lift",
    "p_value",
    "ci_low",
    "ci_high",
    "significant_at_alpha",
    "bonferroni_alpha",
    "significant_after_correction",
]

PITFALLS = [
    (
        "Peeking / stopping early",
        "Checking the p-value every day and stopping the moment it dips under 0.05 inflates the false-positive "
        "rate far above 5%. Decide the sample size BEFORE the test (Sample Size Calculator sheet), run until you "
        "reach it, then look once. The Daily Breakdown sheet shows how the cumulative p-value wanders on the way.",
    ),
    (
        "Seasonality",
        "Weekends, paydays and Black Friday change everyone's behaviour. Both groups must run over the SAME "
        "calendar window so seasonality cancels out. Never compare 'this week with the promo' to 'last week "
        "without' - that is a before/after comparison, not an A/B test.",
    ),
    (
        "Novelty effect",
        "A new layout or a surprise discount often performs best in its first days simply because it is new. "
        "Check whether the treatment's lift shrinks over the campaign (Daily Breakdown). Run at least one full "
        "weekly cycle, ideally two, so the result reflects steady-state behaviour.",
    ),
    (
        "Sample ratio mismatch",
        "If you planned a 50/50 split but the data shows 47/53 with thousands of recipients, the randomisation "
        "or the export is broken (bounces, bot filters, duplicate rows). The health check on the Results sheet "
        "flags this; if it fires, fix the data before trusting anything else.",
    ),
    (
        "Slicing until something is significant",
        "Testing 'new customers', 'returning customers', 'mobile', 'weekend' etc. multiplies your chances of a "
        "fluke. Every extra slice at alpha = 0.05 is another 5% chance of a false alarm. Pre-register ONE primary "
        "metric and ONE population; treat segment results as hypotheses for the next test (see the Bonferroni "
        "column on the Segments sheet).",
    ),
    (
        "Wrong metric",
        "A discount raises conversion but lowers basket size. Judge the promotion on the metric that pays the "
        "bills - revenue (or margin) per recipient - not on conversion alone. All three are reported side by side.",
    ),
    (
        "Confusing statistical with practical significance",
        "With a huge list, a 0.1% lift can be 'significant' and worthless; with a small list, a 30% lift can be "
        "'not significant' and still real. Read the confidence interval: it tells you the range of lifts the data "
        "are consistent with. Then ask whether the low end would still be worth acting on.",
    ),
    (
        "Underpowered tests",
        "If the test could only detect a 40% lift, a 'not significant' result says almost nothing about a real "
        "15% lift. The Results sheet reports the minimum detectable lift for the sample you actually had.",
    ),
    (
        "Changing the test mid-flight",
        "Editing the subject line, the discount, or the audience halfway through creates two different tests "
        "glued together. If you must change something, restart the clock.",
    ),
    (
        "Contamination between groups",
        "Discount codes get shared. Customers appearing in both groups, or control customers redeeming the "
        "treatment code, blur the difference you are trying to measure. Deduplicate (validate_data.py does) and "
        "make codes single-use or personalised where possible.",
    ),
]


# --------------------------------------------------------------------------- #
# Loading
# --------------------------------------------------------------------------- #
def read_items(path: str) -> list[CampaignRecipient]:
    with open(path, newline="", encoding="utf-8") as f:
        return [CampaignRecipient(**row) for row in csv.DictReader(f)]


# --------------------------------------------------------------------------- #
# Analysis
# --------------------------------------------------------------------------- #
def group_summary(items: list[CampaignRecipient]) -> dict[str, dict]:
    summary = {}
    for g in GROUPS:
        rows = [r for r in items if r.group == g]
        n = len(rows)
        opened = sum(r.opened for r in rows)
        clicked = sum(r.clicked for r in rows)
        converted = sum(r.converted for r in rows)
        revenue = sum(r.order_value for r in rows)
        order_values = [r.order_value for r in rows if r.converted]
        summary[g] = {
            "recipients": n,
            "opened": opened,
            "open_rate": opened / n if n else float("nan"),
            "clicked": clicked,
            "click_rate": clicked / n if n else float("nan"),
            "converted": converted,
            "conversion_rate": converted / n if n else float("nan"),
            "revenue": revenue,
            "revenue_per_recipient": revenue / n if n else float("nan"),
            "avg_order_value": mean(order_values) if order_values else float("nan"),
            "order_values": order_values,
            "revenue_per_recipient_values": [r.order_value for r in rows],
        }
    return summary


def run_tests(summary: dict[str, dict], alpha: float, power: float, expected_treatment_share: float) -> dict:
    c, t = summary["control"], summary["treatment"]
    conv = two_proportion_ztest(c["converted"], c["recipients"], t["converted"], t["recipients"], alpha)
    chi = chi_square_2x2(c["converted"], c["recipients"], t["converted"], t["recipients"], alpha)
    aov = welch_ttest(c["order_values"], t["order_values"], alpha)
    rpr = welch_ttest(c["revenue_per_recipient_values"], t["revenue_per_recipient_values"], alpha)
    srm = sample_ratio_mismatch(c["recipients"], t["recipients"], expected_treatment_share)

    n_per_group = min(c["recipients"], t["recipients"])
    planning = {
        "achieved_power": achieved_power_proportions(conv["p1"], conv["p2"], n_per_group, alpha) if conv["p1"] != conv["p2"] else 0.0,
        "mde_abs": minimum_detectable_lift(conv["p1"], n_per_group, alpha, power),
        "n_needed_for_observed_lift": required_sample_size_proportions(conv["p1"], conv["p2"], alpha, power),
    }
    planning["mde_rel"] = planning["mde_abs"] / conv["p1"] if conv["p1"] else float("nan")
    return {"conversion": conv, "chi_square": chi, "aov": aov, "revenue_per_recipient": rpr, "srm": srm, "planning": planning}


def daily_breakdown(items: list[CampaignRecipient], alpha: float) -> list[dict]:
    by_day: dict[date, dict[str, list[int]]] = defaultdict(lambda: {g: [0, 0] for g in GROUPS})
    for r in items:
        by_day[r.send_date][r.group][0] += 1
        by_day[r.send_date][r.group][1] += int(r.converted)

    rows = []
    cum = {g: [0, 0] for g in GROUPS}
    for day in sorted(by_day):
        cs, cc = by_day[day]["control"]
        ts, tc = by_day[day]["treatment"]
        for g in GROUPS:
            cum[g][0] += by_day[day][g][0]
            cum[g][1] += by_day[day][g][1]
        c_rate = cc / cs if cs else float("nan")
        t_rate = tc / ts if ts else float("nan")
        cum_test = (
            two_proportion_ztest(cum["control"][1], cum["control"][0], cum["treatment"][1], cum["treatment"][0], alpha)
            if cum["control"][0] and cum["treatment"][0]
            else None
        )
        rows.append(
            {
                "send_date": day,
                "weekday": day.strftime("%a"),
                "control_sent": cs,
                "control_converted": cc,
                "control_rate": c_rate,
                "treatment_sent": ts,
                "treatment_converted": tc,
                "treatment_rate": t_rate,
                "daily_relative_lift": (t_rate - c_rate) / c_rate if cs and ts and c_rate else float("nan"),
                "cum_control_sent": cum["control"][0],
                "cum_control_converted": cum["control"][1],
                "cum_treatment_sent": cum["treatment"][0],
                "cum_treatment_converted": cum["treatment"][1],
                "cum_relative_lift": cum_test["relative_lift"] if cum_test else float("nan"),
                "cum_p_value": cum_test["p_value"] if cum_test else float("nan"),
                "looks_significant_today": bool(cum_test and cum_test["significant"]),
            }
        )
    return rows


def segment_breakdown(items: list[CampaignRecipient], alpha: float) -> list[dict]:
    bonferroni_alpha = alpha / len(SEGMENTS)
    rows = []
    for seg in SEGMENTS:
        sub = [r for r in items if r.segment == seg]
        c = [r for r in sub if r.group == "control"]
        t = [r for r in sub if r.group == "treatment"]
        if not c or not t:
            continue
        test = two_proportion_ztest(sum(r.converted for r in c), len(c), sum(r.converted for r in t), len(t), alpha)
        rows.append(
            {
                "segment": seg,
                "control_sent": len(c),
                "control_converted": sum(r.converted for r in c),
                "control_rate": test["p1"],
                "treatment_sent": len(t),
                "treatment_converted": sum(r.converted for r in t),
                "treatment_rate": test["p2"],
                "relative_lift": test["relative_lift"],
                "p_value": test["p_value"],
                "ci_low": test["ci_low"],
                "ci_high": test["ci_high"],
                "significant_at_alpha": test["significant"],
                "bonferroni_alpha": bonferroni_alpha,
                "significant_after_correction": test["p_value"] < bonferroni_alpha,
            }
        )
    return rows


def next_test_plan(baseline: float, lifts: list[float], alpha: float, power: float, daily_recipients: float, share: float) -> list[dict]:
    plan = []
    for lift in lifts:
        target = baseline * (1 + lift)
        n = required_sample_size_proportions(baseline, target, alpha, power)
        plan.append(
            {
                "relative_lift": lift,
                "target_rate": target,
                "recipients_per_group": n,
                "recipients_total": n * 2 if n != math.inf else math.inf,
                "days_to_run": days_to_run(n, daily_recipients, share),
            }
        )
    return plan


def verdict_text(tests: dict, alpha: float) -> list[str]:
    conv, aov, rpr, srm, plan = (
        tests["conversion"], tests["aov"], tests["revenue_per_recipient"], tests["srm"], tests["planning"]
    )
    lines = []
    if srm["mismatch"]:
        lines.append(
            "WARNING: the group split is far from the planned ratio (sample ratio mismatch). "
            "Fix the randomisation / export before trusting any result below."
        )
    if conv["significant"]:
        lines.append(
            f"Conversion: the treatment converted at {conv['p2']:.2%} vs {conv['p1']:.2%} for control, a relative lift of "
            f"{conv['relative_lift']:+.1%} (p = {conv['p_value']:.4f}). The 95% CI for the absolute difference is "
            f"{conv['ci_low']:+.2%} to {conv['ci_high']:+.2%}, which excludes zero - this is unlikely to be noise."
        )
    else:
        lines.append(
            f"Conversion: the treatment converted at {conv['p2']:.2%} vs {conv['p1']:.2%} for control "
            f"({conv['relative_lift']:+.1%} relative), but p = {conv['p_value']:.4f} and the 95% CI "
            f"({conv['ci_low']:+.2%} to {conv['ci_high']:+.2%}) includes zero. This test cannot rule out noise."
        )
    if not math.isnan(aov["p_value"]):
        direction = "lower" if aov["diff"] < 0 else "higher"
        lines.append(
            f"Basket: converters in the treatment spent {aov['mean2']:,.2f} on average vs {aov['mean1']:,.2f} "
            f"({direction} by {abs(aov['diff']):,.2f}, p = {aov['p_value']:.4f}). "
            + ("A discount that lifts conversion but shrinks baskets is common - do not judge on conversion alone."
               if aov["diff"] < 0 and aov["significant"] else "")
        )
    if not math.isnan(rpr["p_value"]):
        lines.append(
            f"Money metric: revenue per recipient was {rpr['mean2']:,.3f} vs {rpr['mean1']:,.3f} "
            f"({rpr['relative_lift']:+.1%}, p = {rpr['p_value']:.4f}, 95% CI {rpr['ci_low']:+.3f} to {rpr['ci_high']:+.3f}). "
            + ("The promotion paid off in revenue terms." if rpr["significant"] and rpr["diff"] > 0
               else "Revenue per recipient is noisier than conversion; this difference is not conclusive."
               if not rpr["significant"] else "The promotion LOST revenue per recipient.")
        )
    lines.append(
        f"Power: with this sample the test could reliably detect a relative conversion lift of about "
        f"{plan['mde_rel']:.0%} or more; to confirm the observed lift with 80% power you would need roughly "
        f"{plan['n_needed_for_observed_lift']:,} recipients per group."
        if plan["n_needed_for_observed_lift"] != math.inf else
        f"Power: with this sample the test could reliably detect a relative conversion lift of about {plan['mde_rel']:.0%} or more."
    )
    return lines


# --------------------------------------------------------------------------- #
# Workbook helpers
# --------------------------------------------------------------------------- #
def autosize_columns(ws, max_width: int = 60) -> None:
    for col_cells in ws.columns:
        length = max((len(str(c.value)) for c in col_cells if c.value is not None and not str(c.value).startswith("=")), default=10)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(length + 2, max_width)


def style_header_row(ws, row: int, ncols: int) -> None:
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def write_table(ws, fieldnames: list[str], rows: list[dict], formats: dict[str, str] | None = None) -> None:
    ws.append(fieldnames)
    style_header_row(ws, 1, len(fieldnames))
    for row in rows:
        ws.append([_clean(row.get(f, "")) for f in fieldnames])
    if formats:
        for f, fmt in formats.items():
            if f in fieldnames:
                col = fieldnames.index(f) + 1
                for r in range(2, len(rows) + 2):
                    ws.cell(row=r, column=col).number_format = fmt
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(fieldnames))}{len(rows) + 1}"
    autosize_columns(ws)


def _clean(value):
    """Excel cannot store NaN/inf; write them as blanks / text."""
    if isinstance(value, float):
        if math.isnan(value):
            return None
        if math.isinf(value):
            return "n/a"
    return value


class SheetWriter:
    """Tiny cursor-based helper for the label/value style sheets."""

    def __init__(self, ws):
        self.ws = ws
        self.row = 1

    def title(self, text: str) -> None:
        self.ws.cell(row=self.row, column=1, value=text).font = Font(bold=True, size=14, color="1F4E78")
        self.row += 2

    def section(self, text: str) -> None:
        self.ws.cell(row=self.row, column=1, value=text).font = SECTION_FONT
        self.row += 1

    def header(self, labels: list[str]) -> None:
        for i, label in enumerate(labels, start=1):
            self.ws.cell(row=self.row, column=i, value=label)
        style_header_row(self.ws, self.row, len(labels))
        self.row += 1

    def kv(self, label: str, value, fmt: str | None = None, fill: PatternFill | None = None, note: str | None = None) -> int:
        self.ws.cell(row=self.row, column=1, value=label)
        cell = self.ws.cell(row=self.row, column=2, value=_clean(value))
        if fmt:
            cell.number_format = fmt
        if fill:
            cell.fill = fill
        if note:
            self.ws.cell(row=self.row, column=3, value=note).font = Font(italic=True, color="595959")
        self.row += 1
        return self.row - 1

    def values(self, cells: list, fmt: str | None = None) -> None:
        for i, v in enumerate(cells, start=1):
            cell = self.ws.cell(row=self.row, column=i, value=_clean(v))
            if fmt and i > 1:
                cell.number_format = fmt
        self.row += 1

    def text(self, text: str, bold: bool = False) -> None:
        cell = self.ws.cell(row=self.row, column=1, value=text)
        cell.font = Font(bold=bold)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        self.row += 1

    def blank(self, n: int = 1) -> None:
        self.row += n



def finish_chart(chart) -> None:
    """Newer Excel hides axes unless told otherwise; straight lines, not splines."""
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    if chart.legend is not None:
        chart.legend.position = "b"
        chart.legend.overlay = False
    chart.y_axis.majorGridlines.spPr = None
    for s in chart.series:
        if hasattr(s, "smooth"):
            s.smooth = False

def sig_fill(significant: bool) -> PatternFill:
    return SIGNIFICANT_FILL if significant else NOT_SIGNIFICANT_FILL


# --------------------------------------------------------------------------- #
# Sheets
# --------------------------------------------------------------------------- #
def build_results_sheet(ws, summary: dict, tests: dict, alpha: float, power: float) -> None:
    w = SheetWriter(ws)
    c, t = summary["control"], summary["treatment"]
    conv, chi, aov, rpr, srm, plan = (
        tests["conversion"], tests["chi_square"], tests["aov"], tests["revenue_per_recipient"], tests["srm"], tests["planning"]
    )

    w.title("Did the promotion actually work?  A/B test results")

    w.section("Group comparison")
    w.header(["metric", "control", "treatment", "difference (treatment - control)"])
    comparison = [
        ("Recipients", "recipients", "#,##0"),
        ("Opened", "opened", "#,##0"),
        ("Open rate", "open_rate", PCT),
        ("Clicked", "clicked", "#,##0"),
        ("Click rate (of recipients)", "click_rate", PCT),
        ("Converted (placed an order)", "converted", "#,##0"),
        ("Conversion rate", "conversion_rate", PCT),
        ("Revenue", "revenue", MONEY),
        ("Revenue per recipient", "revenue_per_recipient", MONEY),
        ("Average order value (converters)", "avg_order_value", MONEY),
    ]
    chart_first_row = None
    for label, key, fmt in comparison:
        if key == "conversion_rate":
            chart_first_row = w.row
        w.values([label, c[key], t[key], t[key] - c[key]], fmt)
    w.blank()

    w.section("Health check: sample ratio mismatch (SRM)")
    w.kv("Observed share in treatment", srm["observed_share_2"], PCT)
    w.kv("Chi-square statistic", srm["statistic"], NUM3)
    w.kv("p-value", srm["p_value"], PVAL, WARNING_FILL if srm["mismatch"] else SIGNIFICANT_FILL,
         "Below 0.001 means the split is broken - stop and fix the data.")
    w.blank()

    w.section(f"Test 1 - Conversion rate: two-proportion z-test (alpha = {alpha})")
    w.kv("Control conversion rate", conv["p1"], PCT)
    w.kv("Treatment conversion rate", conv["p2"], PCT)
    w.kv("Absolute difference", conv["diff"], PCT)
    w.kv("Relative lift", conv["relative_lift"], PCT)
    w.kv("Pooled standard error", conv["se_pooled"], "0.00000")
    w.kv("z statistic", conv["z"], NUM3)
    w.kv("p-value (two-sided)", conv["p_value"], PVAL, sig_fill(conv["significant"]),
         "Probability of seeing a difference at least this big if the emails were truly identical.")
    w.kv(f"{1 - alpha:.0%} CI for the difference - low", conv["ci_low"], PCT)
    w.kv(f"{1 - alpha:.0%} CI for the difference - high", conv["ci_high"], PCT,
         note="If this range includes 0.00%, the data are consistent with 'no effect'.")
    w.kv("Statistically significant?", "YES" if conv["significant"] else "NO", fill=sig_fill(conv["significant"]))
    w.blank()

    w.section("Test 1b - Conversion rate: chi-square test of independence (same question, same answer)")
    w.kv("Chi-square statistic (df = 1)", chi["statistic"], NUM3, note="Equals z squared - the two tests are equivalent for a 2x2 table.")
    w.kv("p-value", chi["p_value"], PVAL, sig_fill(chi["significant"]))
    w.blank()

    w.section("Test 2 - Average order value among converters: Welch's t-test")
    w.kv("Control orders / mean / sd", f"{aov['n1']} / {aov['mean1']:,.2f} / {aov['sd1']:,.2f}" if aov["n1"] > 1 else f"{aov['n1']} orders")
    w.kv("Treatment orders / mean / sd", f"{aov['n2']} / {aov['mean2']:,.2f} / {aov['sd2']:,.2f}" if aov["n2"] > 1 else f"{aov['n2']} orders")
    w.kv("Difference in means", aov["diff"], MONEY)
    w.kv("t statistic", aov["t"], NUM3)
    w.kv("Degrees of freedom (Welch)", aov["df"], "0.0")
    w.kv("p-value (two-sided)", aov["p_value"], PVAL, sig_fill(aov["significant"]))
    w.kv(f"{1 - alpha:.0%} CI - low", aov["ci_low"], MONEY)
    w.kv(f"{1 - alpha:.0%} CI - high", aov["ci_high"], MONEY)
    w.kv("Statistically significant?", "YES" if aov["significant"] else "NO", fill=sig_fill(aov["significant"]))
    w.blank()

    w.section("Test 3 - Revenue per recipient (the money metric): Welch's t-test on ALL recipients")
    w.kv("Control mean revenue per recipient", rpr["mean1"], "#,##0.000")
    w.kv("Treatment mean revenue per recipient", rpr["mean2"], "#,##0.000")
    w.kv("Difference", rpr["diff"], "#,##0.000")
    w.kv("Relative lift", rpr["relative_lift"], PCT)
    w.kv("t statistic", rpr["t"], NUM3)
    w.kv("p-value (two-sided)", rpr["p_value"], PVAL, sig_fill(rpr["significant"]),
         "Zeros for non-buyers are included, so this is noisier than the conversion test.")
    w.kv(f"{1 - alpha:.0%} CI - low", rpr["ci_low"], "#,##0.000")
    w.kv(f"{1 - alpha:.0%} CI - high", rpr["ci_high"], "#,##0.000")
    w.kv("Statistically significant?", "YES" if rpr["significant"] else "NO", fill=sig_fill(rpr["significant"]))
    w.blank()

    w.section(f"What this test could actually see (power = {power:.0%})")
    w.kv("Recipients per group (smaller group)", min(c["recipients"], t["recipients"]), "#,##0")
    w.kv("Achieved power for the observed lift", plan["achieved_power"], PCT,
         note="Chance this test would flag the observed effect if it is real. Below 80% = underpowered.")
    w.kv("Minimum detectable absolute lift", plan["mde_abs"], PCT)
    w.kv("Minimum detectable relative lift", plan["mde_rel"], PCT,
         note="Smaller true effects than this would usually come out 'not significant'.")
    w.kv("Recipients per group needed to confirm the observed lift", plan["n_needed_for_observed_lift"], "#,##0")
    w.blank()

    w.section("Verdict")
    for line in verdict_text(tests, alpha):
        w.text(line)

    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 34

    if chart_first_row:
        chart = BarChart()
        chart.type = "col"
        chart.title = "Conversion rate by group"
        chart.y_axis.number_format = "0.0%"
        chart.legend = None
        data = Reference(ws, min_col=2, max_col=3, min_row=chart_first_row, max_row=chart_first_row)
        chart.add_data(data, from_rows=True, titles_from_data=False)
        chart.set_categories(Reference(ws, min_col=2, max_col=3, min_row=4, max_row=4))
        series = chart.series[0]
        series.graphicalProperties.solidFill = COLOR_CONTROL
        pt = DataPoint(idx=1)
        pt.graphicalProperties.solidFill = COLOR_TREATMENT
        series.dPt.append(pt)
        chart.height, chart.width = 7.5, 11
        finish_chart(chart)
        ws.add_chart(chart, "F3")


def build_daily_sheet(ws, daily: list[dict], alpha: float) -> None:
    formats = {
        "control_rate": PCT, "treatment_rate": PCT, "daily_relative_lift": PCT,
        "cum_relative_lift": PCT, "cum_p_value": PVAL, "send_date": "yyyy-mm-dd",
    }
    write_table(ws, DAILY_FIELDS, daily, formats)
    threshold_col = len(DAILY_FIELDS) + 1
    ws.cell(row=1, column=threshold_col, value="alpha_threshold")
    style_header_row(ws, 1, threshold_col)
    sig_col = DAILY_FIELDS.index("looks_significant_today") + 1
    p_col = DAILY_FIELDS.index("cum_p_value") + 1
    for r in range(2, len(daily) + 2):
        ws.cell(row=r, column=threshold_col, value=alpha).number_format = "0.00"
        looks_sig = ws.cell(row=r, column=sig_col).value
        ws.cell(row=r, column=p_col).fill = sig_fill(bool(looks_sig))
    ws.column_dimensions[get_column_letter(threshold_col)].width = 16
    last = len(daily) + 1

    # Chart 1: cumulative p-value over the campaign - the "peeking" picture.
    p_chart = LineChart()
    p_chart.title = "If you had peeked: cumulative p-value by day (dashed = alpha)"
    p_chart.y_axis.scaling.min = 0
    p_chart.y_axis.majorUnit = 0.1
    p_chart.y_axis.number_format = "0.00"
    p_chart.add_data(Reference(ws, min_col=p_col, min_row=1, max_row=last), titles_from_data=True)
    p_chart.add_data(Reference(ws, min_col=threshold_col, min_row=1, max_row=last), titles_from_data=True)
    p_chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
    p_chart.series[0].graphicalProperties.line.solidFill = COLOR_TREATMENT
    p_chart.series[0].graphicalProperties.line.width = 25000
    p_chart.series[0].marker.symbol = "circle"
    p_chart.series[0].marker.size = 6
    p_chart.series[1].graphicalProperties.line.solidFill = COLOR_THRESHOLD
    p_chart.series[1].graphicalProperties.line.dashStyle = "dash"
    p_chart.series[1].graphicalProperties.line.width = 15000
    p_chart.height, p_chart.width = 9, 18
    finish_chart(p_chart)
    ws.add_chart(p_chart, f"A{last + 3}")

    # Chart 2: daily conversion rate per group - seasonality hits both lines.
    rate_chart = LineChart()
    rate_chart.title = "Daily conversion rate by group (seasonality moves both lines)"
    rate_chart.y_axis.number_format = "0.0%"
    c_col = DAILY_FIELDS.index("control_rate") + 1
    t_col = DAILY_FIELDS.index("treatment_rate") + 1
    rate_chart.add_data(Reference(ws, min_col=c_col, min_row=1, max_row=last), titles_from_data=True)
    rate_chart.add_data(Reference(ws, min_col=t_col, min_row=1, max_row=last), titles_from_data=True)
    rate_chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
    for s, colour in zip(rate_chart.series, (COLOR_CONTROL, COLOR_TREATMENT)):
        s.graphicalProperties.line.solidFill = colour
        s.graphicalProperties.line.width = 25000
        s.marker.symbol = "circle"
        s.marker.size = 6
    rate_chart.height, rate_chart.width = 9, 18
    finish_chart(rate_chart)
    ws.add_chart(rate_chart, f"K{last + 3}")


def build_segments_sheet(ws, segments: list[dict]) -> None:
    formats = {"control_rate": PCT, "treatment_rate": PCT, "relative_lift": PCT, "p_value": PVAL,
               "ci_low": PCT, "ci_high": PCT, "bonferroni_alpha": PVAL}
    write_table(ws, SEGMENT_FIELDS, segments, formats)
    p_col = SEGMENT_FIELDS.index("p_value") + 1
    corr_col = SEGMENT_FIELDS.index("significant_after_correction") + 1
    for r in range(2, len(segments) + 2):
        ws.cell(row=r, column=p_col).fill = sig_fill(bool(ws.cell(row=r, column=corr_col).value))
    note_row = len(segments) + 3
    ws.cell(row=note_row, column=1, value=(
        "Segment results are exploratory. Each extra slice is another chance of a fluke, so the Bonferroni "
        "column divides alpha by the number of segments tested. Use a 'significant' segment as the hypothesis "
        "for your NEXT test, not as this test's conclusion."
    )).font = Font(italic=True)


def build_calculator_sheet(ws, defaults: dict, plan_rows: list[dict], aov_sd: float, aov_mean: float, alpha: float, power: float) -> None:
    w = SheetWriter(ws)
    w.title("Sample size calculator - how big and how long should the NEXT test be?")
    w.text("Edit the yellow cells. Everything else recalculates when you open the file in Excel.", bold=True)
    w.blank()

    w.section("Inputs")
    r_base = w.kv("Baseline conversion rate (control)", defaults["baseline"], PCT, INPUT_FILL, "Pre-filled with the control rate observed in this test.")
    r_lift = w.kv("Smallest relative lift worth detecting (MDE)", defaults["mde_rel"], PCT, INPUT_FILL, "e.g. 15% means 3.00% -> 3.45%. Smaller = needs far more people.")
    r_alpha = w.kv("Significance level (alpha)", alpha, "0.00", INPUT_FILL, "0.05 = accept a 5% chance of a false alarm.")
    r_power = w.kv("Power (1 - beta)", power, "0.00", INPUT_FILL, "0.80 = an 80% chance of catching a real effect of size MDE.")
    r_daily = w.kv("Recipients you can email per day", defaults["daily_recipients"], "#,##0", INPUT_FILL, "Pre-filled with the average daily send volume in this dataset.")
    r_share = w.kv("Share of recipients sent the treatment", defaults["treatment_share"], PCT, INPUT_FILL, "50% is most efficient; the SMALLER group sets the duration.")
    for r in (r_base, r_lift, r_alpha, r_power, r_daily, r_share):
        ws.cell(row=r, column=2).border = INPUT_BORDER
    w.blank()

    w.section("Outputs (live formulas)")
    B = lambda r: f"$B${r}"  # noqa: E731
    r_target = w.kv("Target conversion rate", f"={B(r_base)}*(1+{B(r_lift)})", PCT)
    r_za = w.kv("z for alpha (two-sided)", f"=NORMSINV(1-{B(r_alpha)}/2)", NUM3)
    r_zb = w.kv("z for power", f"=NORMSINV({B(r_power)})", NUM3)
    r_pbar = w.kv("Average of the two rates", f"=({B(r_base)}+{B(r_target)})/2", PCT)
    n_formula = (
        f"=ROUNDUP((({B(r_za)}*SQRT(2*{B(r_pbar)}*(1-{B(r_pbar)}))"
        f"+{B(r_zb)}*SQRT({B(r_base)}*(1-{B(r_base)})+{B(r_target)}*(1-{B(r_target)})))^2)"
        f"/(({B(r_target)}-{B(r_base)})^2),0)"
    )
    r_n = w.kv("Recipients needed PER GROUP", n_formula, "#,##0", SIGNIFICANT_FILL)
    w.kv("Recipients needed in TOTAL", f"={B(r_n)}*2", "#,##0")
    w.kv("Days to run the test", f"=ROUNDUP({B(r_n)}/({B(r_daily)}*MIN({B(r_share)},1-{B(r_share)})),0)", "#,##0", SIGNIFICANT_FILL,
         "Round UP to whole weeks so every weekday is covered at least once.")
    w.kv("Whole weeks to run", f"=ROUNDUP({B(w.row - 1)}/7,0)", "#,##0")
    w.blank()

    w.section("Reference table - pre-computed at the observed baseline and daily volume")
    w.header(["relative lift to detect", "target rate", "recipients per group", "recipients total", "days to run"])
    for row in plan_rows:
        w.values([row["relative_lift"], row["target_rate"], row["recipients_per_group"], row["recipients_total"], row["days_to_run"]])
        ws.cell(row=w.row - 1, column=1).number_format = PCT
        ws.cell(row=w.row - 1, column=2).number_format = PCT
        for col in (3, 4, 5):
            ws.cell(row=w.row - 1, column=col).number_format = "#,##0"
    w.blank()

    w.section("Bonus - testing a difference in AVERAGE ORDER VALUE instead")
    r_sd = w.kv("Standard deviation of order value", aov_sd, MONEY, INPUT_FILL, "Pre-filled from this test's orders.")
    r_mde_abs = w.kv("Smallest difference in mean order value worth detecting", round(aov_mean * 0.10, 2) if aov_mean else 5.0, MONEY, INPUT_FILL, "Pre-filled as 10% of the observed mean.")
    for r in (r_sd, r_mde_abs):
        ws.cell(row=r, column=2).border = INPUT_BORDER
    r_n_aov = w.kv("ORDERS needed per group", f"=ROUNDUP(2*(({B(r_za)}+{B(r_zb)})*{B(r_sd)}/{B(r_mde_abs)})^2,0)", "#,##0", SIGNIFICANT_FILL)
    w.kv("Recipients needed per group to get those orders", f"=ROUNDUP({B(r_n_aov)}/{B(r_base)},0)", "#,##0",
         note="Only converters have a basket, so divide by the conversion rate.")
    w.blank()
    w.text("Formula: n per group = [ z(1-a/2) * sqrt(2 * pbar * (1 - pbar)) + z(power) * sqrt(p1(1-p1) + p2(1-p2)) ]^2 / (p2 - p1)^2")

    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 16


def build_pitfalls_sheet(ws) -> None:
    w = SheetWriter(ws)
    w.title("Common pitfalls - read before believing any A/B result")
    w.header(["pitfall", "what goes wrong and what to do instead"])
    for name, text in PITFALLS:
        ws.cell(row=w.row, column=1, value=name).font = Font(bold=True)
        ws.cell(row=w.row, column=1).alignment = Alignment(vertical="top", wrap_text=True)
        cell = ws.cell(row=w.row, column=2, value=text)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[w.row].height = 62
        w.row += 1
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 120


def build_workbook(input_rows, summary, tests, daily, segments, plan_rows, calc_defaults, alpha, power) -> Workbook:
    wb = Workbook()

    ws_input = wb.active
    ws_input.title = "Input Data"
    write_table(ws_input, CSV_FIELDS, input_rows)

    build_results_sheet(wb.create_sheet("Results"), summary, tests, alpha, power)
    build_daily_sheet(wb.create_sheet("Daily Breakdown"), daily, alpha)
    build_segments_sheet(wb.create_sheet("Segments"), segments)

    all_orders = summary["control"]["order_values"] + summary["treatment"]["order_values"]
    aov_sd = stdev(all_orders) if len(all_orders) > 1 else 0.0
    aov_mean = mean(all_orders) if all_orders else 0.0
    build_calculator_sheet(wb.create_sheet("Sample Size Calculator"), calc_defaults, plan_rows, aov_sd, aov_mean, alpha, power)
    build_pitfalls_sheet(wb.create_sheet("Pitfalls"))

    wb.move_sheet("Results", offset=-1)  # open on the answer, not the raw data
    wb.active = 0
    return wb


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze the validated A/B test data and write results to Excel.")
    parser.add_argument("--input", type=str, default=VALIDATED_DATA_PATH, help="Path to the validated CSV")
    parser.add_argument("--output", type=str, default=OUTPUT_WORKBOOK_PATH, help="Path to write the Excel workbook to")
    parser.add_argument("--alpha", type=float, default=0.05, help="Significance level for every test (default: 0.05)")
    parser.add_argument("--power", type=float, default=0.80, help="Target power for sample-size planning (default: 0.8)")
    parser.add_argument("--mde-rel", type=float, default=0.15, help="Relative lift the NEXT test should detect, for the calculator defaults (default: 0.15)")
    parser.add_argument("--daily-recipients", type=float, default=None, help="Recipients you can email per day for the next test (default: observed average)")
    parser.add_argument("--treatment-share", type=float, default=0.5, help="Planned share of recipients in treatment, for the SRM check (default: 0.5)")
    args = parser.parse_args()

    items = read_items(args.input)
    with open(args.input, newline="", encoding="utf-8") as f:
        input_rows = list(csv.DictReader(f))
    if not any(r.group == "control" for r in items) or not any(r.group == "treatment" for r in items):
        raise SystemExit("Need at least one control and one treatment recipient to compare.")

    summary = group_summary(items)
    tests = run_tests(summary, args.alpha, args.power, args.treatment_share)
    daily = daily_breakdown(items, args.alpha)
    segments = segment_breakdown(items, args.alpha)

    n_days = len({r.send_date for r in items}) or 1
    daily_recipients = args.daily_recipients if args.daily_recipients else len(items) / n_days
    observed_share = summary["treatment"]["recipients"] / len(items)
    baseline = tests["conversion"]["p1"]
    plan_rows = next_test_plan(baseline, [0.05, 0.10, 0.15, 0.20, 0.25, 0.30, 0.50], args.alpha, args.power, daily_recipients, observed_share)
    calc_defaults = {
        "baseline": baseline,
        "mde_rel": args.mde_rel,
        "daily_recipients": round(daily_recipients),
        "treatment_share": round(observed_share, 2),
    }

    wb = build_workbook(input_rows, summary, tests, daily, segments, plan_rows, calc_defaults, args.alpha, args.power)
    os.makedirs(os.path.dirname(args.output) or ".", exist_ok=True)
    wb.save(args.output)

    conv, rpr = tests["conversion"], tests["revenue_per_recipient"]
    c, t = summary["control"], summary["treatment"]
    print(f"Analyzed {len(items)} recipients ({c['recipients']} control / {t['recipients']} treatment) -> {args.output}")
    print(f"Conversion: {conv['p1']:.2%} -> {conv['p2']:.2%} ({conv['relative_lift']:+.1%}), p = {conv['p_value']:.4f}, "
          f"95% CI [{conv['ci_low']:+.2%}, {conv['ci_high']:+.2%}] -> {'SIGNIFICANT' if conv['significant'] else 'not significant'}")
    print(f"Revenue/recipient: {rpr['mean1']:.3f} -> {rpr['mean2']:.3f} ({rpr['relative_lift']:+.1%}), p = {rpr['p_value']:.4f} "
          f"-> {'SIGNIFICANT' if rpr['significant'] else 'not significant'}")
    next_n = required_sample_size_proportions(baseline, baseline * (1 + args.mde_rel), args.alpha, args.power)
    print(f"Next test: to detect a {args.mde_rel:.0%} relative lift you need {next_n:,} per group "
          f"(~{days_to_run(next_n, daily_recipients, observed_share)} days at {daily_recipients:,.0f} recipients/day).")
    peeks = sum(1 for d in daily if d["looks_significant_today"])
    print(f"Peeking: the cumulative p-value was below {args.alpha} on {peeks} of {len(daily)} days.")


if __name__ == "__main__":
    main()
