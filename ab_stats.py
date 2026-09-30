"""Pure-Python statistics for two-group A/B tests.

No numpy / scipy: everything is built on the standard library
(``statistics.NormalDist`` for the normal distribution, ``math.lgamma`` plus a
continued-fraction incomplete-beta for Student's t). Kept separate from the
pipeline scripts so the formulas are easy to read, reuse and unit-test.

Methods:
  - Two-proportion z-test          conversion rate A vs B (pooled SE for the test,
                                    unpooled Wald SE for the confidence interval)
  - Chi-square test (2x2, df=1)     the same question asked a different way;
                                    the statistic equals z^2 exactly
  - Welch's t-test                  average order value / revenue per recipient,
                                    no equal-variance assumption
  - Sample-size calculators         n per group for proportions and for means
  - Achieved power / detectable MDE  what the test you already ran could see
  - Sample ratio mismatch (SRM)     was the split actually 50/50?
"""

from __future__ import annotations

import math
from statistics import NormalDist, mean, stdev

_N = NormalDist()


# --------------------------------------------------------------------------- #
# Distribution helpers
# --------------------------------------------------------------------------- #
def normal_cdf(x: float) -> float:
    return _N.cdf(x)


def normal_ppf(p: float) -> float:
    return _N.inv_cdf(p)


def _betacf(a: float, b: float, x: float) -> float:
    """Continued fraction for the incomplete beta function (Lentz's method)."""
    max_iter, eps, tiny = 300, 3e-14, 1e-300
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c, d = 1.0, 1.0 - qab * x / qap
    d = 1.0 / (d if abs(d) > tiny else tiny)
    h = d
    for m in range(1, max_iter + 1):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        d = 1.0 / (d if abs(d) > tiny else tiny)
        c = 1.0 + aa / (c if abs(c) > tiny else tiny)
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        d = 1.0 / (d if abs(d) > tiny else tiny)
        c = 1.0 + aa / (c if abs(c) > tiny else tiny)
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < eps:
            break
    return h


def regularized_incomplete_beta(a: float, b: float, x: float) -> float:
    """I_x(a, b) - the CDF of a Beta(a, b) distribution evaluated at x."""
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    ln_front = math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log(1.0 - x)
    front = math.exp(ln_front)
    if x < (a + 1.0) / (a + b + 2.0):
        return front * _betacf(a, b, x) / a
    return 1.0 - front * _betacf(b, a, 1.0 - x) / b


def student_t_cdf(t: float, df: float) -> float:
    """P(T <= t) for Student's t with ``df`` degrees of freedom."""
    if df <= 0:
        raise ValueError("df must be positive")
    if math.isinf(t):
        return 1.0 if t > 0 else 0.0
    x = df / (df + t * t)
    tail = 0.5 * regularized_incomplete_beta(df / 2.0, 0.5, x)  # P(T > |t|)
    return 1.0 - tail if t >= 0 else tail


def student_t_ppf(p: float, df: float) -> float:
    """Inverse CDF of Student's t by bisection (plenty precise for CIs)."""
    if not 0.0 < p < 1.0:
        raise ValueError("p must be in (0, 1)")
    lo, hi = -1000.0, 1000.0
    for _ in range(200):
        mid = (lo + hi) / 2.0
        if student_t_cdf(mid, df) < p:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2.0


def chi_square_1df_sf(stat: float) -> float:
    """Survival function (p-value) of a chi-square with 1 df: exact via the normal."""
    return 2.0 * (1.0 - normal_cdf(math.sqrt(max(stat, 0.0))))


# --------------------------------------------------------------------------- #
# Tests
# --------------------------------------------------------------------------- #
def two_proportion_ztest(x1: int, n1: int, x2: int, n2: int, alpha: float = 0.05) -> dict:
    """Two-sided z-test for p2 - p1, plus a (1 - alpha) Wald confidence interval.

    Group 1 is the control, group 2 the treatment, so a positive ``diff`` means
    the treatment converted better.
    """
    if min(n1, n2) <= 0:
        raise ValueError("both groups need at least one observation")
    p1, p2 = x1 / n1, x2 / n2
    diff = p2 - p1
    pooled = (x1 + x2) / (n1 + n2)
    se_pooled = math.sqrt(pooled * (1 - pooled) * (1 / n1 + 1 / n2))
    se_unpooled = math.sqrt(p1 * (1 - p1) / n1 + p2 * (1 - p2) / n2)
    z = diff / se_pooled if se_pooled > 0 else 0.0
    p_value = 2.0 * (1.0 - normal_cdf(abs(z)))
    z_crit = normal_ppf(1 - alpha / 2)
    return {
        "p1": p1,
        "p2": p2,
        "diff": diff,
        "relative_lift": diff / p1 if p1 > 0 else float("nan"),
        "se_pooled": se_pooled,
        "se_unpooled": se_unpooled,
        "z": z,
        "p_value": p_value,
        "ci_low": diff - z_crit * se_unpooled,
        "ci_high": diff + z_crit * se_unpooled,
        "significant": p_value < alpha,
        "alpha": alpha,
    }


def chi_square_2x2(x1: int, n1: int, x2: int, n2: int, alpha: float = 0.05) -> dict:
    """Pearson chi-square test of independence on the 2x2 converted/not table."""
    table = [[x1, n1 - x1], [x2, n2 - x2]]
    row_totals = [sum(r) for r in table]
    col_totals = [table[0][0] + table[1][0], table[0][1] + table[1][1]]
    grand = sum(row_totals)
    stat = 0.0
    for i in range(2):
        for j in range(2):
            expected = row_totals[i] * col_totals[j] / grand
            if expected > 0:
                stat += (table[i][j] - expected) ** 2 / expected
    p_value = chi_square_1df_sf(stat)
    return {"statistic": stat, "df": 1, "p_value": p_value, "significant": p_value < alpha, "alpha": alpha}


def welch_ttest(sample_a: list[float], sample_b: list[float], alpha: float = 0.05) -> dict:
    """Welch's unequal-variance t-test for mean(b) - mean(a) with a CI."""
    n1, n2 = len(sample_a), len(sample_b)
    if n1 < 2 or n2 < 2:
        return {
            "n1": n1, "n2": n2,
            "mean1": mean(sample_a) if n1 else float("nan"),
            "mean2": mean(sample_b) if n2 else float("nan"),
            "sd1": float("nan"), "sd2": float("nan"), "diff": float("nan"), "se": float("nan"),
            "t": float("nan"), "df": float("nan"), "p_value": float("nan"),
            "ci_low": float("nan"), "ci_high": float("nan"), "significant": False, "alpha": alpha,
        }
    m1, m2 = mean(sample_a), mean(sample_b)
    s1, s2 = stdev(sample_a), stdev(sample_b)
    v1, v2 = s1 * s1 / n1, s2 * s2 / n2
    se = math.sqrt(v1 + v2)
    diff = m2 - m1
    if se == 0:
        t_stat, df, p_value, t_crit = 0.0, float(n1 + n2 - 2), 1.0, 0.0
    else:
        t_stat = diff / se
        df = (v1 + v2) ** 2 / (v1 * v1 / (n1 - 1) + v2 * v2 / (n2 - 1))
        p_value = 2.0 * (1.0 - student_t_cdf(abs(t_stat), df))
        t_crit = student_t_ppf(1 - alpha / 2, df)
    return {
        "n1": n1, "n2": n2, "mean1": m1, "mean2": m2, "sd1": s1, "sd2": s2,
        "diff": diff, "relative_lift": diff / m1 if m1 else float("nan"), "se": se,
        "t": t_stat, "df": df, "p_value": p_value,
        "ci_low": diff - t_crit * se, "ci_high": diff + t_crit * se,
        "significant": p_value < alpha, "alpha": alpha,
    }


def sample_ratio_mismatch(n1: int, n2: int, expected_share_2: float = 0.5) -> dict:
    """Chi-square goodness-of-fit: did the split land near the planned ratio?

    A tiny p-value here means the randomisation or the data export is broken
    and every other result in the report is suspect.
    """
    total = n1 + n2
    e2 = total * expected_share_2
    e1 = total - e2
    stat = (n1 - e1) ** 2 / e1 + (n2 - e2) ** 2 / e2 if total else 0.0
    p_value = chi_square_1df_sf(stat)
    return {"observed_share_2": n2 / total if total else float("nan"), "statistic": stat, "p_value": p_value,
            "mismatch": p_value < 0.001}


# --------------------------------------------------------------------------- #
# Planning: sample size, power, duration
# --------------------------------------------------------------------------- #
def required_sample_size_proportions(p1: float, p2: float, alpha: float = 0.05, power: float = 0.8) -> int:
    """Recipients needed PER GROUP to detect p1 -> p2 (two-sided test).

    n = [ z_{1-a/2} * sqrt(2 * pbar * (1 - pbar)) + z_{power} * sqrt(p1(1-p1) + p2(1-p2)) ]^2 / (p2 - p1)^2
    """
    if p1 == p2:
        return math.inf
    pbar = (p1 + p2) / 2
    z_a = normal_ppf(1 - alpha / 2)
    z_b = normal_ppf(power)
    numerator = (z_a * math.sqrt(2 * pbar * (1 - pbar)) + z_b * math.sqrt(p1 * (1 - p1) + p2 * (1 - p2))) ** 2
    return math.ceil(numerator / (p2 - p1) ** 2)


def required_sample_size_means(sd: float, mde_abs: float, alpha: float = 0.05, power: float = 0.8) -> int:
    """Observations needed PER GROUP to detect a difference in means of ``mde_abs``."""
    if mde_abs == 0:
        return math.inf
    z_a = normal_ppf(1 - alpha / 2)
    z_b = normal_ppf(power)
    return math.ceil(2 * ((z_a + z_b) * sd / mde_abs) ** 2)


def achieved_power_proportions(p1: float, p2: float, n_per_group: int, alpha: float = 0.05) -> float:
    """Probability the test you actually ran would flag a true p1 -> p2 change."""
    if n_per_group <= 0:
        return 0.0
    pbar = (p1 + p2) / 2
    se0 = math.sqrt(2 * pbar * (1 - pbar) / n_per_group)
    se1 = math.sqrt((p1 * (1 - p1) + p2 * (1 - p2)) / n_per_group)
    if se1 == 0:
        return 1.0
    z_a = normal_ppf(1 - alpha / 2)
    delta = abs(p2 - p1)
    return (1 - normal_cdf((z_a * se0 - delta) / se1)) + normal_cdf((-z_a * se0 - delta) / se1)


def minimum_detectable_lift(p1: float, n_per_group: int, alpha: float = 0.05, power: float = 0.8) -> float:
    """Smallest absolute increase in conversion the given sample could reliably detect."""
    lo, hi = 0.0, 1.0 - p1
    for _ in range(100):
        mid = (lo + hi) / 2
        if achieved_power_proportions(p1, p1 + mid, n_per_group, alpha) < power:
            lo = mid
        else:
            hi = mid
    return hi


def days_to_run(n_per_group: int, daily_recipients: float, treatment_share: float = 0.5) -> int:
    """Calendar days until BOTH groups reach ``n_per_group``, given a daily send volume."""
    if daily_recipients <= 0 or n_per_group == math.inf:
        return math.inf
    smaller_share = min(treatment_share, 1 - treatment_share)
    return math.ceil(n_per_group / (daily_recipients * smaller_share))
