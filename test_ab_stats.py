"""Unit tests for ab_stats.py against textbook / scipy reference values.

Run with:  python -m unittest test_ab_stats.py
"""

import math
import unittest

from ab_stats import (
    achieved_power_proportions,
    chi_square_2x2,
    days_to_run,
    minimum_detectable_lift,
    required_sample_size_means,
    required_sample_size_proportions,
    sample_ratio_mismatch,
    student_t_cdf,
    student_t_ppf,
    two_proportion_ztest,
    welch_ttest,
)


class DistributionTests(unittest.TestCase):
    def test_t_cdf_matches_tables(self):
        self.assertAlmostEqual(student_t_cdf(2.0, 10), 0.96331, places=5)
        self.assertAlmostEqual(student_t_cdf(0.0, 5), 0.5, places=10)
        self.assertAlmostEqual(student_t_cdf(-2.228, 10), 0.025, places=3)
        self.assertAlmostEqual(student_t_cdf(1.96, 1e6), 0.975, places=4)  # -> normal

    def test_t_ppf_inverts_cdf(self):
        self.assertAlmostEqual(student_t_ppf(0.975, 10), 2.2281, places=4)
        self.assertAlmostEqual(student_t_ppf(0.95, 3), 2.3534, places=4)


class ProportionTests(unittest.TestCase):
    def test_ztest_known_example(self):
        r = two_proportion_ztest(150, 5000, 195, 5000)
        self.assertAlmostEqual(r["z"], 2.4656, places=4)
        self.assertAlmostEqual(r["p_value"], 0.01368, places=5)
        self.assertAlmostEqual(r["relative_lift"], 0.30, places=10)
        self.assertLess(r["ci_low"], r["diff"])
        self.assertGreater(r["ci_high"], r["diff"])
        self.assertTrue(r["significant"])

    def test_chi_square_equals_z_squared(self):
        z = two_proportion_ztest(150, 5000, 195, 5000)
        chi = chi_square_2x2(150, 5000, 195, 5000)
        self.assertAlmostEqual(chi["statistic"], z["z"] ** 2, places=9)
        self.assertAlmostEqual(chi["p_value"], z["p_value"], places=9)

    def test_no_difference_is_not_significant(self):
        r = two_proportion_ztest(100, 4000, 100, 4000)
        self.assertEqual(r["z"], 0.0)
        self.assertAlmostEqual(r["p_value"], 1.0)
        self.assertFalse(r["significant"])

    def test_sample_ratio_mismatch(self):
        self.assertFalse(sample_ratio_mismatch(4937, 4863)["mismatch"])
        self.assertTrue(sample_ratio_mismatch(4400, 5600)["mismatch"])


class MeanTests(unittest.TestCase):
    def test_welch_matches_scipy(self):
        # scipy.stats.ttest_ind([1,2,3,4,5],[2,3,4,5,7], equal_var=False) -> t=1.0776, p=0.3138
        r = welch_ttest([1, 2, 3, 4, 5.0], [2, 3, 4, 5, 7.0])
        self.assertAlmostEqual(r["t"], 1.0776, places=4)
        self.assertAlmostEqual(r["p_value"], 0.31375, places=4)
        self.assertAlmostEqual(r["df"], 7.7111, places=3)

    def test_welch_degenerate_samples(self):
        r = welch_ttest([10.0], [12.0, 14.0])
        self.assertTrue(math.isnan(r["p_value"]))
        self.assertFalse(r["significant"])


class PlanningTests(unittest.TestCase):
    def test_sample_size_textbook_value(self):
        # 2.0% -> 2.4%, alpha 0.05, power 0.8: standard calculators give ~21,100 per group
        n = required_sample_size_proportions(0.02, 0.024)
        self.assertTrue(21000 <= n <= 21200, n)

    def test_sample_size_for_means(self):
        # 2 * ((1.96 + 0.8416) * 30 / 5)^2 = 565.2 -> 566
        self.assertEqual(required_sample_size_means(30, 5), 566)

    def test_power_and_mde_are_consistent(self):
        n = required_sample_size_proportions(0.03, 0.039, 0.05, 0.8)
        self.assertGreaterEqual(achieved_power_proportions(0.03, 0.039, n), 0.80)
        mde = minimum_detectable_lift(0.03, n)
        self.assertAlmostEqual(mde, 0.009, places=3)

    def test_days_to_run_uses_smaller_group(self):
        self.assertEqual(days_to_run(7000, 1000, 0.5), 14)
        self.assertEqual(days_to_run(7000, 1000, 0.2), 35)
        self.assertEqual(days_to_run(7000, 0), math.inf)


if __name__ == "__main__":
    unittest.main()
