"""Regression and boundary checks: python -m unittest -v"""
from copy import deepcopy
from dataclasses import replace
import json
from pathlib import Path
import unittest
from unittest.mock import patch

import pandas as pd

from finance import (Rules, normalize, load_workbook, calculate_accounts, pacing_chart_data,
                     classify, reconcile_source_import)
from briefs import account_facts, deterministic_brief, validate_brief, generate_ai_brief

SOURCE = Path.home() / "Downloads" / "Product Finance - AI Fluency Test Sample Data.xlsx"


def synthetic(start="2026-01-01", dates=None, credits=None, entitlement=36500):
    c = pd.DataFrame([dict(customer_id="A", contract_start=start, term_months=12,
                           annual_entitlement_credits=entitlement, annual_contract_value_usd=entitlement/2)])
    u = pd.DataFrame(dict(customer_id=["A"] * len(dates), date=dates, credits_used=credits))
    return normalize(c, u)


class FinanceTests(unittest.TestCase):
    def test_fractional_exact_entitlement_and_chart_reconcile(self):
        c, u = synthetic(dates=pd.date_range("2026-01-01", periods=10), credits=[.1]*10, entitlement=1)
        m = calculate_accounts(c, u).iloc[0]
        self.assertEqual(m.total_credits_used, 1)
        self.assertEqual(m.status, "OVER ENTITLEMENT")
        self.assertEqual(m.estimated_exhaustion_date, pd.Timestamp("2026-01-10"))
        self.assertEqual(m.overage_credits, 0)
        self.assertEqual(pacing_chart_data(m, u)["Actual cumulative usage"].dropna().iloc[-1], 1)

    def test_sparse_usage_events_are_forecast_and_classified(self):
        c, u = synthetic(dates=pd.date_range("2026-07-02", periods=30), credits=[100]*30)
        m = calculate_accounts(c, u).iloc[0]
        self.assertEqual(m.status, "UNDERUTILIZING")
        self.assertTrue(m.usage_accelerating)
        self.assertTrue(m.forecast_available)
        self.assertAlmostEqual(m.projected_total_contract_usage, 18_300)
        self.assertNotEqual(account_facts(m)["projected_utilization"], "unavailable")

    def test_observed_overage_remains_urgent_despite_gaps(self):
        c, u = synthetic(dates=["2026-07-31"], credits=[120], entitlement=100)
        m = calculate_accounts(c, u).iloc[0]
        self.assertEqual(m.status, "OVER ENTITLEMENT")
        self.assertEqual(m.estimated_overage_value, 10)
        self.assertAlmostEqual(m.projected_overage_value, 316)

    def test_explicit_zero_days_and_expired_complete_contract(self):
        c, u = synthetic(dates=pd.date_range("2026-01-01", "2026-12-31"), credits=[0]*365)
        m = calculate_accounts(c, u).iloc[0]
        self.assertEqual(m.status, "UNDERUTILIZING")
        self.assertEqual(m.contract_lifecycle, "EXPIRED")
        self.assertEqual(m.forecast_remaining_usage, 0)
        self.assertEqual(m.usage_growth_30_day_pct, 0)

    def test_numeric_dates_rejected(self):
        with self.assertRaisesRegex(ValueError, "numeric dates"):
            synthetic(start=46000, dates=[46001], credits=[1])

    def test_threshold_validation(self):
        for kwargs in [dict(over_entitlement_threshold=.9), dict(early_exhaustion_days=.5),
                       dict(underutilization_min_elapsed=2), dict(acceleration_min_history_days=30)]:
            with self.assertRaises(ValueError):
                Rules(**kwargs)

    def test_constant_usage_paces_exactly_and_windows_do_not_overlap(self):
        c, u = synthetic(dates=pd.date_range("2026-01-01", periods=60), credits=[100]*60)
        m = calculate_accounts(c, u).iloc[0]
        self.assertEqual(m.days_elapsed, 60)
        self.assertEqual(m.days_remaining, 305)
        self.assertEqual(m.trailing_30_day_credits, 3000)
        self.assertEqual(m.prior_30_day_credits, 3000)
        self.assertAlmostEqual(m.pacing_index, 1)
        self.assertAlmostEqual(m.projected_utilization_pct, 1)
        self.assertEqual(m.estimated_exhaustion_date, pd.Timestamp("2026-12-31"))
        self.assertEqual(m.status, "ON TRACK")
        chart = pacing_chart_data(m, u)
        self.assertEqual(chart.iloc[0]["Expected linear pacing"], 0)
        self.assertEqual(chart.iloc[-1]["Expected linear pacing"], 36500)
        self.assertEqual(chart["Actual cumulative usage"].dropna().iloc[-1], 6000)

    def test_short_history_uses_contract_days(self):
        c, u = synthetic(dates=pd.date_range("2026-01-01", periods=5), credits=[100]*5)
        m = calculate_accounts(c, u).iloc[0]
        self.assertEqual(m.trailing_window_days, 5)
        self.assertEqual(m.trailing_30_day_average_daily_usage, 100)
        self.assertTrue(pd.isna(m.usage_growth_30_day_pct))
        self.assertFalse(m.usage_accelerating)

    def test_zero_usage_and_missing_days(self):
        c, u = synthetic(dates=["2026-04-02"], credits=[0])
        m = calculate_accounts(c, u).iloc[0]
        self.assertEqual(m.missing_usage_days, 91)
        self.assertEqual(m.status, "UNDERUTILIZING")
        self.assertEqual(m.usage_growth_30_day_pct, 0)
        self.assertTrue(pd.isna(m.days_to_exhaustion))
        self.assertTrue(pd.isna(m.estimated_exhaustion_date))

    def test_future_expired_and_no_usage_customer(self):
        c, u = synthetic(dates=["2026-07-31"], credits=[0])
        future = c.iloc[0].copy()
        future["customer_id"], future["contract_start"] = "F", pd.Timestamp("2026-08-01")
        expired = c.iloc[0].copy()
        expired["customer_id"], expired["contract_start"] = "E", pd.Timestamp("2025-01-01")
        c, u = normalize(pd.concat([c, pd.DataFrame([future, expired])]), u)
        a = calculate_accounts(c, u).set_index("customer_id")
        self.assertEqual(a.loc["F", "contract_lifecycle"], "NOT STARTED")
        self.assertTrue(pd.isna(a.loc["F", "pacing_index"]))
        self.assertTrue(pd.isna(a.loc["F", "forecast_remaining_usage"]))
        self.assertEqual(a.loc["F", "status"], "NOT STARTED")
        self.assertEqual(a.loc["E", "contract_lifecycle"], "EXPIRED")
        self.assertEqual(a.loc["E", "days_elapsed"], 365)
        self.assertEqual(a.loc["E", "forecast_remaining_usage"], 0)

    def test_active_contract_without_usage_rows_remains_visible(self):
        c, u = synthetic(dates=["2026-07-31"], credits=[0])
        no_events = c.iloc[0].copy()
        no_events["customer_id"] = "CUST-17"
        c, u = normalize(pd.concat([c, pd.DataFrame([no_events])]), u)
        accounts = calculate_accounts(c, u).set_index("customer_id")
        self.assertIn("CUST-17", accounts.index)
        self.assertEqual(accounts.loc["CUST-17", "total_credits_used"], 0)
        self.assertTrue(accounts.loc["CUST-17", "forecast_available"])
        self.assertEqual(accounts.loc["CUST-17", "status"], "UNDERUTILIZING")

    def test_exhausted_date_and_exact_entitlement(self):
        c, u = synthetic(dates=["2026-01-01", "2026-01-02"], credits=[60, 40], entitlement=100)
        m = calculate_accounts(c, u).iloc[0]
        self.assertEqual(m.status, "OVER ENTITLEMENT")
        self.assertEqual(m.overage_credits, 0)
        self.assertEqual(m.days_to_exhaustion, 0)
        self.assertEqual(m.estimated_exhaustion_date, pd.Timestamp("2026-01-02"))

    def test_acceleration_and_zero_baseline(self):
        for baseline in (0, 100):
            c, u = synthetic(dates=pd.date_range("2026-01-01", periods=60), credits=[baseline]*30+[200]*30)
            self.assertTrue(calculate_accounts(c, u).iloc[0].usage_accelerating)
        c, u = synthetic(dates=pd.date_range("2026-01-01", periods=60), credits=[1]*30+[2]*30)
        self.assertFalse(calculate_accounts(c, u).iloc[0].usage_accelerating)

    def test_rule_strict_boundaries(self):
        c, u = synthetic(dates=pd.date_range("2026-01-01", periods=100), credits=[100]*100)
        m = calculate_accounts(c, u).iloc[0].to_dict()
        m.update(projected_utilization_pct=1.10, estimated_exhaustion_date=m["contract_end"]-pd.Timedelta(days=30))
        self.assertEqual(classify(m, Rules())["status"], "ON TRACK")
        m["estimated_exhaustion_date"] -= pd.Timedelta(days=1)
        self.assertEqual(classify(m, Rules())["status"], "EARLY EXHAUSTION RISK")
        m.update(estimated_exhaustion_date=pd.NaT, projected_utilization_pct=.70)
        self.assertEqual(classify(m, Rules())["status"], "ON TRACK")
        m["projected_utilization_pct"] = .6999
        self.assertEqual(classify(m, Rules())["status"], "UNDERUTILIZING")

    def test_leap_year(self):
        c, u = synthetic(start="2024-02-29", dates=["2024-02-29"], credits=[100])
        m = calculate_accounts(c, u).iloc[0]
        self.assertEqual(m.contract_end, pd.Timestamp("2025-02-28"))
        self.assertEqual(m.total_contract_days, 365)

    def test_invalid_source_rejected(self):
        c, u = synthetic(dates=["2026-01-01"], credits=[100])
        for field, bad in [("annual_entitlement_credits", 0), ("annual_contract_value_usd", -1), ("term_months", 24)]:
            copy = c.copy()
            copy[field] = bad
            with self.assertRaises(ValueError):
                normalize(copy, u)
        with self.assertRaises(ValueError):
            normalize(c, pd.concat([u, u]))
        for field, bad in [("customer_id", "UNKNOWN"), ("date", "2027-01-01"), ("credits_used", -1)]:
            copy = u.copy()
            copy[field] = bad
            with self.assertRaises(ValueError):
                normalize(c, copy)

    @unittest.skipUnless(SOURCE.exists(), "Supplied workbook not present")
    def test_source_manual_regressions(self):
        c, u = load_workbook(SOURCE.read_bytes())
        a = calculate_accounts(c, u).set_index("customer_id")
        self.assertEqual(len(c), 12)
        self.assertEqual(len(u), 1943)
        self.assertAlmostEqual(a.total_credits_used.sum(), 912704.1)
        self.assertEqual(a.annual_entitlement_credits.sum(), 1752000)
        checks = {
            "CUST-03": (180000.7, 43905.8, 450753.1333333333, "OVER ENTITLEMENT"),
            "CUST-02": (224399.3, 12114.4, 290624.6866666667, "EARLY EXHAUSTION RISK"),
            "CUST-04": (6287.2, 1195.2, 14733.28, "UNDERUTILIZING"),
            "CUST-01": (129928.5, 20082.3, 235025.87, "ON TRACK"),
        }
        for cid, (total, recent, projected, status) in checks.items():
            self.assertAlmostEqual(a.loc[cid, "total_credits_used"], total)
            self.assertAlmostEqual(a.loc[cid, "trailing_30_day_credits"], recent)
            self.assertAlmostEqual(a.loc[cid, "projected_total_contract_usage"], projected)
            self.assertEqual(a.loc[cid, "status"], status)
        self.assertAlmostEqual(a.estimated_overage_value.sum(), 30000.35)
        lower = calculate_accounts(c, u, replace(Rules(), usage_acceleration_threshold=.20)).set_index("customer_id")
        self.assertTrue(lower.loc["CUST-05", "usage_accelerating"])

        audit = reconcile_source_import(SOURCE.read_bytes(), c, u, a.reset_index())
        self.assertTrue(audit["passed"])
        self.assertEqual(audit["usage_min"], pd.Timestamp("2026-01-05"))
        self.assertEqual(audit["usage_max"], pd.Timestamp("2026-07-31"))
        self.assertEqual(audit["failed_checks"], 0)
        changed = a.reset_index().copy()
        changed.loc[changed.customer_id.eq("CUST-01"), "total_credits_used"] += 1
        failed = reconcile_source_import(SOURCE.read_bytes(), c, u, changed)
        self.assertFalse(failed["passed"])
        self.assertIn("CUST-01", next(
            check["Imported / Calculated"] for check in failed["checks"]
            if check["Check"] == "Credits used by customer"))

    def test_ai_fact_substitution_and_invalid_output(self):
        c, u = synthetic(dates=["2026-01-01"], credits=[100])
        m = calculate_accounts(c, u).iloc[0].to_dict()
        before = deepcopy(m)
        facts = account_facts(m)
        output = dict(what_happened="Usage is {used}.", why_it_matters="{reason}", impact="Overage proxy: {overage_value}.", action="{owner}: {action}")
        self.assertEqual(len(validate_brief(json.dumps(output), facts)), 4)
        self.assertEqual(len(deterministic_brief(m)), 4)
        with patch("openai.OpenAI") as client:
            client.return_value.responses.create.return_value.output_text = json.dumps(output)
            self.assertEqual(len(generate_ai_brief(m)), 4)
            sent = client.return_value.responses.create.call_args.kwargs
            self.assertEqual(json.loads(sent["input"]), facts)
            self.assertFalse(sent["store"])
            output.update(impact="A refund is guaranteed.", action="Cancel the contract.")
            client.return_value.responses.create.return_value.output_text = json.dumps(output)
            result = generate_ai_brief(m)
            self.assertEqual(result[2:], deterministic_brief(m)[2:])
        self.assertEqual(m["total_credits_used"], before["total_credits_used"])
        for bad in ["Overage is $999.", "Overage is one million.", "{unknown}"]:
            output["impact"] = bad
            with self.assertRaises(ValueError):
                validate_brief(json.dumps(output), facts)


if __name__ == "__main__":
    unittest.main()
