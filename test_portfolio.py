import unittest
import pandas as pd
from portfolio import portfolio_chart

class PortfolioTests(unittest.TestCase):
    def test_entitlement_added_only_on_start_date(self):
        accounts = pd.DataFrame([
            dict(customer_id=cid, contract_start=pd.Timestamp(start),
                 contract_end=pd.Timestamp("2026-12-31"), annual_entitlement_credits=credits,
                 total_contract_days=364, forecast_available=True, trailing_30_day_average_daily_usage=0)
            for cid, start, credits in [("A", "2026-01-05", 100), ("B", "2026-02-01", 200)]])
        usage = pd.DataFrame({"customer_id": ["A"], "date": [pd.Timestamp("2026-01-05")], "credits_used": [0]})
        chart = portfolio_chart(accounts, usage, pd.Timestamp("2026-01-01"),
                                pd.Timestamp("2026-03-01"), pd.Timestamp("2026-01-31")).set_index("date")
        for date, expected in [("2026-01-01", 0), ("2026-01-05", 100), ("2026-01-31", 100)]:
            self.assertEqual(chart.loc[date, "Total contracted credits"], expected)
        self.assertTrue(pd.isna(chart.loc["2026-02-01", "Total contracted credits"]))
        self.assertEqual(chart.loc["2026-01-31", "Projected contracted credits"], 100)
        self.assertGreater(chart.loc["2026-02-01", "Projected contracted credits"], 100)
        self.assertLess(chart.loc["2026-02-01", "Projected contracted credits"], 200)
        self.assertEqual(chart.loc["2026-03-01", "Projected contracted credits"], 200)
        self.assertEqual(chart.attrs["average_monthly_contracted_addition"], 100)
        accounts.loc[accounts.customer_id.eq("B"), "forecast_available"] = False
        chart = portfolio_chart(accounts, usage, pd.Timestamp("2026-01-01"),
                                pd.Timestamp("2026-03-01"), pd.Timestamp("2026-01-31"))
        self.assertTrue(chart.loc[chart.date.gt(pd.Timestamp("2026-02-01")),
                                  "Projected consumption"].notna().all())

    def test_projection_stops_at_contract_end_and_pacing_matches(self):
        accounts = pd.DataFrame([dict(customer_id="A", contract_start=pd.Timestamp("2026-01-01"),
            contract_end=pd.Timestamp("2026-01-05"), annual_entitlement_credits=100,
            total_contract_days=4, forecast_available=True, trailing_30_day_average_daily_usage=10)])
        usage = pd.DataFrame(dict(customer_id=["A","A"], date=pd.to_datetime(["2026-01-01","2026-01-02"]), credits_used=[10,10]))
        chart = portfolio_chart(accounts, usage, pd.Timestamp("2026-01-01"), pd.Timestamp("2026-01-08"), pd.Timestamp("2026-01-02"))
        self.assertEqual(chart.iloc[0]["Credits consumed"], 0)
        self.assertEqual(chart.iloc[2]["Credits consumed"], 20)
        self.assertTrue(pd.isna(chart.iloc[3]["Credits consumed"]))
        self.assertEqual(chart.iloc[-1]["Projected consumption"], 40)
        self.assertEqual(chart.iloc[-1]["Contract pacing"], 100)
        accounts.loc[0, "forecast_available"] = False
        chart = portfolio_chart(accounts, usage, pd.Timestamp("2026-01-01"), pd.Timestamp("2026-01-08"), pd.Timestamp("2026-01-02"))
        self.assertTrue(chart["Projected consumption"].isna().all())
