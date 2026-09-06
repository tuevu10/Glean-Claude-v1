import os
from pathlib import Path
import unittest
import tempfile
from unittest.mock import patch

from streamlit.testing.v1 import AppTest


@unittest.skipUnless((Path.home()/"Downloads"/"Product Finance - AI Fluency Test Sample Data.xlsx").exists(), "Sample workbook required")
class AppTests(unittest.TestCase):
    def test_resolve_and_reopen_priority(self):
        app = AppTest.from_file("app.py", default_timeout=30).run()
        toggles = [t for t in app.button if t.key and t.key.startswith("resolved_")]
        self.assertGreater(len(toggles), 1)
        key = toggles[0].key
        toggles[0].click().run()
        self.assertTrue(app.session_state["priority_resolutions"])
        self.assertEqual([t.key for t in app.button if t.key and t.key.startswith("resolved_")][-1], key)
        app.button(key=key).click().run()
        self.assertFalse(app.session_state["priority_resolutions"])
        self.assertEqual([t.key for t in app.button if t.key and t.key.startswith("resolved_")][0], key)
        self.assertFalse(app.exception)

    def test_summary_period_selection(self):
        from finance import load_workbook, credit_sum
        import pandas as pd
        source = Path.home()/"Downloads"/"Product Finance - AI Fluency Test Sample Data.xlsx"
        _, usage = load_workbook(source.read_bytes())
        app = AppTest.from_file("app.py", default_timeout=30).run()
        app.selectbox(key="summary_view").set_value("Quarterly").run()
        app.selectbox(key="summary_quarter").set_value(2).run()
        expected = credit_sum(usage.loc[(usage.date >= pd.Timestamp("2026-04-01")) &
                                       (usage.date < pd.Timestamp("2026-07-01")), "credits_used"])
        self.assertEqual(app.metric[4].value, f"{expected:,.1f}")
        self.assertFalse(app.exception)
        app.selectbox(key="summary_view").set_value("Monthly").run()
        app.selectbox(key="summary_month").set_value(12).run()
        self.assertTrue(any("No usage records" in i.value for i in app.info))
        self.assertFalse(app.exception)

    def test_priority_card_opens_details_tab(self):
        app = AppTest.from_file("app.py", default_timeout=30).run()
        self.assertEqual(app.session_state["workspace_view"], "Dashboard")
        next(b for b in app.button if b.label == "Review CUST-02").click().run()
        self.assertEqual(app.session_state["workspace_view"], "Customer Details")
        self.assertEqual(app.selectbox(key="detail_customer").value, "CUST-02")
        self.assertFalse(app.exception)

    def test_refresh_rejects_invalid_replacement_and_recovers(self):
        original = (Path.home()/"Downloads"/"Product Finance - AI Fluency Test Sample Data.xlsx").read_bytes()
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder)/"input.xlsx"
            source.write_bytes(original)
            with patch.dict(os.environ, {"FINANCE_WORKBOOK": str(source), "OPENAI_API_KEY": ""}):
                app = AppTest.from_file("app.py", default_timeout=30).run()
                self.assertFalse(app.exception)
                self.assertEqual(app.metric[2].value, "$30,000.35")
                source.write_bytes(b"invalid replacement workbook")
                next(b for b in app.button if b.label == "Refresh source").click().run()
                self.assertTrue(app.error)
                self.assertEqual(len(app.metric), 0)
                source.write_bytes(original)
                next(b for b in app.button if b.label == "Refresh source").click().run()
                self.assertFalse(app.exception)
                self.assertFalse(app.error)
                self.assertEqual(app.metric[2].value, "$30,000.35")

    def test_dashboard_filters_details_and_thresholds(self):
        with patch.dict(os.environ, {"OPENAI_API_KEY": ""}):
            app = AppTest.from_file("app.py", default_timeout=30).run()
            self.assertFalse(app.exception)
            self.assertEqual(app.metric[0].value, "1")
            self.assertEqual(app.metric[3].value, "1,752,000")
            self.assertEqual(len(app.dataframe[0].value), 8)
            self.assertEqual(app.selectbox(key="detail_customer").value, "CUST-03")
            app.toggle(key="exceptions_only").set_value(False).run()
            self.assertEqual(len(app.dataframe[0].value), 12)
            app.multiselect[0].set_value(["CUST-03"]).run()
            self.assertEqual(len(app.dataframe[0].value), 1)
            app.multiselect[1].set_value(["ON TRACK"]).run()
            self.assertEqual(len(app.dataframe[0].value), 0)
            self.assertFalse(app.exception)
            app.multiselect[0].set_value([])
            app.multiselect[1].set_value([])
            app.selectbox(key="detail_customer").set_value("CUST-03").run()
            self.assertEqual(app.metric[6].value, "180,000.7")
            app.number_input[3].set_value(20.0).run()
            queue = app.dataframe[1].value.set_index("Customer")
            self.assertTrue(queue.loc["CUST-05", "Usage Accelerating"])
            self.assertFalse(app.exception)
            self.assertFalse(any(b.label == "Generate Finance Brief" for b in app.button))

    def test_manual_navigation_survives_stale_row_selection(self):
        app = AppTest.from_file("app.py", default_timeout=30).run()
        queue_key = next(k for k in app.session_state.filtered_state if k.startswith("queue_"))
        # Replay the same selection state delivered by the dataframe widget.
        app.session_state[queue_key] = {"selection": {"rows": [1], "columns": [], "cells": []}}
        app.run()
        # Callback is verified in the browser; manual selection must not be reset by a stale row.
        app.selectbox(key="detail_customer").set_value("CUST-01").run()
        self.assertEqual(app.selectbox(key="detail_customer").value, "CUST-01")
        app.toggle(key="exceptions_only").set_value(False).run()
        self.assertEqual(app.selectbox(key="detail_customer").value, "CUST-01")
        self.assertFalse(app.exception)

    def test_ai_failure_falls_back(self):
        with patch.dict(os.environ, {"OPENAI_API_KEY": "test-placeholder"}), patch("briefs.generate_ai_brief", side_effect=RuntimeError("simulated failure")):
            app = AppTest.from_file("app.py", default_timeout=30).run()
            next(b for b in app.button if b.label == "Generate Finance Brief").click().run()
            self.assertFalse(app.exception)
            self.assertTrue(any("failed validation" in w.value for w in app.warning))

    def test_invalid_workbook_path(self):
        app = AppTest.from_file("app.py", default_timeout=30).run()
        app.text_input[0].set_value("does-not-exist.xlsx").run()
        self.assertFalse(app.exception)
        self.assertTrue(app.error)
        self.assertEqual(len(app.metric), 0)
