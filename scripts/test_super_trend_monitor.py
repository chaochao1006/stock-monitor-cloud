"""Offline checks: original algorithm parity, lookahead, sessions, stale/error states."""
import ast
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import numpy as np
import pandas as pd

import super_trend_monitor as monitor


class MonitorTests(unittest.TestCase):
    def prices(self):
        rng = np.random.default_rng(42)
        close = 100 + np.cumsum(rng.normal(0, 4, 300))
        return pd.DataFrame({"Open": close, "High": close + 2, "Low": close - 2,
                             "Close": close, "Adjusted Close": close, "Volume": 1000},
                            index=pd.bdate_range("2024-01-01", periods=300))

    def test_prefix_and_flips(self):
        data = self.prices()
        full = monitor.calculate_supertrend(data)
        prefix = monitor.calculate_supertrend(data.iloc[:150])
        pd.testing.assert_frame_equal(prefix, full.iloc[:len(prefix)])
        self.assertGreater((full.event == "买入").sum(), 0)
        self.assertGreater((full.event == "卖出").sum(), 0)
        for _, row in full[full.event != ""].iterrows():
            self.assertEqual(row.direction, 1 if row.event == "买入" else -1)

    def test_old_algorithm_parity(self):
        old_path = Path(r"D:\0-codex\vibe-trading-home\runs\20261026_mrvl_supertrend_local\strategy.py")
        if not old_path.exists():
            self.skipTest("Original script is a local verification fixture")
        tree = ast.parse(old_path.read_text(encoding="utf-8"))
        definitions = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in ("wilder_atr", "add_supertrend")]
        namespace = {"pd": pd, "np": np, "ATR_LENGTH": 10, "MULTIPLIER": 3.0}
        exec(compile(ast.Module(body=definitions, type_ignores=[]), "original", "exec"), namespace)
        data = self.prices()
        old = namespace["add_supertrend"](data.rename(columns=str.lower))
        new = monitor.calculate_supertrend(data)
        np.testing.assert_allclose(old.supertrend, new.supertrend)
        np.testing.assert_equal(old.direction.to_numpy(), new.direction.to_numpy())

    def test_completed_sessions(self):
        self.assertEqual(monitor.session_dates("2026-11-02T15:00:00Z")[0], "2026-10-30")
        self.assertEqual(monitor.session_dates("2026-11-02T21:16:00Z")[0], "2026-11-02")
        # US Thanksgiving Friday closes at 13:00 ET, and Christmas is a holiday.
        self.assertEqual(monitor.session_dates("2025-11-28T18:16:00Z")[0], "2025-11-28")
        self.assertEqual(monitor.session_dates("2025-12-25T22:00:00Z")[0], "2025-12-24")

    def test_stale_live_and_failure(self):
        data = self.prices()
        expected = data.index[-2].strftime("%Y-%m-%d")
        _, schedule = monitor.session_dates("2026-11-02T15:00:00Z")
        with patch.object(monitor.market, "fetch_ohlcv", return_value=(data, "fixture", "")):
            row, _ = monitor.analyze("TEST", expected, schedule)
            self.assertEqual(row["data_date"], expected)  # excludes future/live row
            self.assertEqual(row["health"], "正常")
        with patch.object(monitor, "session_dates", return_value=("2026-10-30", schedule)), \
             patch.object(monitor.market, "fetch_yfinance", return_value=(pd.DataFrame(), "offline")), \
             patch.object(monitor.market, "fetch_ohlcv", return_value=(data, "fixture", "")):
            snapshot = monitor.build_snapshot(["TEST"])
            self.assertEqual(snapshot["stale_count"], 1)
            self.assertEqual(snapshot["triggers"], [])
        with patch.object(monitor.market, "fetch_ohlcv", return_value=(pd.DataFrame(), "", "offline")):
            snapshot = monitor.build_snapshot(["TEST"])
            self.assertIsNone(snapshot["records"][0]["state"])
            self.assertEqual(snapshot["failure_count"], 1)
            json.dumps(snapshot, allow_nan=False)

    def test_repeated_run_no_duplicate_and_state(self):
        expected, schedule = monitor.session_dates("2026-11-02T15:00:00Z")
        row = {"ticker": "TEST", "data_date": expected, "event": "买入", "state": "持有", "health": "正常"}
        event = {"ticker": "TEST", "data_date": expected, "signal": "买入"}
        with patch.object(monitor, "session_dates", return_value=(expected, schedule)), \
             patch.object(monitor, "analyze", return_value=(row, [event])):
            first = monitor.build_snapshot(["TEST", "TEST"])
            second = monitor.build_snapshot(["TEST"])
            self.assertEqual(first["triggers"], second["triggers"])
            self.assertEqual(len(first["triggers"]), 1)
            self.assertEqual(first["records"][0]["state"], "持有")


if __name__ == "__main__":
    unittest.main()
