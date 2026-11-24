import ast
import copy
import datetime as dt
from pathlib import Path
import unittest
from unittest.mock import Mock, patch

import numpy as np
import pandas as pd

import super_trend_monitor as st


class DailyRecoveryTests(unittest.TestCase):
    def fixture(self, early_close=False):
        day = '2026-11-20'
        opening = pd.Timestamp(day+'T14:30:00Z')
        closing = pd.Timestamp(day+('T18:00:00Z' if early_close else 'T21:00:00Z'))
        dates = pd.date_range(opening, closing, freq='60min', inclusive='left').append(pd.DatetimeIndex([closing]))
        bars = {'open':[100.]*len(dates), 'high':[102.]*len(dates), 'low':[99.]*len(dates),
                'close':[101.]*len(dates), 'volume':[1000.]*len(dates)}
        payload = {'chart':{'result':[{
            'meta':{'dataGranularity':'60m', 'exchangeTimezoneName':'America/New_York',
                    'regularMarketTime':int(closing.timestamp())+1, 'regularMarketPrice':101.,
                    'regularMarketDayHigh':102., 'regularMarketDayLow':99., 'regularMarketVolume':9000.},
            'timestamp':[int(t.timestamp()) for t in dates], 'indicators':{'quote':[bars]}}]}}
        schedule = pd.DataFrame({'market_open':[opening-pd.Timedelta(days=1), opening],
                                 'market_close':[closing-pd.Timedelta(days=1), closing]},
                                index=pd.to_datetime(['2026-11-19', day]))
        frame = pd.DataFrame({'Open':100., 'High':102., 'Low':99., 'Close':101.,
                              'Adjusted Close':95., 'Volume':1000.},
                             index=pd.bdate_range(end='2026-11-19', periods=20))
        return frame, payload, day, schedule

    def test_full_regular_session_only_preserves_adjusted_history(self):
        for early in (False, True):
            frame, payload, day, schedule = self.fixture(early)
            repaired = st.complete_latest_session(frame, payload, day, schedule)
            pd.testing.assert_frame_equal(repaired.loc[frame.index], frame, check_freq=False)
            self.assertEqual(repaired.index[-1], pd.Timestamp(day))
            self.assertEqual(repaired.loc[day, 'Adjusted Close'], 101.)
            self.assertEqual(repaired.loc[day, 'Volume'], 9000.)
            self.assertEqual(len(repaired), len(frame)+1)

    def test_missing_hours_wrong_quote_wrong_day_and_gap_are_rejected(self):
        frame, original, day, schedule = self.fixture()
        broken = copy.deepcopy(original)
        block = broken['chart']['result'][0]
        block['timestamp'].pop(1)
        for values in block['indicators']['quote'][0].values():
            values.pop(1)
        with self.assertRaisesRegex(ValueError, '完整常规交易时段'):
            st.complete_latest_session(frame, broken, day, schedule)
        broken = copy.deepcopy(original)
        broken['chart']['result'][0]['meta']['regularMarketDayHigh'] = 105.
        with self.assertRaisesRegex(ValueError, '不一致'):
            st.complete_latest_session(frame, broken, day, schedule)
        broken = copy.deepcopy(original)
        broken['chart']['result'][0]['meta']['regularMarketTime'] -= 86400
        with self.assertRaisesRegex(ValueError, '未达到'):
            st.complete_latest_session(frame, broken, day, schedule)
        with self.assertRaisesRegex(ValueError, '不跨越缺失历史'):
            st.complete_latest_session(frame.iloc[:-1], original, day, schedule)
        broken = copy.deepcopy(original)
        broken['chart']['result'][0]['meta']['dataGranularity'] = '1d'
        with self.assertRaisesRegex(ValueError, '周期或时区'):
            st.complete_latest_session(frame, broken, day, schedule)

    def test_small_tape_extreme_difference_uses_final_daily_quote(self):
        frame, payload, day, schedule = self.fixture()
        payload['chart']['result'][0]['meta']['regularMarketDayHigh'] = 102.01
        payload['chart']['result'][0]['meta']['regularMarketDayLow'] = 98.9
        repaired = st.complete_latest_session(frame, payload, day, schedule)
        self.assertEqual(repaired.loc[day, 'High'], 102.01)
        self.assertEqual(repaired.loc[day, 'Low'], 98.9)

    def test_stale_first_host_tries_second_and_then_completed_session(self):
        frame, payload, day, schedule = self.fixture()
        stale_response = Mock()
        fallback_response = Mock()
        fallback_response.json.return_value = payload
        with patch.object(st, 'download_session_context', return_value=(day,schedule)), \
             patch.object(st, 'decode_daily', return_value=frame), \
             patch.object(st.requests, 'get', side_effect=[stale_response,stale_response,fallback_response]) as get:
            result, source, error = st.fetch_daily('NVDA')
        self.assertEqual(str(result.index[-1].date()), day)
        self.assertIn('complete regular session', source)
        self.assertFalse(error)
        self.assertIn('query2', get.call_args_list[1].args[0])
        self.assertEqual(get.call_args_list[2].kwargs['params']['interval'], '60m')
        # A failed verification must leave the old dates intact.
        with patch.object(st, 'download_session_context', return_value=(day,schedule)), \
             patch.object(st, 'decode_daily', return_value=frame), \
             patch.object(st.requests, 'get', side_effect=[stale_response,stale_response,ValueError('offline'),ValueError('offline')]):
            result, _, error = st.fetch_daily('NVDA')
        pd.testing.assert_frame_equal(result, frame)
        self.assertIn('offline', error)

    def test_partial_outage_is_not_success(self):
        self.assertEqual(st.snapshot_exit_code({'stock_pool_count':88,'success_count':1}, 'fixture'), 1)
        self.assertEqual(st.snapshot_exit_code({'stock_pool_count':88,'success_count':87}, 'fixture'), 0)
        self.assertEqual(st.snapshot_exit_code({'records':[]}, 'fixture'), 1)

    def test_short_risk_retries_stale_yfinance_instead_of_returning_it(self):
        source = Path(__file__).with_name('short-risk.py').read_text(encoding='utf-8-sig')
        node = next(n for n in ast.parse(source).body if isinstance(n, ast.FunctionDef) and n.name == 'fetch_stock_data')
        stale, _, day, _ = self.fixture()
        fresh = stale.copy()
        fresh.loc[pd.Timestamp(day)] = fresh.iloc[-1]
        backup = Mock(return_value=(fresh, ''))
        yahoo = Mock()
        yahoo.download.return_value = stale
        namespace = {'pd':pd,'dt':dt,'Tuple':tuple,'yf':yahoo,'HISTORY_PERIOD':'18mo','INTERVAL':'1d',
                     'normalize_ohlcv':lambda frame, target:(frame,''), 'fetch_yahoo_chart':backup}
        exec(compile(ast.Module(body=[node],type_ignores=[]), 'short-risk', 'exec'), namespace)
        result, error = namespace['fetch_stock_data']('NVDA', pd.Timestamp(day).date())
        self.assertEqual(result.index[-1], pd.Timestamp(day))
        self.assertFalse(error)
        backup.assert_called_once()
        backup.reset_mock()
        yahoo.download.return_value = fresh
        namespace['fetch_stock_data']('NVDA', pd.Timestamp(day).date())
        backup.assert_not_called()


if __name__ == '__main__':
    unittest.main()
