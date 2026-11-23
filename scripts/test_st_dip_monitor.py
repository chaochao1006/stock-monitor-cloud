import copy
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import numpy as np
import pandas as pd
from streamlit.testing.v1 import AppTest

import st_dip_monitor as dip
import super_trend_monitor as st


class DipTests(unittest.TestCase):
    def test_same_supertrend_rules_sell_only_and_exact_pool(self):
        expected_pool = set('AAOI AMD AMZN APLD AXTI BE BMNR CBRS COHR COIN CRCL CRWV IONQ IREN LITE META MSTR MU NBIS NVDA ORCL PLTR QBTS RGTI SKHY SMCI SNDK SPCX TSLA'.split())
        self.assertEqual(set(dip.STOCK_POOL), expected_pool)
        self.assertEqual(len(dip.STOCK_POOL), 29)
        close = np.r_[np.full(10, 100.), 110., 90., 120., 70.]
        data = pd.DataFrame({'Open':close, 'High':close+1, 'Low':close-1, 'Close':close,
                             'Adjusted Close':close, 'Volume':1000},
                            index=pd.bdate_range('2026-10-01', periods=len(close)))
        day = str(data.index[-1].date())
        schedule = pd.DataFrame(index=pd.bdate_range(data.index[0], periods=40))
        with patch.object(st, 'fetch_daily', return_value=(data, 'fixture', '')), \
             patch.object(st, 'session_dates', return_value=(day, schedule)):
            original = st.build_snapshot(dip.STOCK_POOL)
            actual = dip.build_snapshot()
        self.assertTrue(any(r['signal'] == '买入' for r in original['history']))
        self.assertEqual(actual['triggers'], [r for r in original['triggers'] if r['signal'] == '卖出'])
        self.assertEqual(actual['history'], [r for r in original['history'] if r['signal'] == '卖出'])
        self.assertEqual(len(actual['triggers']), 29)
        self.assertEqual(actual['atr_length'], 10)
        self.assertEqual(actual['multiplier'], 3.)
        for row, old in zip(actual['records'], original['records']):
            for field in ('close','supertrend','state','execution_date','daily_bar_count'):
                self.assertEqual(row[field], old[field])
        # A live bar cannot flip an otherwise completed signal; stale stocks do not trigger.
        with patch.object(st, 'fetch_daily', return_value=(data, 'fixture', '')), \
             patch.object(st, 'session_dates', return_value=(str(schedule.index[len(data)].date()), schedule)):
            stale = dip.build_snapshot()
        self.assertEqual(stale['triggers'], [])
        self.assertEqual(stale['stale_count'], 29)
        with tempfile.TemporaryDirectory() as folder, patch.object(dip, 'OUTPUT', Path(folder)), \
             patch.object(dip, 'build_snapshot', return_value=actual):
            self.assertEqual(dip.main(), 0)
            self.assertEqual(json.loads((Path(folder)/'signals.json').read_text(encoding='utf-8'))['module'], 'st-dip')

    def test_dashboard_sell_filter_and_orchestrator_status(self):
        root = Path(__file__).resolve().parents[1]
        row = {'ticker':'NVDA','data_date':'2026-11-20','close':100.,'supertrend':110.,
               'health':'正常','signal':'卖出','execution_date':'2026-11-23'}
        payload = {'module':'st-dip','schema_version':2,'data_interval':'1d','signal_filter':'sell_only',
                   'latest_date':'2026-11-20','generated_at_cn':'2026-11-23T12:00:00+08:00',
                   'stock_pool_count':29,'success_count':29,'failure_count':0,'stale_count':0,
                   'records':[row], 'triggers':[row,dict(row,ticker='AMD',signal='买入'),
                                              dict(row,ticker='BE',health='行情过期')],
                   'history':[row,dict(row,ticker='AMD',signal='买入')]}
        with tempfile.TemporaryDirectory() as folder, patch.dict(os.environ, {'REPORT_BASE_DIR':folder}):
            path = Path(folder)/'st-dip'
            path.mkdir()
            (path/'signals.json').write_text(json.dumps(payload), encoding='utf-8')
            (Path(folder)/'last_run_status.json').write_text(json.dumps({'模块':[{'模块':'ST-dip','退出码':0}]}), encoding='utf-8')
            app = AppTest.from_file(str(root/'app.py')).run(timeout=20)
            app.sidebar.radio[0].set_value('ST-dip').run(timeout=20)
            self.assertFalse(app.exception)
            self.assertEqual(app.subheader[0].value, '当日卖出触发记录')
            self.assertEqual(app.dataframe[0].value['触发信号'].tolist(), ['卖出'])
            self.assertEqual(app.dataframe[-1].value['触发信号'].tolist(), ['卖出'])
            self.assertFalse(any('持有（' in item.value for item in app.subheader))
            self.assertIn('全部成功', app.sidebar.caption[1].value)
        import sys
        sys.path.insert(0, str(root))
        import run_all_monitors as runner
        self.assertIn('ST-dip', [label for label, _, _ in runner.MONITORS])


if __name__ == '__main__':
    unittest.main()
