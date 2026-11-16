"""Selected-pool integration checks without network or production data writes."""
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import numpy as np
import pandas as pd

import import_pick_pool as importer
import pick_monitors as pick
import super_trend_monitor as st
import smc_monitor as smc


class PickTests(unittest.TestCase):
    def prices(self, count=300):
        rng = np.random.default_rng(42)
        close = 100 + np.cumsum(rng.normal(0, 4, count))
        return pd.DataFrame({'Open':close, 'High':close+2, 'Low':close-2, 'Close':close,
                             'Adjusted Close':close, 'Volume':1000},
                            index=pd.bdate_range('2024-01-01', periods=count))

    def pool(self, tickers=('AAA', 'BBB')):
        return {'source_file':'fixture.xlsx', 'stock_pool_count':len(tickers),
                'stocks':[{'ticker':t, 'company_name':t, 'display_ticker':t,
                           'sector':'信息技术', 'industry':'测试', 'indexes':['标普500']} for t in tickers]}

    def context(self, frame):
        expected = frame.index[-1].date().isoformat()
        schedule = pd.DataFrame(index=pd.bdate_range(frame.index[-1], periods=5))
        return expected, schedule

    def test_one_download_per_stock_matches_both_original_rules(self):
        frame = self.prices()
        original = frame.copy(deep=True)
        expected, schedule = self.context(frame)
        with patch.object(st, 'session_dates', return_value=(expected,schedule)), \
                patch.object(st, 'fetch_daily', return_value=(frame,'fixture','')) as fetch:
            snapshots = pick.build_snapshots(self.pool(), max_workers=1)
            self.assertEqual(fetch.call_count, 2)
        with patch.object(st, 'fetch_daily', return_value=(frame,'fixture','')):
            reference, st_events = st.analyze('AAA', expected, schedule)
        with patch.object(smc, 'fetch_daily', return_value=(frame,'fixture','')):
            reference_smc, smc_events = smc.analyze('AAA', expected)
        for key,value in reference.items():
            self.assertEqual(snapshots['st-pick']['records'][0][key], value)
        for key,value in reference_smc.items():
            self.assertEqual(snapshots['smc-pick']['records'][0][key], value)
        self.assertEqual(len(snapshots['st-pick']['history']), len(st_events)*2)
        self.assertEqual(len(snapshots['smc-pick']['history']), len(smc_events)*2)
        self.assertEqual(snapshots['st-pick']['atr_length'], 10)
        self.assertEqual(snapshots['st-pick']['multiplier'], 3)
        self.assertEqual(snapshots['smc-pick']['swing_length'], 50)
        pd.testing.assert_frame_equal(frame, original)

    def test_stale_and_failed_stocks_do_not_trigger(self):
        frame = self.prices()
        expected,schedule = self.context(frame)
        def fetch(ticker):
            return (frame.iloc[:-1] if ticker=='OLD' else pd.DataFrame(), 'fixture', 'offline')
        with patch.object(st, 'session_dates', return_value=(expected,schedule)), \
                patch.object(st, 'fetch_daily', side_effect=fetch):
            snapshots = pick.build_snapshots(self.pool(('OLD','BAD')), max_workers=1)
        for payload in snapshots.values():
            self.assertEqual(payload['triggers'], [])
            self.assertEqual(payload['success_count'], 0)
            self.assertEqual(payload['records'][0]['health'], '行情过期')
            self.assertEqual(payload['records'][1]['health'], '更新失败')
        self.assertIsNone(snapshots['st-pick']['records'][1]['state'])

    def test_short_history_is_independent_between_modules(self):
        frame = self.prices(30)
        expected,schedule = self.context(frame)
        with patch.object(st,'session_dates',return_value=(expected,schedule)), \
                patch.object(st,'fetch_daily',return_value=(frame,'fixture','')):
            snapshots = pick.build_snapshots(self.pool(('AAA',)),max_workers=1)
        self.assertEqual(snapshots['st-pick']['success_count'],1)
        self.assertEqual(snapshots['smc-pick']['records'][0]['health'],'日线不足')
        self.assertEqual(snapshots['smc-pick']['triggers'],[])

    def test_cache_reuses_same_invocation_but_retries_next_run(self):
        frame=self.prices()
        with patch.object(st,'session_dates',return_value=self.context(frame)), \
                patch.object(st,'fetch_daily',return_value=(frame,'fixture','')):
            snapshots=pick.build_snapshots(self.pool(),max_workers=1)
        with tempfile.TemporaryDirectory() as folder, patch.object(pick,'DATA_DIR',Path(folder)), \
                patch.object(pick,'build_snapshots',side_effect=lambda pool:json.loads(json.dumps(snapshots))) as build:
            with patch.dict(os.environ,{'PICK_RUN_ID':'first'}):
                pick.update_pool(self.pool());pick.update_pool(self.pool())
            self.assertEqual(build.call_count,1)
            with patch.dict(os.environ,{'PICK_RUN_ID':'second'}):
                pick.update_pool(self.pool())
                changed=self.pool();changed['source_file']='changed.xlsx'
                pick.update_pool(changed)
            self.assertEqual(build.call_count,3)
            for module in pick.MODULE_LABELS:
                saved=json.loads((Path(folder)/module/'signals.json').read_text(encoding='utf-8'))
                self.assertEqual(saved['pick_run_id'],'second')

    def test_import_merges_sheets_and_preserves_share_classes(self):
        class Sheet:
            def __init__(self,title,codes):self.title,self.codes=title,codes
            def iter_rows(self,values_only=True):
                yield ('股票代码','公司名称（英文）','一级板块','细分行业')
                for ticker in self.codes:yield (ticker,ticker,'金融','测试')
        class Book(list):
            def close(self):pass
        book=Book([Sheet('标普500',['AAPL','BRK.B','GOOG']),Sheet('纳指100',['AAPL','GOOGL','BF.B'])])
        with tempfile.TemporaryDirectory() as folder:
            source=Path(folder)/'source.xlsx';source.write_bytes(b'test source fingerprint')
            with patch.object(importer,'load_workbook',return_value=book):
                result=importer.import_pool(source,Path(folder)/'pool.json')
        self.assertEqual(result['stock_pool_count'],5)
        stocks={r['ticker']:r for r in result['stocks']}
        self.assertEqual(stocks['AAPL']['indexes'],['标普500','纳指100'])
        self.assertEqual(stocks['BRK-B']['display_ticker'],'BRK.B')
        self.assertEqual(stocks['BF-B']['display_ticker'],'BF.B')
        self.assertIn('GOOG',stocks);self.assertIn('GOOGL',stocks)

    def test_committed_pool_is_deduplicated_and_old_pools_unchanged(self):
        pool=pick.load_pool()
        self.assertEqual(len(pool['stocks']),len({r['ticker'] for r in pool['stocks']}))
        self.assertNotEqual({r['ticker'] for r in pool['stocks']},set(st.STOCK_POOL))
        self.assertEqual(len(st.STOCK_POOL),len(set(st.STOCK_POOL)))


if __name__=='__main__':
    unittest.main()
