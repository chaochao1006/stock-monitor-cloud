"""Exercise both selected pages and five-module status in Streamlit itself."""
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import numpy as np
import pandas as pd
from streamlit.testing.v1 import AppTest

import pick_monitors as pick
import super_trend_monitor as st


class PickDashboardTests(unittest.TestCase):
    def test_selected_pages_and_status_are_wired(self):
        close=100+np.cumsum(np.random.default_rng(42).normal(0,4,300))
        frame=pd.DataFrame({'Open':close,'High':close+2,'Low':close-2,'Close':close,
                            'Adjusted Close':close,'Volume':1000},index=pd.bdate_range('2024-01-01',periods=300))
        expected=frame.index[-1].date().isoformat()
        schedule=pd.DataFrame(index=pd.bdate_range(frame.index[-1],periods=5))
        pool={'source_file':'fixture.xlsx','stock_pool_count':1,
              'stocks':[{'ticker':'AAPL','company_name':'Apple','industry':'测试'}]}
        with patch.object(st,'session_dates',return_value=(expected,schedule)), \
                patch.object(st,'fetch_daily',return_value=(frame,'fixture','')):
            snapshots=pick.build_snapshots(pool,max_workers=1)
        with tempfile.TemporaryDirectory() as folder, patch.dict(os.environ,{'REPORT_BASE_DIR':folder}):
            root=Path(folder)
            for module,payload in snapshots.items():
                pick.write_json(root/module/'signals.json',payload)
            labels=['super-trend','SMC','ST-pick','SMC-pick','短线风险']
            pick.write_json(root/'last_run_status.json',{'状态':'全部成功','模块':[
                {'模块':label,'退出码':0,'状态':'成功'} for label in labels]})
            app=AppTest.from_file(str(Path(__file__).resolve().parents[1]/'app.py')).run(timeout=30)
            self.assertEqual(len(app.exception),0)
            self.assertIn('5/5',[metric.value for metric in app.metric])
            for label in ['ST-pick','SMC-pick']:
                app.sidebar.radio[0].set_value(label).run(timeout=30)
                self.assertEqual(len(app.exception),0)
                self.assertIn(label,[header.value for header in app.header])
                self.assertGreaterEqual(len(app.dataframe),1)


if __name__=='__main__':
    unittest.main()
