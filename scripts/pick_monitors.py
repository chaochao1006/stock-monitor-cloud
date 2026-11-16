"""Calculate ST-pick and SMC-pick from one download per selected security."""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import re
from zoneinfo import ZoneInfo

if __package__:
    from . import super_trend_monitor as st_monitor, smc_monitor
else:
    import super_trend_monitor as st_monitor
    import smc_monitor

POOL_PATH = Path(__file__).with_name('pick_stock_pool.json')
DATA_DIR = Path(os.environ.get('REPORT_BASE_DIR', Path(__file__).resolve().parents[1] / 'data'))
MODULE_LABELS = {'st-pick': 'ST-pick', 'smc-pick': 'SMC-pick'}


def load_pool(path=POOL_PATH):
    pool = json.loads(Path(path).read_text(encoding='utf-8'))
    stocks = pool.get('stocks', [])
    tickers = [r['ticker'] for r in stocks]
    if not stocks or len(tickers) != len(set(tickers)):
        raise ValueError('pick股票池为空或含重复代码')
    if any(not re.fullmatch(r'[A-Z][A-Z0-9]*(?:-[A-Z0-9]+)?', t) for t in tickers):
        raise ValueError('pick股票池含非法Yahoo代码')
    if pool.get('stock_pool_count') != len(stocks):
        raise ValueError('pick股票池数量不一致')
    return pool


def build_snapshots(pool, now=None, max_workers=4):
    expected, schedule = st_monitor.session_dates(now)
    generated = datetime.now(ZoneInfo('Asia/Shanghai')).isoformat(timespec='seconds')
    stocks = pool['stocks']

    def worker(stock):
        ticker = stock['ticker']
        details = {k: stock[k] for k in ('company_name', 'display_ticker', 'sector', 'industry', 'indexes') if k in stock}
        try:
            frame, source, error = st_monitor.fetch_daily(ticker)
        except Exception as exc:
            frame, source, error = None, '', str(exc)
        results = []
        for module in ('st-pick', 'smc-pick'):
            try:
                if frame is None:
                    raise ValueError(error or '行情为空')
                if module == 'st-pick':
                    row, events = st_monitor.analyze_prices(ticker, expected, schedule, frame, source, error)
                else:
                    row, events = smc_monitor.analyze_prices(ticker, expected, frame, source, error)
                row.update(details)
                results.append((row, [{**event, **details} for event in events]))
            except Exception as exc:
                results.append(({'ticker': ticker, **details, 'state': None,
                                 'event': '', 'health': '更新失败', 'error': str(exc)}, []))
        return results

    st_records, smc_records, st_history, smc_history = [], [], [], []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        for i, results in enumerate(executor.map(worker, stocks), 1):
            (st_row, st_events), (smc_row, smc_events) = results
            st_records.append(st_row); st_history.extend(st_events)
            smc_records.append(smc_row); smc_history.extend(smc_events)
            if i % 50 == 0 or i == len(stocks):
                print(f'pick: 已计算 {i}/{len(stocks)}', flush=True)
    common = {'data_interval': '1d', 'latest_date': expected, 'generated_at_cn': generated,
              'stock_pool_count': len(stocks), 'pool_source': pool['source_file']}
    snapshots = {}
    for module, records, history in [('st-pick', st_records, st_history), ('smc-pick', smc_records, smc_history)]:
        fresh = sum(r['health'] == '正常' for r in records)
        failures = [{'ticker': r['ticker'], 'error': r.get('error', r['health'])}
                    for r in records if r['health'] == '更新失败']
        snapshots[module] = {**common, 'module': module, 'records': records, 'history': history,
                             'success_count': fresh, 'failure_count': len(failures),
                             'stale_count': sum(r['health'] == '行情过期' for r in records), 'failures': failures}
    snapshots['st-pick'].update({
        'schema_version': st_monitor.SCHEMA_VERSION, 'atr_length': st_monitor.LENGTH,
        'multiplier': st_monitor.MULTIPLIER, 'price_basis': 'Yahoo adjusted OHLC; fixed start 2016-01-01 or listing',
        'triggers': [dict(row, signal=row['event']) for row in st_records
                     if row['health'] == '正常' and row.get('event') in ('买入', '卖出')],
    })
    snapshots['smc-pick'].update({
        'swing_length': smc_monitor.SWING_LENGTH,
        'triggers': [event for event in smc_history if event['data_date'] == expected],
    })
    return snapshots


def write_json(path, payload):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix('.tmp')
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False), encoding='utf-8')
    temporary.replace(path)


def update_pool(pool=None):
    pool = load_pool() if pool is None else pool
    run_id = os.environ.get('PICK_RUN_ID', '')
    fingerprint = hashlib.sha256(json.dumps(pool, sort_keys=True).encode()).hexdigest()
    manifest_path = DATA_DIR / 'pick' / 'last_batch.json'
    # Only reuse a batch inside the same orchestrator invocation. A later run,
    # even on the same date, must download again and retry any failed securities.
    if run_id:
        try:
            saved = json.loads(manifest_path.read_text(encoding='utf-8'))
            if saved.get('run_id') == run_id and saved.get('pool_hash') == fingerprint:
                snapshots = {m: json.loads((DATA_DIR / m / 'signals.json').read_text(encoding='utf-8')) for m in MODULE_LABELS}
                if all(p.get('pick_run_id') == run_id and p.get('module') == m for m, p in snapshots.items()):
                    print('pick: 复用本次运行已计算的共享行情结果', flush=True)
                    return snapshots
        except (OSError, ValueError, TypeError):
            pass
    snapshots = build_snapshots(pool)
    for module, payload in snapshots.items():
        payload['pick_run_id'] = run_id
        write_json(DATA_DIR / module / 'signals.json', payload)
    if run_id:
        write_json(manifest_path, {'run_id': run_id, 'pool_hash': fingerprint})
    return snapshots


def main(module=None):
    snapshots = update_pool()
    chosen = [module] if module else list(MODULE_LABELS)
    failed = False
    for key in chosen:
        payload = snapshots[key]
        print(f"{MODULE_LABELS[key]}: {payload['success_count']}/{payload['stock_pool_count']} 正常，"
              f"{payload['failure_count']}失败，{len(payload['triggers'])}个当日触发", flush=True)
        failed |= payload['success_count'] == 0
    return int(failed)


if __name__ == '__main__':
    raise SystemExit(main())
