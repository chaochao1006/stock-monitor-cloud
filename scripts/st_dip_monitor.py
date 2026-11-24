"""Daily Supertrend sell flips for a separate 29-stock watchlist."""
from pathlib import Path
import os

if __package__:
    from . import super_trend_monitor as st
    from .pick_monitors import write_json
else:
    import super_trend_monitor as st
    from pick_monitors import write_json


STOCK_POOL = [
    'AAOI', 'AMD', 'AMZN', 'APLD', 'AXTI', 'BE', 'BMNR', 'CBRS', 'COHR', 'COIN',
    'CRCL', 'CRWV', 'IONQ', 'IREN', 'LITE', 'META', 'MSTR', 'MU', 'NBIS', 'NVDA',
    'ORCL', 'PLTR', 'QBTS', 'RGTI', 'SKHY', 'SMCI', 'SNDK', 'SPCX', 'TSLA',
]
OUTPUT = Path(os.environ.get('REPORT_BASE_DIR', Path(__file__).resolve().parents[1] / 'data')) / 'st-dip'


def build_snapshot():
    payload = st.build_snapshot(STOCK_POOL)
    payload['module'] = 'st-dip'
    payload['signal_filter'] = 'sell_only'
    payload['triggers'] = [row for row in payload['triggers']
                           if row.get('signal') == '卖出' and row.get('health') == '正常'
                           and row.get('data_date') == payload['latest_date']]
    payload['history'] = [row for row in payload['history'] if row.get('signal') == '卖出']
    latest_sell = {}
    for event in payload['history']:
        ticker = event['ticker']
        if event['data_date'] > latest_sell.get(ticker, ''):
            latest_sell[ticker] = event['data_date']
    for row in payload['records']:
        row['event'] = '卖出' if row.get('event') == '卖出' and row.get('health') == '正常' else ''
        row['last_event_date'] = latest_sell.get(row['ticker'], '')
        row['last_event'] = '卖出' if row['last_event_date'] else ''
    return payload


def main():
    payload = build_snapshot()
    write_json(OUTPUT / 'signals.json', payload)
    print(f"ST-dip: {payload['success_count']}/{payload['stock_pool_count']} 正常，"
          f"{len(payload['triggers'])} 个最新交易日卖出触发", flush=True)
    return st.snapshot_exit_code(payload, 'ST-dip')


if __name__ == '__main__':
    raise SystemExit(main())
