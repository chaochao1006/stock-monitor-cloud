"""Daily Supertrend ATR(10) x 3; signals use completed US sessions only."""
from __future__ import annotations

import json
import os
import sys
from functools import lru_cache
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import pandas_market_calendars as mcal
import requests

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
from super_trend_stock_pool import STOCK_POOL

OUTPUT = Path(os.environ.get("REPORT_BASE_DIR", Path(__file__).resolve().parents[1] / "data")) / "super-trend"
LENGTH, MULTIPLIER = 10, 3.0
SCHEMA_VERSION = 2


def validate_daily(frame):
    """Reject weekly/monthly/quarterly series, including a final daily quote."""
    if frame.empty:
        raise ValueError("行情为空")
    if not frame.index.is_unique or not frame.index.is_monotonic_increasing:
        raise ValueError("行情日期重复或未排序")
    gaps = frame.index.to_series().diff().dt.total_seconds().dropna() / 86400
    if not gaps.empty and (gaps.median() > 4 or (gaps <= 4).mean() < 0.8):
        raise ValueError("数据周期异常：日期间隔不符合日线，拒绝计算")


def decode_daily(payload):
    chart = payload.get("chart", {})
    if chart.get("error") or not chart.get("result"):
        raise ValueError(str(chart.get("error") or "Yahoo返回空数据"))
    block = chart["result"][0]
    granularity = block.get("meta", {}).get("dataGranularity")
    if granularity != "1d":
        raise ValueError(f"数据周期异常：请求1d，实际返回{granularity}")
    timestamps = block.get("timestamp") or []
    quote = (block.get("indicators", {}).get("quote") or [{}])[0]
    adjusted = (block.get("indicators", {}).get("adjclose") or [{}])[0].get("adjclose")
    if not timestamps or adjusted is None:
        raise ValueError("缺少时间戳或复权收盘价")
    frame = pd.DataFrame(quote, index=pd.to_datetime(timestamps, unit="s", utc=True)
                         .tz_convert("America/New_York").tz_localize(None).normalize())
    frame = frame.rename(columns=str.capitalize)
    frame["Adjusted Close"] = adjusted
    frame = frame[["Open", "High", "Low", "Close", "Adjusted Close", "Volume"]].apply(pd.to_numeric, errors="coerce")
    frame = frame.dropna(subset=["Open", "High", "Low", "Close", "Adjusted Close"]).sort_index()
    validate_daily(frame)
    return frame


@lru_cache(maxsize=8)
def download_session_context(minute):
    return session_dates(minute)


def complete_latest_session(frame, payload, expected, schedule):
    """Recover one missing daily bar only from a fully closed, verified regular session."""
    day = pd.Timestamp(expected)
    previous = schedule.index[schedule.index < day]
    if not len(previous) or frame.empty or frame.index[-1] != previous[-1]:
        raise ValueError('备用仅补齐紧邻上一交易日的一根日K，不跨越缺失历史')
    block = payload['chart']['result'][0]
    meta = block['meta']
    if meta.get('dataGranularity') not in ('60m', '1h') or meta.get('exchangeTimezoneName') != 'America/New_York':
        raise ValueError('备用行情周期或时区不符合美股常规交易时段')
    opening, closing = schedule.loc[day, ['market_open', 'market_close']]
    quoted_at = pd.to_datetime(meta['regularMarketTime'], unit='s', utc=True)
    if quoted_at.tz_convert('America/New_York').date().isoformat() != expected or quoted_at < closing - pd.Timedelta(seconds=60):
        raise ValueError('备用报价未达到目标交易日收盘')
    if pd.Timestamp.now(tz='UTC') < closing + pd.Timedelta(minutes=15):
        raise ValueError('目标交易日尚未完成收盘校验等待时间')
    bars = pd.DataFrame(block['indicators']['quote'][0],
                        index=pd.to_datetime(block['timestamp'], unit='s', utc=True))
    bars = bars.loc[(bars.index >= opening) & (bars.index <= closing)]
    required = pd.date_range(opening, closing, freq='60min', inclusive='left')
    if not bars.index.is_unique or not required.isin(bars.index).all() or bars.index[-1] != closing:
        raise ValueError('备用缺少完整常规交易时段或最终收盘记录')
    prices = bars[['open', 'high', 'low', 'close']].apply(pd.to_numeric, errors='coerce')
    if not np.isfinite(prices.to_numpy()).all() or (prices <= 0).any().any():
        raise ValueError('备用时段包含无效OHLC')
    if (prices.high < prices.max(axis=1)).any() or (prices.low > prices.min(axis=1)).any():
        raise ValueError('备用时段OHLC关系无效')
    close = float(meta['regularMarketPrice'])
    high, low = float(meta['regularMarketDayHigh']), float(meta['regularMarketDayLow'])
    volume = float(meta['regularMarketVolume'])
    if not np.isfinite([close, high, low, volume]).all() or min(close, high, low, volume) <= 0:
        raise ValueError('缺少有效的常规交易日收盘报价或成交量')
    # Intraday bars can omit auction/consolidated-tape extremes. Use the final
    # daily quote's high/low, allowing at most 25bp discrepancy for verification;
    # the final closing print must still agree at quote precision.
    if (not np.allclose([prices.high.max(), prices.low.min()], [high, low], rtol=.0025, atol=1e-4)
            or not np.isclose(prices.close.iloc[-1], close, rtol=1e-5, atol=1e-4)):
        raise ValueError('备用时段OHLC与最终常规交易日报价不一致')
    open_price = float(prices.open.iloc[0])
    if high < max(open_price, close, low) or low > min(open_price, close, high):
        raise ValueError('备用日线OHLC关系无效')
    # The quoted session must be the most recent actual market session. Its adjusted
    # close equals its raw close; all previous Yahoo adjusted bars remain untouched.
    result = frame.copy()
    result.loc[day] = [open_price, high, low, close, close, volume]
    validate_daily(result)
    return result


def fetch_daily(symbol):
    # Explicit timestamps prevent Yahoo's range=max automatic quarterly aggregation.
    params = {"period1": int(pd.Timestamp("2016-01-01", tz="America/New_York").timestamp()),
              "period2": int(pd.Timestamp.now(tz="UTC").timestamp()),
              "interval": "1d", "events": "history", "includeAdjustedClose": "true"}
    errors, candidates = [], []
    expected, schedule = download_session_context(pd.Timestamp.now(tz='UTC').floor('min').isoformat())
    for host in ("query1.finance.yahoo.com", "query2.finance.yahoo.com"):
        try:
            response = requests.get(f"https://{host}/v8/finance/chart/{symbol}", params=params,
                                    headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
            response.raise_for_status()
            frame = decode_daily(response.json())
            source = f"Yahoo Chart API {host} (verified 1d)"
            if not frame.loc[:expected].empty and frame.loc[:expected].index[-1].date().isoformat() == expected:
                return frame, source, ''
            candidates.append((frame, source))
        except Exception as exc:
            errors.append(f"{host}: {exc}")
    if candidates:
        frame, source = max(candidates, key=lambda pair: pair[0].index[-1])
        for host in ('query1.finance.yahoo.com', 'query2.finance.yahoo.com'):
            try:
                response = requests.get(f'https://{host}/v8/finance/chart/{symbol}',
                                        params={'range':'5d', 'interval':'60m', 'includePrePost':'false'},
                                        headers={'User-Agent':'Mozilla/5.0'}, timeout=20)
                response.raise_for_status()
                repaired = complete_latest_session(frame, response.json(), expected, schedule)
                return repaired, source + '; latest 1d OHLCV verified from complete regular session and closing quote', ''
            except Exception as exc:
                errors.append(f'{host} completed-session fallback: {exc}')
        # Preserve its real date. Never relabel an old bar as today's signal.
        return frame, source, '; '.join(errors)
    return pd.DataFrame(), "", "; ".join(errors)


def session_dates(now=None):
    now = pd.Timestamp(now if now is not None else datetime.now(ZoneInfo("UTC")))
    if now.tzinfo is None:
        raise ValueError("now must be timezone aware")
    schedule = mcal.get_calendar("NYSE").schedule(
        start_date=(now - timedelta(days=20)).date(), end_date=(now + timedelta(days=15)).date())
    closed = schedule[schedule.market_close + pd.Timedelta(minutes=15) <= now]
    latest = closed.index[-1].strftime("%Y-%m-%d")
    return latest, schedule


def calculate_supertrend(frame, length=LENGTH, multiplier=MULTIPLIER):
    """Same SMA-seeded Wilder ATR and band recursion as the original daily script."""
    df = frame.copy().sort_index()
    prev = df.Close.shift(1)
    tr = pd.concat([df.High - df.Low, (df.High - prev).abs(), (df.Low - prev).abs()], axis=1).max(axis=1)
    atr = np.full(len(df), np.nan)
    if len(df) < length:
        raise ValueError("数据不足：至少需要10根已收盘日K")
    atr[length - 1] = tr.iloc[:length].mean()
    for i in range(length, len(df)):
        atr[i] = (atr[i - 1] * (length - 1) + tr.iloc[i]) / length
    basic_upper = (df.High.to_numpy() + df.Low.to_numpy()) / 2 + multiplier * atr
    basic_lower = (df.High.to_numpy() + df.Low.to_numpy()) / 2 - multiplier * atr
    upper, lower = basic_upper.copy(), basic_lower.copy()
    direction, line = np.full(len(df), np.nan), np.full(len(df), np.nan)
    first = length - 1
    direction[first], line[first] = -1, upper[first]
    close = df.Close.to_numpy()
    for i in range(length, len(df)):
        upper[i] = basic_upper[i] if basic_upper[i] < upper[i-1] or close[i-1] > upper[i-1] else upper[i-1]
        lower[i] = basic_lower[i] if basic_lower[i] > lower[i-1] or close[i-1] < lower[i-1] else lower[i-1]
        if line[i-1] == upper[i-1]:
            direction[i] = 1 if close[i] > upper[i] else -1
        else:
            direction[i] = -1 if close[i] < lower[i] else 1
        line[i] = lower[i] if direction[i] == 1 else upper[i]
    df["atr"], df["direction"], df["supertrend"] = atr, direction, line
    df["event"] = np.where((df.direction == 1) & (df.direction.shift(1) == -1), "买入",
                           np.where((df.direction == -1) & (df.direction.shift(1) == 1), "卖出", ""))
    return df.dropna(subset=["supertrend"])


def analyze(symbol, expected, schedule):
    df, source, error = fetch_daily(symbol)
    return analyze_prices(symbol, expected, schedule, df, source, error)


def analyze_prices(symbol, expected, schedule, df, source, error=""):
    """Same monitor rule, accepting already fetched prices for shared scans."""
    if df.empty:
        raise ValueError(error or "行情为空")
    # The fetcher normalizes dates to New York dates. Exclude live and future bars.
    df = df.loc["2016-01-01":expected].copy()
    if df.empty:
        raise ValueError("没有已收盘行情")
    validate_daily(df)
    daily_bar_count = len(df)
    factor = df["Adjusted Close"] / df.Close
    if factor.isna().any() or not np.isfinite(factor).all() or (factor <= 0).any():
        raise ValueError("缺少可靠复权因子，无法计算")
    for col in ["Open", "High", "Low", "Close"]:
        df[col] = df[col] * factor
    if not np.isfinite(df[["Open", "High", "Low", "Close"]]).all().all() or (df[["Open", "High", "Low", "Close"]] <= 0).any().any():
        raise ValueError("行情含无效价格")
    df = calculate_supertrend(df)
    if len(df) < 2:
        raise ValueError("数据不足：无法确认趋势翻转")
    latest = df.iloc[-1]
    data_date = df.index[-1].strftime("%Y-%m-%d")
    events = [{"ticker": symbol, "data_date": idx.strftime("%Y-%m-%d"), "signal": row.event,
               "close": round(float(row.Close), 4), "supertrend": round(float(row.supertrend), 4)}
              for idx, row in df[df.event != ""].iterrows()]
    future = schedule.index[schedule.index > pd.Timestamp(data_date)]
    return {
        "ticker": symbol, "data_date": data_date,
        "state": "持有" if latest.direction == 1 else "观望",
        "event": latest.event, "close": round(float(latest.Close), 4),
        "supertrend": round(float(latest.supertrend), 4),
        "last_event_date": events[-1]["data_date"] if events else "",
        "last_event": events[-1]["signal"] if events else "",
        "execution_date": future[0].strftime("%Y-%m-%d") if len(future) else "",
        "data_source": source, "health": "正常" if data_date == expected else "行情过期",
        "data_interval": "1d", "daily_bar_count": daily_bar_count,
        "history_start": df.index[0].strftime("%Y-%m-%d"),
    }, events


def build_snapshot(symbols=None):
    symbols = list(dict.fromkeys(symbols if symbols is not None else STOCK_POOL))
    expected, schedule = session_dates()
    records, events, failures = [], [], []
    def worker(symbol):
        try:
            row, history = analyze(symbol, expected, schedule)
            return row, history, None
        except Exception as exc:
            return {"ticker": symbol, "state": None, "health": "更新失败"}, [], {"ticker": symbol, "error": str(exc)}
    with ThreadPoolExecutor(max_workers=4) as executor:
        for row, history, failure in executor.map(worker, symbols):
            records.append(row)
            events.extend(history)
            if failure:
                failures.append(failure)
    today = [dict(row, signal=row["event"]) for row in records
             if row.get("health") == "正常" and row.get("event") in ("买入", "卖出")]
    return {"module": "super-trend", "schema_version": SCHEMA_VERSION, "data_interval": "1d", "latest_date": expected,
            "generated_at_cn": datetime.now(ZoneInfo("Asia/Shanghai")).isoformat(timespec="seconds"),
            "atr_length": LENGTH, "multiplier": MULTIPLIER, "stock_pool_count": len(symbols),
            "price_basis": "Yahoo adjusted OHLC; fixed calculation start 2016-01-01 or listing",
            "success_count": sum(r["health"] == "正常" for r in records),
            "failure_count": len(failures), "stale_count": sum(r["health"] == "行情过期" for r in records),
            "records": records, "triggers": today, "history": events, "failures": failures}


def snapshot_exit_code(payload, label):
    records = payload.get('records', [])
    total = payload.get('stock_pool_count', len(records))
    fresh = payload.get('success_count', sum(row.get('health') == '正常' for row in records))
    if not total or fresh / total < .9:
        print(f'{label}: 新鲜行情覆盖不足 {fresh}/{total}，要求至少90%；真实状态已保存。', flush=True)
        return 1
    return 0


def main():
    payload = build_snapshot()
    OUTPUT.mkdir(parents=True, exist_ok=True)
    path = OUTPUT / "signals.json"
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False), encoding="utf-8")
    temporary.replace(path)
    print(f"super-trend: {payload['success_count']}/{payload['stock_pool_count']} 正常，"
          f"{len(payload['triggers'])} 个最新交易日触发；输出 {path}", flush=True)
    return snapshot_exit_code(payload, 'super-trend')


if __name__ == "__main__":
    raise SystemExit(main())
