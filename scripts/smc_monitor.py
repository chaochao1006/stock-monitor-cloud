"""Daily bullish Swing events from the existing LuxAlgo SMC Python port.

Derived from LuxAlgo Smart Money Concepts, CC BY-NC-SA 4.0.
https://www.luxalgo.com/library/indicator/smart-money-concepts-smc/
"""
from __future__ import annotations

import json
import os
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np

if __package__:
    from .super_trend_monitor import fetch_daily, session_dates, validate_daily
else:
    from super_trend_monitor import fetch_daily, session_dates, validate_daily

STOCK_POOL = """
SMR XE NNE OKLO BE IONQ QBTS INSP BWAY MANE CRSP MRNA TEM QSI ABCL IBRX TWST
RKLB ASTS LUNR FLY ERAS VST CORZ VRT IREN AAOI SPCX ORCL CRCL NOK CBRS POET
GLW AMAT NVTS MRVL SHAZ CRWV NBIS COHR FLNC MXL AXTI APP AVGO ALAB ARM INTC FORM LITE
""".split()
SWING_LENGTH = 50
OUTPUT = Path(os.environ.get("REPORT_BASE_DIR", Path(__file__).resolve().parents[1] / "data")) / "smc"


def swing_events(frame, length=SWING_LENGTH):
    """Yield bullish events on the confirmation bar and track bearish breaks for bias."""
    high, low, close = (frame[col].to_numpy(dtype=float) for col in ("High", "Low", "Close"))
    leg = bias = 0
    upper = lower = np.nan
    upper_crossed = lower_crossed = False
    for i in range(length, len(frame)):
        previous_upper, previous_lower = upper, lower
        previous_leg = leg
        origin = i - length
        if high[origin] > high[origin + 1:i + 1].max():
            leg = 0
        elif low[origin] < low[origin + 1:i + 1].min():
            leg = 1
        if leg != previous_leg:
            if leg == 1:
                lower, lower_crossed = low[origin], False
            else:
                upper, upper_crossed = high[origin], False

        if not upper_crossed and close[i] > upper and close[i - 1] <= previous_upper:
            yield i, "Swing CHoCH" if bias == -1 else "Swing BOS", float(upper)
            upper_crossed = True
            bias = 1
        if not lower_crossed and close[i] < lower and close[i - 1] >= previous_lower:
            lower_crossed = True
            bias = -1


def analyze(symbol, expected):
    frame, source, error = fetch_daily(symbol)
    if frame.empty:
        raise ValueError(error or "行情为空")
    frame = frame.loc["2016-01-01":expected]
    validate_daily(frame)
    prices = frame[["Open", "High", "Low", "Close"]]
    adjusted = prices.mul(frame["Adjusted Close"] / frame["Close"], axis=0)
    if not np.isfinite(adjusted.to_numpy()).all() or (adjusted <= 0).any().any():
        raise ValueError("复权行情包含无效价格")
    if ((adjusted.High < adjusted[["Open", "Close", "Low"]].max(axis=1)) |
        (adjusted.Low > adjusted[["Open", "Close", "High"]].min(axis=1))).any():
        raise ValueError("复权行情的开高低收不一致")
    data_date = frame.index[-1].date().isoformat()
    if data_date != expected:
        return {"ticker": symbol, "data_date": data_date, "health": "行情过期"}, []
    if len(frame) <= SWING_LENGTH:
        return {"ticker": symbol, "data_date": data_date, "health": "日线不足"}, []
    events = [
        {"ticker": symbol, "data_date": frame.index[i].date().isoformat(),
         "signal": tag, "close": round(float(adjusted.Close.iloc[i]), 4), "level": round(level, 4)}
        for i, tag, level in swing_events(adjusted)
    ]
    last = events[-1] if events else {}
    return {"ticker": symbol, "data_date": data_date, "health": "正常",
            "data_source": source, "last_event": last.get("signal"),
            "last_event_date": last.get("data_date")}, events


def build_snapshot(symbols=STOCK_POOL):
    expected, _ = session_dates()

    def worker(symbol):
        try:
            return analyze(symbol, expected)
        except Exception as exc:
            return {"ticker": symbol, "health": "更新失败", "error": str(exc)}, []

    with ThreadPoolExecutor(max_workers=4) as executor:
        results = list(executor.map(worker, symbols))
    records = [row for row, _ in results]
    history = [event for _, events in results for event in events]
    return {"module": "smc", "data_interval": "1d", "swing_length": SWING_LENGTH,
            "latest_date": expected,
            "generated_at_cn": datetime.now(ZoneInfo("Asia/Shanghai")).isoformat(timespec="seconds"),
            "records": records, "history": history,
            "triggers": [event for event in history if event["data_date"] == expected]}


def main():
    payload = build_snapshot()
    OUTPUT.mkdir(parents=True, exist_ok=True)
    path = OUTPUT / "signals.json"
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False), encoding="utf-8")
    temporary.replace(path)
    fresh = sum(row["health"] == "正常" for row in payload["records"])
    print(f"SMC: {fresh}/{len(payload['records'])} 正常，{len(payload['triggers'])} 个当日触发")
    return 0 if fresh else 1


if __name__ == "__main__":
    raise SystemExit(main())
