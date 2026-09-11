from __future__ import annotations

import json
import math
import os
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from us_stock_boll_monitor import (  # noqa: E402
    BOLL_LOW_LOOKBACK_DAYS,
    BOLL_PERIOD,
    BOLL_STD_MULTIPLIER,
    STOCK_POOL,
    add_bollinger,
    fetch_ohlcv,
    get_company_name,
    safe_float,
)


REPORT_BASE_DIR = Path(os.environ.get("REPORT_BASE_DIR", SCRIPT_DIR.parent / "data"))
OUTPUT_DIR = REPORT_BASE_DIR / "BOLL"
RANK_PATH = OUTPUT_DIR / "BOLL_percentile_rank.json"
BEIJING_TZ = ZoneInfo("Asia/Shanghai")


def fmt_date(value) -> str:
    """把 pandas 日期统一输出为 YYYY-MM-DD。"""
    if value is None or pd.isna(value):
        return ""
    return pd.Timestamp(value).strftime("%Y-%m-%d")


def safe_round(value, digits: int = 6):
    """安全四舍五入，避免 NaN/Inf 写入 JSON。"""
    number = safe_float(value)
    if number is None or not math.isfinite(number):
        return None
    return round(number, digits)


def current_bandwidth_percentile(df: pd.DataFrame) -> float | None:
    """计算最新 Bandwidth 在近8个月历史区间中的分位。"""
    bandwidth = df["Bandwidth"].dropna().tail(BOLL_LOW_LOOKBACK_DAYS)
    if len(bandwidth) < 60:
        return None
    latest = bandwidth.iloc[-1]
    return float((bandwidth <= latest).mean())


def analyze_symbol(symbol: str) -> tuple[dict | None, dict | None]:
    """下载单只股票数据并计算 BOLL 带宽排名字段。"""
    company_name = get_company_name(symbol)
    df, source, error = fetch_ohlcv(symbol)
    if df.empty:
        return None, {
            "ticker": symbol,
            "company_name": company_name,
            "error": error or "无法获取行情数据",
        }

    df = add_bollinger(df)
    df = df.dropna(subset=["Middle Band", "Upper Band", "Lower Band", "Bandwidth", "Percent B"])
    if len(df) < 60:
        return None, {
            "ticker": symbol,
            "company_name": company_name,
            "error": "BOLL 有效数据少于60个交易日，无法计算8个月分位",
        }

    latest = df.iloc[-1]
    percentile = current_bandwidth_percentile(df)
    if percentile is None:
        return None, {
            "ticker": symbol,
            "company_name": company_name,
            "data_date": fmt_date(df.index[-1]),
            "error": "近8个月 Bandwidth 有效数据不足",
        }

    bandwidth = safe_float(latest["Bandwidth"])
    previous_20 = df["Bandwidth"].dropna().tail(21)
    if len(previous_20) >= 21 and previous_20.iloc[0]:
        bandwidth_20d_change = bandwidth / previous_20.iloc[0] - 1 if bandwidth is not None else None
    else:
        bandwidth_20d_change = None

    record = {
        "rank": None,
        "ticker": symbol,
        "company_name": company_name,
        "data_date": fmt_date(df.index[-1]),
        "close": safe_round(latest["Close"], 4),
        "upper_band": safe_round(latest["Upper Band"], 4),
        "middle_band": safe_round(latest["Middle Band"], 4),
        "lower_band": safe_round(latest["Lower Band"], 4),
        "bandwidth": safe_round(bandwidth, 6),
        "history_percentile": safe_round(percentile, 6),
        "percent_b": safe_round(latest["Percent B"], 6),
        "bandwidth_20d_change": safe_round(bandwidth_20d_change, 6),
        "data_source": source,
        "status": "OK",
    }
    return record, None


def build_ranking() -> dict:
    """生成全股票池 BOLL 带宽历史分位排名。"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    records: list[dict] = []
    failures: list[dict] = []

    for symbol in STOCK_POOL:
        try:
            record, failure = analyze_symbol(symbol)
            if record:
                records.append(record)
            if failure:
                failures.append(failure)
        except Exception as exc:
            failures.append({"ticker": symbol, "company_name": symbol, "error": str(exc)})

    records.sort(key=lambda item: (
        item["history_percentile"] if item["history_percentile"] is not None else 999,
        item["bandwidth"] if item["bandwidth"] is not None else 999,
        item["ticker"],
    ))
    for index, record in enumerate(records, start=1):
        record["rank"] = index

    data_dates = [item["data_date"] for item in records if item.get("data_date")]
    date_counts = {value: data_dates.count(value) for value in sorted(set(data_dates))}
    ranking_date = ""
    if date_counts:
        ranking_date = max(date_counts.items(), key=lambda item: (item[1], item[0]))[0]
    return {
        "module": "BOLL分位排名",
        "generated_at_cn": datetime.now(BEIJING_TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "ranking_date": ranking_date,
        "latest_data_date": max(data_dates) if data_dates else "",
        "data_date_counts": date_counts,
        "stock_pool_count": len(STOCK_POOL),
        "success_count": len(records),
        "failure_count": len(failures),
        "lookback_days": BOLL_LOW_LOOKBACK_DAYS,
        "lookback_note": "近8个月约168个交易日",
        "boll_period": BOLL_PERIOD,
        "std_multiplier": BOLL_STD_MULTIPLIER,
        "records": records,
        "failures": failures,
    }


def main() -> None:
    """入口函数：写入 JSON 排名文件，并在终端输出摘要。"""
    payload = build_ranking()
    RANK_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"BOLL分位排名已更新：{RANK_PATH}")
    print(
        f"排名日期：{payload.get('ranking_date') or 'N/A'}；"
        f"成功：{payload['success_count']}；失败：{payload['failure_count']}"
    )


if __name__ == "__main__":
    main()
