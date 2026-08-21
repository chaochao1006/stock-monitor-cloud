from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo


PROJECT_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("REPORT_BASE_DIR", PROJECT_DIR / "data"))
SCRIPTS_DIR = PROJECT_DIR / "scripts"
STATUS_PATH = DATA_DIR / "last_run_status.json"
BEIJING_TZ = ZoneInfo("Asia/Shanghai")

MONITORS = [
    ("BOLL", SCRIPTS_DIR / "us_stock_boll_monitor.py", DATA_DIR / "BOLL" / "BOLL.txt"),
    ("CROSS", SCRIPTS_DIR / "cross_monitor.py", DATA_DIR / "CROSS" / "CROSS.txt"),
    ("短线风险", SCRIPTS_DIR / "short-risk.py", DATA_DIR / "drop" / "short-risk" / "RISK.txt"),
    ("中长期风险", SCRIPTS_DIR / "drop_monitor.py", DATA_DIR / "drop" / "drop.txt"),
    ("BOLL中长期下跌", SCRIPTS_DIR / "BOLL_MACD_RSI_Downtrend_monitor.py", DATA_DIR / "drop" / "drop-BOLL" / "BOLL-dip.txt"),
]


def now_bj() -> datetime:
    """返回北京时间。"""
    return datetime.now(BEIJING_TZ)


def read_trigger_lines(path: Path) -> list[str]:
    """读取触发记录行。"""
    if not path.exists():
        return []
    for encoding in ("utf-8-sig", "utf-8", "gbk"):
        try:
            return [line.strip() for line in path.read_text(encoding=encoding).splitlines() if line.strip()]
        except UnicodeDecodeError:
            continue
    return []


def latest_trigger_date(lines: list[str]) -> str:
    """从触发记录中提取最新日期。"""
    latest = ""
    for line in lines:
        parts = re.split(r"\t+|\s+", line.strip())
        if not parts:
            continue
        value = parts[0].strip()
        if re.match(r"^\d{4}-\d{2}-\d{2}$", value) and value > latest:
            latest = value
    return latest


def run_monitor(label: str, script_path: Path, trigger_file: Path) -> dict:
    """运行单个监控脚本，并返回运行状态。"""
    env = os.environ.copy()
    env["REPORT_BASE_DIR"] = str(DATA_DIR)
    env.setdefault("PYTHONIOENCODING", "utf-8")
    started_at = now_bj()
    before_lines = read_trigger_lines(trigger_file)
    print(f"\n========== {started_at:%Y-%m-%d %H:%M:%S} 开始运行：{label} ==========")
    if not script_path.exists():
        print(f"脚本不存在：{script_path}")
        return {
            "模块": label,
            "状态": "脚本不存在",
            "退出码": 127,
            "开始时间": started_at.strftime("%Y-%m-%d %H:%M:%S"),
            "结束时间": now_bj().strftime("%Y-%m-%d %H:%M:%S"),
            "触发记录数": len(before_lines),
            "本次新增触发": 0,
            "最新触发日期": latest_trigger_date(before_lines),
        }

    result = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=str(PROJECT_DIR),
        env=env,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60 * 45,
    )
    finished_at = now_bj()
    after_lines = read_trigger_lines(trigger_file)
    print(f"========== {label} 运行结束，退出码：{result.returncode} ==========")
    return {
        "模块": label,
        "状态": "成功" if result.returncode == 0 else "失败",
        "退出码": result.returncode,
        "开始时间": started_at.strftime("%Y-%m-%d %H:%M:%S"),
        "结束时间": finished_at.strftime("%Y-%m-%d %H:%M:%S"),
        "耗时秒": round((finished_at - started_at).total_seconds(), 1),
        "触发记录数": len(after_lines),
        "本次新增触发": max(len(after_lines) - len(before_lines), 0),
        "最新触发日期": latest_trigger_date(after_lines),
    }


def write_status(started_at: datetime, finished_at: datetime, module_results: list[dict]) -> None:
    """写入云端运行状态文件，供 Streamlit 首页展示。"""
    failed = [item for item in module_results if item.get("退出码") != 0]
    payload = {
        "状态": "全部成功" if not failed else "部分失败",
        "开始时间": started_at.strftime("%Y-%m-%d %H:%M:%S"),
        "结束时间": finished_at.strftime("%Y-%m-%d %H:%M:%S"),
        "耗时秒": round((finished_at - started_at).total_seconds(), 1),
        "模块总数": len(module_results),
        "成功模块数": len(module_results) - len(failed),
        "失败模块数": len(failed),
        "模块": module_results,
    }
    STATUS_PATH.parent.mkdir(parents=True, exist_ok=True)
    STATUS_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已写入运行状态：{STATUS_PATH}")


def main() -> None:
    """顺序运行所有监控模块。任何单个模块失败，都不阻断其他模块。"""
    started_at = now_bj()
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    results = []
    for label, script_path, trigger_file in MONITORS:
        try:
            results.append(run_monitor(label, script_path, trigger_file))
        except subprocess.TimeoutExpired:
            print(f"{label} 运行超时，已跳过。")
            lines = read_trigger_lines(trigger_file)
            results.append(
                {
                    "模块": label,
                    "状态": "超时",
                    "退出码": 124,
                    "触发记录数": len(lines),
                    "本次新增触发": 0,
                    "最新触发日期": latest_trigger_date(lines),
                }
            )
        except Exception as exc:
            print(f"{label} 运行异常：{exc}")
            lines = read_trigger_lines(trigger_file)
            results.append(
                {
                    "模块": label,
                    "状态": f"异常：{exc}",
                    "退出码": 1,
                    "触发记录数": len(lines),
                    "本次新增触发": 0,
                    "最新触发日期": latest_trigger_date(lines),
                }
            )

    print("\n========== 全部监控完成 ==========")
    for item in results:
        state = item["状态"] if item["退出码"] == 0 else f"{item['状态']}({item['退出码']})"
        print(f"{item['模块']}: {state}")
    write_status(started_at, now_bj(), results)


if __name__ == "__main__":
    main()
