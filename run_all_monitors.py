from __future__ import annotations

import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
DATA_DIR = PROJECT_DIR / "data"
SCRIPTS_DIR = PROJECT_DIR / "scripts"

MONITORS = [
    ("BOLL", SCRIPTS_DIR / "us_stock_boll_monitor.py"),
    ("CROSS", SCRIPTS_DIR / "cross_monitor.py"),
    ("短线风险", SCRIPTS_DIR / "short-risk.py"),
    ("中长期风险", SCRIPTS_DIR / "drop_monitor.py"),
    ("BOLL中长期下跌", SCRIPTS_DIR / "BOLL_MACD_RSI_Downtrend_monitor.py"),
]


def run_monitor(label: str, script_path: Path) -> tuple[str, int]:
    """运行单个监控脚本，并返回脚本退出码。"""
    env = os.environ.copy()
    env["REPORT_BASE_DIR"] = str(DATA_DIR)
    env.setdefault("PYTHONIOENCODING", "utf-8")
    print(f"\n========== {datetime.now():%Y-%m-%d %H:%M:%S} 开始运行：{label} ==========")
    if not script_path.exists():
        print(f"脚本不存在：{script_path}")
        return label, 127

    result = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=str(PROJECT_DIR),
        env=env,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60 * 45,
    )
    print(f"========== {label} 运行结束，退出码：{result.returncode} ==========")
    return label, result.returncode


def main() -> None:
    """顺序运行所有监控模块。任何单个模块失败，都不阻断其他模块。"""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    results = []
    for label, script_path in MONITORS:
        try:
            results.append(run_monitor(label, script_path))
        except subprocess.TimeoutExpired:
            print(f"{label} 运行超时，已跳过。")
            results.append((label, 124))
        except Exception as exc:
            print(f"{label} 运行异常：{exc}")
            results.append((label, 1))

    print("\n========== 全部监控完成 ==========")
    for label, code in results:
        state = "成功" if code == 0 else f"失败/异常({code})"
        print(f"{label}: {state}")


if __name__ == "__main__":
    main()
