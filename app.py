from __future__ import annotations

import json
import re
import os
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
import streamlit as st

try:
    from docx import Document
except Exception:  # pragma: no cover - optional local dependency
    Document = None


BASE_DIR = Path(os.environ.get("REPORT_BASE_DIR", Path(__file__).resolve().parent / "data"))
RUN_STATUS_PATH = BASE_DIR / "last_run_status.json"
SUPER_TREND_PATH = BASE_DIR / "super-trend" / "signals.json"


@dataclass(frozen=True)
class ModuleConfig:
    key: str
    label: str
    folder: Path
    trigger_file: Path | None = None
    report_patterns: tuple[str, ...] = ()


MODULES = [
    ModuleConfig(
        key="short_risk",
        label="短线风险",
        folder=BASE_DIR / "drop" / "short-risk",
        trigger_file=BASE_DIR / "drop" / "short-risk" / "RISK.txt",
        report_patterns=("美股短线反弹风险监控_*.docx",),
    ),
]


def parse_date(value) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return pd.to_datetime(str(value)).date()
    except Exception:
        return None


def clean_symbol(raw: str) -> tuple[str, int | None]:
    text = str(raw).strip()
    match = re.match(r"^([A-Za-z0-9.\-]+)[（(](\d+)[）)]$", text)
    if match:
        return match.group(1).upper(), int(match.group(2))
    return text.upper(), None


def read_text_lines(path: Path) -> list[str]:
    if not path or not path.exists():
        return []
    for encoding in ("utf-8-sig", "utf-8", "gbk"):
        try:
            return [line.strip() for line in path.read_text(encoding=encoding).splitlines() if line.strip()]
        except UnicodeDecodeError:
            continue
    return []


@st.cache_data(ttl=60)
def load_run_status() -> dict:
    if not RUN_STATUS_PATH.exists():
        return {}
    try:
        status = json.loads(RUN_STATUS_PATH.read_text(encoding="utf-8"))
        active_labels = {"super-trend", *(module.label for module in MODULES)}
        modules = [item for item in status.get("模块", []) if item.get("模块") in active_labels]
        succeeded = sum(item.get("退出码") == 0 for item in modules)
        if modules:
            status["状态"] = "全部成功" if succeeded == len(modules) else "部分失败"
        status.update({"模块": modules, "模块总数": len(modules), "成功模块数": succeeded,
                       "失败模块数": len(modules) - succeeded})
        return status
    except Exception as exc:
        return {"状态": "状态文件读取失败", "错误": str(exc)}


@st.cache_data(ttl=60)
def load_trigger_records() -> pd.DataFrame:
    rows: list[dict] = []
    for module in MODULES:
        if not module.trigger_file:
            continue
        for line in read_text_lines(module.trigger_file):
            parts = re.split(r"\t+|\s{2,}", line)
            if len(parts) < 2:
                parts = line.split()
            if len(parts) < 2:
                continue
            trigger_date = parse_date(parts[0])
            symbol, count = clean_symbol(parts[1])
            detail = " ".join(parts[2:]).strip() if len(parts) > 2 else ""
            rows.append(
                {
                    "模块": module.label,
                    "日期": trigger_date,
                    "股票": symbol,
                    "触发次数": count,
                    "信号/评分": detail,
                    "原始记录": line,
                    "来源文件": str(module.trigger_file),
                }
            )
    if not rows:
        return pd.DataFrame(columns=["模块", "日期", "股票", "触发次数", "信号/评分", "原始记录", "来源文件"])
    df = pd.DataFrame(rows)
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    return df.sort_values(["日期", "模块", "股票"], ascending=[False, True, True]).reset_index(drop=True)


@st.cache_data(ttl=60)
def load_reports() -> pd.DataFrame:
    rows = []
    for module in MODULES:
        for pattern in module.report_patterns:
            for path in sorted(module.folder.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True):
                report_date = None
                match = re.search(r"(\d{4}-\d{2}-\d{2})", path.name)
                if match:
                    report_date = parse_date(match.group(1))
                rows.append(
                    {
                        "模块": module.label,
                        "报告日期": report_date,
                        "文件名": path.name,
                        "修改时间": datetime.fromtimestamp(path.stat().st_mtime),
                        "大小KB": round(path.stat().st_size / 1024, 1),
                        "路径": str(path),
                    }
                )
    if not rows:
        return pd.DataFrame(columns=["模块", "报告日期", "文件名", "修改时间", "大小KB", "路径"])
    df = pd.DataFrame(rows)
    df["报告日期"] = pd.to_datetime(df["报告日期"], errors="coerce")
    return df.sort_values(["报告日期", "修改时间"], ascending=[False, False]).reset_index(drop=True)


def read_docx_preview(path: str, limit: int = 18) -> str:
    if Document is None:
        return "当前环境无法读取 Word 正文预览。"
    try:
        doc = Document(path)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        return "\n\n".join(paragraphs[:limit]) or "报告中没有可预览的段落。"
    except Exception as exc:
        return f"读取报告预览失败：{exc}"


def filter_by_modules(df: pd.DataFrame, modules: Iterable[str]) -> pd.DataFrame:
    if df.empty:
        return df
    selected = list(modules)
    if not selected:
        return df.iloc[0:0]
    return df[df["模块"].isin(selected)].copy()


def render_kpis(triggers: pd.DataFrame, reports: pd.DataFrame) -> None:
    latest_trigger_date = triggers["日期"].max() if not triggers.empty else pd.NaT
    latest_report_date = reports["报告日期"].max() if not reports.empty else pd.NaT
    trigger_col, report_col = st.columns(2)
    trigger_col.metric("触发记录数", f"{len(triggers):,}")
    report_col.metric("最新报告日期", latest_report_date.strftime("%Y-%m-%d") if pd.notna(latest_report_date) else "N/A")
    st.caption(f"最新触发日期：{latest_trigger_date.strftime('%Y-%m-%d') if pd.notna(latest_trigger_date) else 'N/A'}")


def render_run_status(status: dict) -> None:
    st.subheader("云端运行状态")
    if not status:
        st.warning("还没有读取到云端运行状态文件。可能是 GitHub Actions 尚未成功运行，或还没有把 data/last_run_status.json 提交到仓库。")
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("最近运行状态", status.get("状态", "未知"))
    c2.metric("结束时间", status.get("结束时间", "N/A"))
    c3.metric("成功模块", f"{status.get('成功模块数', 0)}/{status.get('模块总数', 0)}")
    c4.metric("耗时", f"{status.get('耗时秒', 0)} 秒")

    modules = status.get("模块", [])
    if modules:
        display = pd.DataFrame(modules)
        wanted = ["模块", "状态", "退出码", "触发记录数", "本次新增触发", "最新触发日期", "开始时间", "结束时间", "耗时秒"]
        cols = [col for col in wanted if col in display.columns]
        st.dataframe(display[cols], use_container_width=True, hide_index=True)
    if status.get("错误"):
        st.error(status["错误"])


def render_trigger_section(triggers: pd.DataFrame) -> None:
    st.subheader("触发记录")
    if triggers.empty:
        st.info("当前没有触发记录。")
        return
    date_range = [triggers["日期"].min().date(), triggers["日期"].max().date()]
    chosen = st.date_input("触发日期范围", value=date_range, min_value=date_range[0], max_value=date_range[1])
    if isinstance(chosen, tuple) and len(chosen) == 2:
        start, end = pd.to_datetime(chosen[0]), pd.to_datetime(chosen[1])
        triggers = triggers[(triggers["日期"] >= start) & (triggers["日期"] <= end)]
    display_columns = ["日期", "模块", "股票", "触发次数", "信号/评分", "原始记录"]
    st.dataframe(triggers[display_columns], use_container_width=True, hide_index=True)


def render_super_trend() -> None:
    st.header("super-trend")
    st.caption("日线 Supertrend · Wilder ATR(10) × 3 · 收盘确认翻转，计划下一交易日开盘执行。持有表示多头趋势，不代表实际账户持仓。")
    if not SUPER_TREND_PATH.exists():
        st.info("尚无监控数据，请运行 scripts/super_trend_monitor.py 或等待云端每日更新。")
        return
    try:
        payload = json.loads(SUPER_TREND_PATH.read_text(encoding="utf-8"))
    except Exception as exc:
        st.error(f"读取 super-trend 失败：{exc}")
        return
    if payload.get("schema_version", 0) < 2 or payload.get("data_interval") != "1d":
        st.error("此快照未通过日线周期校验，已停止展示。请更新监控脚本并重新运行 super-trend。")
        return
    st.caption(f"信号交易日：{payload.get('latest_date', 'N/A')} ｜ 最近计算：{payload.get('generated_at_cn', 'N/A')}")
    st.caption(f"股票池 {payload.get('stock_pool_count', 0)} 只 ｜ 正常 {payload.get('success_count', 0)} ｜ 行情过期 {payload.get('stale_count', 0)} ｜ 更新失败 {payload.get('failure_count', 0)}")
    if payload.get("stale_count") or payload.get("failure_count"):
        st.warning("部分股票行情过期或更新失败，请以当前信号表中的行情日期和数据状态为准。")
    # Calendar freshness is checked locally, without making market-data requests.
    try:
        from scripts.super_trend_monitor import session_dates
        expected, _ = session_dates()
        if payload.get("latest_date", "") < expected:
            st.warning(f"当前整份快照已过期，最新应有交易日为 {expected}。以下为历史快照，请等待更新。")
    except Exception:
        st.caption("未能校验快照时效，请核对信号交易日。")
    st.subheader("当日触发记录")
    st.caption("当日指上述最新已收盘交易日；周末沿用最近交易日。行情过期或获取失败的股票不列入当日触发。")
    rename = {"ticker": "股票", "data_date": "行情日期", "state": "当前信号", "signal": "触发信号",
              "close": "复权收盘价", "supertrend": "ST线", "last_event_date": "最近触发日期",
              "last_event": "最近触发", "execution_date": "计划执行日", "health": "数据状态",
              "data_interval": "数据周期", "daily_bar_count": "日K数量"}
    triggered = pd.DataFrame(payload.get("triggers", []))
    if triggered.empty:
        st.info("该交易日无有效买入／卖出触发。")
    else:
        columns = ["ticker", "signal", "data_date", "close", "supertrend", "execution_date"]
        st.dataframe(triggered.reindex(columns=columns).rename(columns=rename), use_container_width=True, hide_index=True)
    st.subheader("股票池当前信号")
    columns = ["ticker", "state", "data_date", "close", "supertrend", "last_event_date", "last_event", "data_interval", "daily_bar_count", "health"]
    rows = pd.DataFrame(payload.get("records", [])).reindex(columns=columns)
    for state in ("持有", "观望"):
        group = rows[rows["state"] == state]
        st.subheader(f"{state}（{len(group)}只）")
        st.dataframe(group.rename(columns=rename), use_container_width=True, hide_index=True)
        if group.empty:
            st.caption(f"当前没有{state}状态的股票。")
    unknown = rows[~rows["state"].isin(["持有", "观望"])]
    if not unknown.empty:
        with st.expander(f"信号暂不可用（{len(unknown)}只）"):
            st.dataframe(unknown.rename(columns=rename), use_container_width=True, hide_index=True)
    st.caption("价格采用Yahoo复权口径，与原AkShare回测或TradingView设置不同可能产生差异。数据状态异常时勿将其当作当天确认信号。")
    if payload.get("failures"):
        with st.expander("获取失败详情"):
            st.dataframe(pd.DataFrame(payload["failures"]), use_container_width=True, hide_index=True)
    with st.expander("历史买卖触发"):
        history = pd.DataFrame(payload.get("history", []))
        if not history.empty:
            st.dataframe(history.sort_values(["data_date", "ticker"], ascending=[False, True]).rename(columns=rename), use_container_width=True, hide_index=True)


def render_reports_section(reports: pd.DataFrame) -> None:
    st.subheader("Word 报告")
    if reports.empty:
        st.info("当前没有报告文件。")
        return
    st.dataframe(reports[["报告日期", "模块", "文件名", "修改时间", "大小KB"]], use_container_width=True, hide_index=True)

    options = [f"{row['模块']} | {row['文件名']}" for _, row in reports.iterrows()]
    selected = st.selectbox("预览/下载报告", options)
    idx = options.index(selected)
    path = reports.iloc[idx]["路径"]
    with st.expander("报告正文预览", expanded=False):
        st.text(read_docx_preview(path))
    with open(path, "rb") as handle:
        st.download_button("下载这个 Word 报告", handle, file_name=Path(path).name)


def render_single_module_page(module_label: str, triggers: pd.DataFrame, reports: pd.DataFrame) -> None:
    st.header(module_label)
    module_triggers = filter_by_modules(triggers, [module_label])
    module_reports = filter_by_modules(reports, [module_label])
    render_kpis(module_triggers, module_reports)
    render_trigger_section(module_triggers)
    with st.expander("查看该模块 Word 报告", expanded=False):
        render_reports_section(module_reports)


def main() -> None:
    st.set_page_config(page_title="股票监控仪表盘", page_icon="📈", layout="wide")
    st.title("股票技术监控仪表盘")
    st.caption("云端版：读取项目 data 目录中的 Supertrend 信号与短线风险报告。")

    triggers = load_trigger_records()
    reports = load_reports()
    run_status = load_run_status()

    module_labels = [m.label for m in MODULES]
    selected_modules = st.sidebar.multiselect("模块筛选", module_labels, default=module_labels)
    page = st.sidebar.radio(
        "页面",
        ["总览", "super-trend", "短线风险", "报告中心", "数据诊断"],
    )
    st.sidebar.caption(f"数据目录：{BASE_DIR}")
    if run_status:
        st.sidebar.caption(f"最近云端运行：{run_status.get('结束时间', 'N/A')} | {run_status.get('状态', '未知')}")
    if st.sidebar.button("刷新数据"):
        st.cache_data.clear()
        st.rerun()

    filtered_triggers = filter_by_modules(triggers, selected_modules)
    filtered_reports = filter_by_modules(reports, selected_modules)

    if page == "总览":
        render_run_status(run_status)
        render_kpis(filtered_triggers, filtered_reports)
        st.subheader("最新触发")
        if filtered_triggers.empty:
            st.info("暂无触发记录。")
        else:
            st.dataframe(filtered_triggers.head(25)[["日期", "模块", "股票", "信号/评分"]], use_container_width=True, hide_index=True)
        st.subheader("最新报告")
        if filtered_reports.empty:
            st.info("暂无报告。")
        else:
            st.dataframe(filtered_reports.head(12)[["报告日期", "模块", "文件名"]], use_container_width=True, hide_index=True)

    elif page == "super-trend":
        render_super_trend()

    elif page == "短线风险":
        render_single_module_page("短线风险", triggers, reports)

    elif page == "报告中心":
        render_reports_section(filtered_reports)

    elif page == "数据诊断":
        st.subheader("数据源状态")
        st.write(
            {
                "触发记录": len(triggers),
                "Word报告": len(reports),
                "基础目录存在": BASE_DIR.exists(),
                "运行状态文件存在": RUN_STATUS_PATH.exists(),
            }
        )
        render_run_status(run_status)
        st.dataframe(pd.DataFrame([m.__dict__ for m in MODULES]), use_container_width=True)


if __name__ == "__main__":
    main()
