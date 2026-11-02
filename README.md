# 股票技术监控云端版

这个目录是云端部署版本，包含以下监控模块：

- BOLL：`scripts/us_stock_boll_monitor.py`
- CROSS：`scripts/cross_monitor.py`
- 短线风险：`scripts/short-risk.py`
- 中长期风险：`scripts/drop_monitor.py`
- BOLL分位排名：`scripts/boll_percentile_rank.py`
- super-trend：`scripts/super_trend_monitor.py`，日线 Wilder ATR(10) × 3，与 BOLL 分位排名共用股票池。

## super-trend

网页侧栏选择 `super-trend`，依次展示最新已收盘交易日的买入/卖出触发，以及全部股票的持有/观望状态。
买入为空转多，卖出为多转空；持续多头为持有，持续空头为观望。信号计划在下一交易日开盘执行。
失败显示空状态，行情滞后单独标注，不把旧信号当作当日触发。历史触发保留在折叠表中。
使用NYSE实际收盘时间（含节假日及提前收盘），收盘15分钟后才接受当日日K。
Yahoo复权OHLC从2016年或上市后开始计算，与旧AkShare数据口径可能不同。
输出：`data/super-trend/signals.json`；支持 `REPORT_BASE_DIR`。
单独更新：`python scripts/super_trend_monitor.py`。全量更新：`python run_all_monitors.py`。
本地改动需提交推送至部署仓库后，GitHub Actions和Streamlit Cloud才会生效。

## 运行逻辑

GitHub Actions 每周二至周六北京时间 07:17 左右运行，覆盖美股周一至周五收盘：

```bash
python run_all_monitors.py
```

运行结果保存到：

```text
data/
```

Streamlit Cloud 读取 `data/` 并展示网页仪表盘。

每次运行结束后还会生成：

```text
data/last_run_status.json
```

网页首页会显示最近云端运行时间、6 个模块是否成功、每个模块的记录数量。这样即使当天没有新触发，也能确认系统是否正常运行过。

## 自动清理规则

为了避免 GitHub 仓库长期变大，定时任务每天会运行：

```bash
python cleanup_old_outputs.py
```

清理规则：

- Word 报告：只保留最近 30 天
- 日志 `.log`：只保留最近 3 天
- Excel 跟踪文件：保留
- TXT 触发记录：保留

## 本地测试

```bash
pip install -r requirements.txt
python run_all_monitors.py
streamlit run app.py
```

## 部署步骤

1. 新建一个 GitHub 仓库。
2. 把本目录 `cloud` 里的所有文件推送到仓库根目录。
3. 打开 GitHub 仓库的 `Actions`，确认 `Daily Stock Monitors` 工作流已启用。
4. 进入 Streamlit Community Cloud，选择这个仓库。
5. Streamlit 入口文件填：

```text
app.py
```

6. 部署后，网页会读取 GitHub Actions 生成并提交的 `data/` 结果。

说明：GitHub Actions 的定时任务可能不是精确到秒启动，通常会有几分钟延迟。
