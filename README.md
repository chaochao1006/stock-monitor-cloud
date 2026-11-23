# 股票技术监控云端版

这个目录是云端部署版本，包含以下监控模块：

- 短线风险：`scripts/short-risk.py`，与super-trend、SMC共用股票池。
- SMC：`scripts/smc_monitor.py`，与super-trend共用同一份合并去重后的股票池。
- super-trend：`scripts/super_trend_monitor.py`，日线 Wilder ATR(10) × 3，与SMC共用股票池，统一配置在 `scripts/super_trend_stock_pool.py`。
- ST-pick：`scripts/st_pick_monitor.py`，沿用super-trend规则，使用选股表导入的独立股票池。
- SMC-pick：`scripts/smc_pick_monitor.py`，沿用SMC绿色Swing BOS/CHoCH规则，使用同一选股池。

原有super-trend、SMC与短线风险已统一为84只证券；以后只需修改`scripts/super_trend_stock_pool.py`，三个模块会同时采用新名单。ST-pick与SMC-pick的独立选股池不受影响。

## super-trend

网页侧栏选择 `super-trend`，依次展示最新已收盘交易日的买入/卖出触发，以及全部股票的持有/观望状态。
买入为空转多，卖出为多转空；持续多头为持有，持续空头为观望。信号计划在下一交易日开盘执行。
失败显示空状态，行情滞后单独标注，不把旧信号当作当日触发。历史触发保留在折叠表中。
使用NYSE实际收盘时间（含节假日及提前收盘），收盘15分钟后才接受当日日K。
Yahoo复权OHLC从2016年或上市后开始计算，与旧AkShare数据口径可能不同。
日线请求使用明确起止时间，不使用可能被Yahoo自动降为季度线的 `range=max`。
逐票校验返回元数据为 `1d` 并检查日期间隔；页面显示数据周期和日K数量。
旧版未校验快照会被页面拒绝展示，必须用修复后的脚本重算。
输出：`data/super-trend/signals.json`；支持 `REPORT_BASE_DIR`。
单独更新：`python scripts/super_trend_monitor.py`。全量更新：`python run_all_monitors.py`。
本地改动需提交推送至部署仓库后，GitHub Actions和Streamlit Cloud才会生效。

## SMC

使用 Yahoo 复权日线，沿用本地 SMC 回测的 Swing 长度 50。已收盘的日线确认绿色 Swing BOS 或 Swing CHoCH 后，显示在当日触发记录。页面也展示历史触发和各股票的数据状态。行情过期、日线不足或取数失败的股票不参与当日触发。

单独更新可运行 `python scripts/smc_monitor.py`，结果写入 `data/smc/signals.json`。定时任务会一起更新 SMC。

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

网页首页会显示最近云端运行时间、5个模块是否成功、每个模块的记录数量。这样即使当天没有新触发，也能确认系统是否正常运行过。

## ST-pick和SMC-pick

网页侧栏分别选择`ST-pick`和`SMC-pick`。ST-pick展示当日买卖触发、持有表和观望表；SMC-pick展示当日绿色Swing BOS/CHoCH、历史触发和股票数据状态。规则直接调用原模块的计算函数，不采用此前回测里的固定止盈止损规则。

股票池于2026年11月18日重新按桌面最新选票表的两个工作表导入：标普页177只、纳指页67只，原始合并去重192只。按要求剔除ACN、EBAY、F、VZ、WDAY，并于11月19日剔除BAX、CI、GDDY、INTU、ABT、FSLR、TMUS、ADBE后，两个pick模块共用179只证券。GOOG和GOOGL是不同证券，均保留；BRK.B、BF.B自动转成Yahoo代码BRK-B、BF-B。

云端读取已提交的`scripts/pick_stock_pool.json`，不依赖桌面的Excel路径。以后修改Excel后，可重新导入再提交这个JSON：

```powershell
python scripts/import_pick_pool.py "C:\Users\81975\Desktop\每日报告\选票\标普500与纳指100_细分行业市值排序_20261116.xlsx"
```

两个pick模块在同一次全量运行中只下载一次当前选股池的行情，分别计算两套信号。后续运行会重新下载，不会因为同一天运行过就沿用失败结果。输出分别保存到`data/st-pick/signals.json`和`data/smc-pick/signals.json`。

单独更新两个pick模块：`python scripts/pick_monitors.py`。全量定时任务会自动运行，无需在云端上传Excel。修改代码与股票池后，需要提交推送到部署仓库才会生效。

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
