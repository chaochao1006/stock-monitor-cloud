# 股票技术监控云端版

这个目录是云端部署版本，包含五个监控模块：

- BOLL：`scripts/us_stock_boll_monitor.py`
- CROSS：`scripts/cross_monitor.py`
- 短线风险：`scripts/short-risk.py`
- 中长期风险：`scripts/drop_monitor.py`
- BOLL中长期下跌：`scripts/BOLL_MACD_RSI_Downtrend_monitor.py`

## 运行逻辑

GitHub Actions 每周一至周五北京时间 09:40 左右运行：

```bash
python run_all_monitors.py
```

运行结果保存到：

```text
data/
```

Streamlit Cloud 读取 `data/` 并展示网页仪表盘。

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
