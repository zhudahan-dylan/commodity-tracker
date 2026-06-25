# 大宗商品价格日报工具 — 技术设计文档

## 1. 项目概述

每日自动从 [生意社 (100ppi.com)](https://www.100ppi.com) 采集大宗商品价格数据，存入 Excel 文件。Excel 按品类分 Sheet，每 Sheet 含折线图 + 原始数据，汇总页提供品类级一览视图。

---

## 2. 数据源分析

### 2.1 目标页面

| 数据层级 | URL | 内容 | 每品类商品数 |
|----------|-----|------|-------------|
| 首页「商品涨跌榜」 | `https://www.100ppi.com` | 各品类精选商品价格 + 七日涨跌幅 | ~2 个 |
| 品类动态页 (后续扩展) | `/news/list-{id}-1.html` | 品类下更多报价动态 | 10~30 个 |

**Phase 1**：只采首页涨跌榜。架构预留品类详情页采集接口，后续按需开启。

### 2.2 首页数据形态

首页 HTML 中「商品涨跌榜」区域的文本结构（经 Playwright 渲染后）：

```
商品涨跌榜
2026-06-22

商品名称         价格          七日涨跌幅

能源
WTI原油          76.79         -12.45%
Brent            79.55         -11.98%

化工
丙烯             7384.33       -16.41%
甲醇             2916.67       -15.38%
...
```

### 2.3 解析策略

1. Playwright 打开首页 → 等待 `networkidle`
2. 获取 `document.body.innerText` 全文
3. 定位「商品涨跌榜」锚点 → 截取后 4000 字符
4. 正则按品类名分段 → 每段内按 3 行一组（名称 / 价格 / 涨跌幅）提取
5. Fallback：文本解析失败时尝试 DOM 选择器遍历

---

## 3. 技术选型

| 组件 | 选型 | 原因 |
|------|------|------|
| 浏览器自动化 | Playwright (Python, `async`) | 页面有 JS 动态渲染，纯 HTTP 请求拿不到完整数据 |
| Excel 读写 | openpyxl | 支持 .xlsx 格式化、折线图、自动筛选、冻结表头 |
| 调度 | macOS `launchd` (Plist) | Mac 原生，比 cron 更适合 GUI 会话环境 |
| 语言 | Python 3.9+ | 生态成熟，openpyxl + Playwright 均为一线库 |

---

## 4. 项目结构

```
commodityTracker/
├── main.py                 # CLI 入口
├── config.py               # 全局配置（URL、路径、超时等）
├── scraper.py              # Playwright 采集模块
├── storage.py              # Excel 读写、去重、格式化
├── chart_builder.py        # 折线图构建（透视 + 按单位拆图）
├── requirements.txt        # playwright, openpyxl
├── demo_preview.py         # [临时] 效果预览脚本，交付前删除
└── data/
    └── commodity_prices.xlsx   # 输出的 Excel 文件（运行时生成）
```

---

## 5. 模块设计

### 5.1 `config.py` — 全局配置

```python
BASE_URL = "https://www.100ppi.com"
OUTPUT_DIR = "./data"
EXCEL_FILE = "commodity_prices.xlsx"
BROWSER_TIMEOUT = 30_000      # ms
PAGE_LOAD_TIMEOUT = 60_000    # ms

# Excel 表头
SUMMARY_HEADERS = ["品类", "代表商品", "单位", "最新价格", "单日涨跌", "7日走势", "趋势"]
DATA_HEADERS   = ["日期", "商品名称", "价格", "单位", "七日涨跌幅(%)", "记录时间"]
```

不再硬编码品类列表——品类名从页面动态提取。

### 5.2 `scraper.py` — 采集模块

#### 核心函数

| 函数 | 输入 | 输出 | 说明 |
|------|------|------|------|
| `scrape_homepage()` | — | `list[dict]` | 异步，打开首页 → 提取涨跌榜数据 |
| `scrape_category_page(cat_cn: str)` | 品类中文名 | `list[dict]` | Phase 2 扩展，抓品类详情页 |
| `_parse_ranking_text(text, date, time)` | 页面纯文本 | `list[dict]` | 纯文本正则解析 |
| `_extract_from_dom(page, date, time)` | Playwright Page | `list[dict]` | DOM 选择器 Fallback |

#### 每条记录结构

```python
{
    "日期": "2026-06-22",          # str, 从页面提取或当日日期
    "品类": "能源",                # str, 从页面动态提取
    "商品名称": "WTI原油",         # str
    "价格": 76.79,                # float
    "单位": "元/吨",              # str, 默认 "元/吨"（页面未标注时）
    "七日涨跌幅(%)": -12.45,      # float | None
    "记录时间": "2026-06-22 09:30:00",  # str, 采集时刻
}
```

#### 正则解析伪代码

```
1. text = page.innerText
2. idx = text.find("商品涨跌榜")
3. section = text[idx : idx+4000]
4. date = regex(r'\d{4}-\d{2}-\d{2}', section[:200]) or today
5. 按品类名 (能源|化工|橡塑|纺织|有色|钢铁|建材|农副) 分段
6. 每段内连续 3 行一组:
    行1 → 商品名称
    行2 → 价格 (去逗号 → float)
    行3 → 涨跌幅 (正则 [+-]?\d+\.?\d*%)
7. 品类名本身出现在数据行时跳过（非商品名）
```

#### 错误处理

- 网络超时 → 打印错误，返回空列表
- 文本解析失败 → 自动 fallback 到 DOM 选择器
- DOM 选择器也失败 → 返回空列表，上层打印警告

### 5.3 `chart_builder.py` — 图表构建

#### 职责

在品类 Sheet 中按单位拆图：同一单位的商品合并到一张折线图，不同单位分图。

#### 核心函数

| 函数 | 说明 |
|------|------|
| `build_charts_for_sheet(ws, cat_name, records)` | 入口：读取该品类全部记录，按单位分组，调用 `_add_chart_block` |
| `_add_chart_block(ws, start_row, unit, comm_names, data_matrix)` | 在指定行写入透视数据 + 嵌入折线图 |
| `_pivot_records(records)` | 将流水记录透视成 `{date: {comm_name: price}}` |

#### 透视逻辑

```
输入: [{日期:"06-20", 商品:"WTI", 价格:75}, {日期:"06-20", 商品:"Brent", 价格:80}, ...]
输出:
      日期      WTI     Brent
      06-20     75.00   80.00
      06-21     80.50   82.00
      ...
```

透视数据写入 Excel 的 A-C 列（日期 + 各商品列），折线图引用此数据区域，X 轴 = 日期列，Y 轴 = 各商品价格列。

#### 图表规格

```python
chart = LineChart()
chart.width  = 24
chart.height = 14
chart.y_axis.delete = False          # 显示 Y 轴刻度数字
chart.x_axis.delete = False          # 显示 X 轴日期标签
chart.y_axis.numFmt = '#,##0'       # Y 轴数字格式
chart.y_axis.majorTickMark = "out"
chart.y_axis.majorGridlines = ChartLines()
chart.x_axis.majorTickMark = "out"
chart.x_axis.tickLblPos = "nextTo"
chart.legend.position = "b"          # 图例在底部
```

每商品一条折线，不同颜色，数据点带圆形标记。

### 5.4 `storage.py` — Excel 存储

#### 核心函数

| 函数 | 说明 |
|------|------|
| `init_workbook()` | 首次创建 Excel → 建「汇总」Sheet + 表头 |
| `append_records(records)` | 追加数据：按品类写到对应 Sheet，去重，刷新图表 |
| `get_or_create_sheet(wb, name)` | 获取 Sheet，不存在则创建并写入表头 |
| `rebuild_summary_sheet(wb)` | 每次追加后重建汇总页 |
| `read_all_records(cat_name=None)` | 读取全量或指定品类数据 |
| `get_stats()` | 返回统计信息（总记录数、品类分布、日期范围） |

#### 去重策略

```
去重键: (日期, 品类, 商品名称)
```

- 写入前读取该 Sheet 已有数据 → 构建 `set` 去重键
- 新记录逐条检查 → 已存在跳过，不存在追加到数据区末尾
- 数据仅追加（append），历史行永不修改或删除

#### Sheet 布局（品类 Sheet）

```
行 1:     标题: "📊 能源 — 价格趋势"
行 3:     [透视数据] 日期 | WTI原油 | Brent原油
行 4-10:  [透视数据] 各日期行
          ← 折线图嵌入在透视表右侧 (H列起) →
行 12:    [透视数据] 日期 | 液化天然气     (不同单位 = 新图)
行 13-19: [透视数据] 各日期行
          ← 折线图 →
行 22:    标题: "📋 能源 — 原始数据"
行 23:    表头: 日期 | 商品名称 | 价格 | 单位 | 七日涨跌幅(%) | 记录时间
行 24+:   数据行... (冻结此行, 自动筛选)
```

#### 汇总 Sheet 布局

```
行 1:     标题: "📊 大宗商品价格日报 — 汇总"
行 2:     副标题: "更新日期: 2026-06-22  数据来源: 生意社(100ppi.com)"
行 4:     表头: 品类 | 代表商品 | 单位 | 最新价格 | 单日涨跌 | 7日走势 | 趋势
行 5+:    分组行: "▎能源" (绿底, 合并单元格)
行 6+:    数据行: | WTI原油 | 美元/桶 | 74.50 | ↓ -2.99% | ▇▆▅▄▂▃▁ | ↓ 下行
```

#### 汇总页计算逻辑

```
对每个品类的每个商品:
  最新价格 = 该商品最近一个日期的价格
  单日涨跌 = (最新价格 - 前一天价格) / 前一天价格 × 100%
           → > 0: 绿底红字 "↑ +X.XX%"
           → < 0: 红底绿字 "↓ X.XX%"
           → ≈ 0: 灰字 "→ 0.00%"
  7日走势  = sparkline (▁▂▃▄▅▆▇), 7天价格缩放到 7 级
  趋势     = 比较第1天 vs 第7天
           → 涨幅 > 1%: "↑ 上行" (红字)
           → 跌幅 > 1%: "↓ 下行" (绿字)
           → 否则: "→ 震荡" (灰字)
```

### 5.5 `main.py` — CLI 入口

```bash
python3 main.py run       # 一键: 采集 → 保存 → 显示统计 (默认)
python3 main.py scrape    # 仅采集保存
python3 main.py stats     # 查看数据统计
```

#### 执行流程 (`run`)

```
1. 调用 scraper.scrape_homepage() → list[dict]
2. 若为空 → 打印错误，退出
3. 调用 storage.append_records(records) → int (新增条数)
4. 若新增 > 0 → 调用 chart_builder 重建所有品类 Sheet 的图表
5. 调用 storage.rebuild_summary_sheet()
6. 打印统计信息
```

---

## 6. 调度方案 (macOS)

### 6.1 launchd Plist

文件路径: `~/Library/LaunchAgents/com.commodity.tracker.plist`

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.commodity.tracker</string>
    <key>ProgramArguments</key>
    <array>
        <string>/usr/bin/python3</string>
        <string>/path/to/commodityTracker/main.py</string>
        <string>run</string>
    </array>
    <key>StartCalendarInterval</key>
    <dict>
        <key>Hour</key><integer>9</integer>
        <key>Minute</key><integer>30</integer>
    </dict>
    <key>StandardOutPath</key>
    <string>/tmp/commodity_tracker.log</string>
    <key>StandardErrorPath</key>
    <string>/tmp/commodity_tracker.err</string>
</dict>
</plist>
```

安装 & 加载:

```bash
mkdir -p ~/Library/LaunchAgents
cp commodityTracker/com.commodity.tracker.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.commodity.tracker.plist
launchctl list | grep commodity   # 验证
```

### 6.2 手动执行

```bash
cd /Users/xiangshaoxiong/Reasonix/commodityTracker
python3 main.py run       # 一键: 采集 + 保存 + 统计
python3 main.py stats     # 仅查看统计
```

### 6.3 日常管理

| 操作 | 命令 |
|------|------|
| 查看是否在运行 | `launchctl list \| grep commodity` |
| 手动触发一次 | `launchctl start com.commodity.tracker` |
| 查看今日日志 | `tail -30 /tmp/commodity_tracker.log` |
| 查看错误 | `cat /tmp/commodity_tracker.err` |
| 暂停定时任务 | `launchctl unload ~/Library/LaunchAgents/com.commodity.tracker.plist` |
| 恢复定时任务 | `launchctl load ~/Library/LaunchAgents/com.commodity.tracker.plist` |
| 新增跟踪商品 | 编辑 `config.py` 的 `TRACKED_COMMODITIES` 列表，加一行即可 |

### 6.4 数据文件

```
data/commodity_prices.xlsx
├── 汇总    按品类汇总: 最新价格 + 单日涨跌箭头 + 7日走势sparkline + 趋势
├── 能源    折线图(按单位拆) + 原始数据表
├── 化工    ...
├── 有色    白银/黄金/铜/铝, 元/吨图 + 元/克图
├── ...
```

---

## 7. 边界条件 & 容错

| 场景 | 处理 |
|------|------|
| 网站不可访问 | 打印错误，退出码 1，不破坏已有数据 |
| 今日数据已存在 | 去重跳过，打印 "0 条新增" |
| 新品类出现 | 自动创建对应 Sheet，无需改代码 |
| 某品类无数据 | 跳过该品类，不影响其他 |
| 已有 Excel 被外部打开 | openpyxl 写入报错，打印提示关闭文件后重试 |
| 图表数据更新 | 每次追加后全量重建图表区（透视 + 折线图），保证最新 |

---

## 8. 交付清单

- [x] 效果预览 Excel → `demo_preview.py`
- [x] `config.py` — 含 `TRACKED_COMMODITIES` 跟踪商品列表
- [x] `scraper.py` — 首页采集 + sf 页面 3 级容错采集
- [x] `chart_builder.py` — 按单位拆图、自动刷新
- [x] `storage.py` — 去重追加、自动建 Sheet、汇总页重建
- [x] `main.py` — CLI `run` / `stats`
- [x] `requirements.txt`
- [x] 首次运行验证（真实数据 12 条）
- [x] launchd 定时任务已安装（每日 9:30）
- [x] `DESIGN.md` 使用说明
