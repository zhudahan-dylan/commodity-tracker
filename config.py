"""
Commodity Tracker — Global Configuration
"""

# ── Target ──
BASE_URL = "https://www.100ppi.com"

# ── Output ──
OUTPUT_DIR = "./data"
EXCEL_FILE = "commodity_prices.xlsx"

# ── Browser ──
BROWSER_TIMEOUT = 30_000       # ms
PAGE_LOAD_TIMEOUT = 60_000     # ms

# ── Excel headers ──
SUMMARY_HEADERS = ["品类", "代表商品", "单位", "最新价格", "单日涨跌", "7日走势", "趋势"]
DATA_HEADERS   = ["日期", "商品名称", "价格", "单位", "七日涨跌幅(%)", "记录时间"]

# ── Tracked commodities (sf pages for individual spot prices) ──
# Each entry: sf_id = the /sf/{id}.html page; used as fallback if dynamic discovery fails.
TRACKED_COMMODITIES = [
    {"sf_id": 551, "name": "黄金", "category": "有色", "unit": "元/克"},
    {"sf_id": 792, "name": "铜",   "category": "有色", "unit": "元/吨"},
    {"sf_id": 827, "name": "铝",   "category": "有色", "unit": "元/吨"},
]

# Category page used for dynamic URL discovery (有色金属)
NONFERROUS_LIST_URL = "https://www.100ppi.com/news/list-359-1.html"

# ── Charts ──
CHART_WIDTH  = 24
CHART_HEIGHT = 14
SERIES_COLORS = ["2F5496", "C00000", "548235", "BF8F00", "7030A0", "31869B"]
