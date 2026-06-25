"""
Scraper — Playwright-based extraction of daily commodity price rankings
from https://www.100ppi.com homepage "商品涨跌榜" section.
"""

import asyncio
import re
from datetime import datetime
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

from config import (
    BASE_URL, BROWSER_TIMEOUT, PAGE_LOAD_TIMEOUT,
    TRACKED_COMMODITIES, NONFERROUS_LIST_URL,
)


# ── Helpers ────────────────────────────────────────────

def _parse_price(raw: str):  # -> float | None
    """'1,945,000.00' | '76.79' → float"""
    if not raw:
        return None
    try:
        return float(raw.replace(",", "").strip())
    except ValueError:
        return None


def _parse_change_pct(raw: str):  # -> float | None
    """'-12.45%' | '+9.06%' → float"""
    if not raw:
        return None
    m = re.search(r"([+-]?\d+\.?\d*)%", raw)
    return float(m.group(1)) if m else None


# ── Main scraper ───────────────────────────────────────

async def scrape_homepage():  # -> list[dict]
    """
    Open 100ppi.com homepage, extract daily 商品涨跌榜 data.
    Returns list of record dicts.
    """
    records: list[dict] = []
    scrape_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1920, "height": 1080},
        )
        page = await context.new_page()
        page.set_default_timeout(PAGE_LOAD_TIMEOUT)

        try:
            print(f"[{scrape_time}] 打开 {BASE_URL} ...")
            await page.goto(BASE_URL, wait_until="networkidle",
                            timeout=PAGE_LOAD_TIMEOUT)
            await asyncio.sleep(2)   # 额外等待动态内容

            # ── 提取页面日期 ──
            page_date = await _extract_page_date(page)

            # ── 通道 1 文本解析 ──
            full_text = await page.evaluate("() => document.body.innerText")
            records = _parse_from_text(full_text, page_date, scrape_time)

            # ── 通道 2 DOM fallback ──
            if not records:
                print("  文本解析无结果, 尝试 DOM 提取...")
                records = await _parse_from_dom(page, page_date, scrape_time)

        except PlaywrightTimeout:
            print("  ⚠️ 页面加载超时")
        except Exception as exc:
            print(f"  ❌ 采集异常: {exc}")
        finally:
            await browser.close()

    return records


# ── Page date ──────────────────────────────────────────

async def _extract_page_date(page) -> str:
    """Try to find date on the page, fallback to today."""
    try:
        # Look near "商品涨跌榜" heading
        el = await page.query_selector("text=商品涨跌榜")
        if el:
            parent_text = await page.evaluate(
                "(el) => el.closest('div')?.textContent || ''", el
            )
            m = re.search(r"(\d{4}-\d{2}-\d{2})", parent_text)
            if m:
                return m.group(1)
    except Exception:
        pass
    # Also try the "大宗行业榜" section
    try:
        el = await page.query_selector("text=大宗行业榜")
        if el:
            parent_text = await page.evaluate(
                "(el) => el.closest('div')?.textContent || ''", el
            )
            m = re.search(r"(\d{4}-\d{2}-\d{2})", parent_text)
            if m:
                return m.group(1)
    except Exception:
        pass
    return datetime.now().strftime("%Y-%m-%d")


# ── Text parser ────────────────────────────────────────

def _parse_from_text(text: str, page_date: str, scrape_time: str) -> list[dict]:
    """
    Parse the 商品涨跌榜 block from page plain text.

    Expected structure:
        商品涨跌榜
        2026-06-22

        商品名称         价格        七日涨跌幅
        
        能源
        WTI原油         76.79       -12.45%
        Brent           79.55       -11.98%
        
        化工
        丙烯            7384.33     -16.41%
        ...
    """
    idx = text.find("商品涨跌榜")
    if idx == -1:
        print("  未找到 '商品涨跌榜' 锚点")
        return []

    section = text[idx: idx + 5000]

    # Try to extract date from the first 300 chars
    dm = re.search(r"(\d{4}-\d{2}-\d{2})", section[:300])
    effective_date = dm.group(1) if dm else page_date

    # Split by category names
    cat_names = "能源|化工|橡塑|纺织|有色|钢铁|建材|农副"
    # Pattern: category name on its own line, then data lines
    pattern = re.compile(
        rf"^({cat_names})\s*$\n((?:(?!^{cat_names}\s*$).+\n)+)",
        re.MULTILINE,
    )

    records: list[dict] = []
    for m in pattern.finditer(section):
        cat = m.group(1)
        block = m.group(2)
        lines = [l.strip() for l in block.split("\n") if l.strip()]

        # Skip non-data lines
        data_lines = [
            l for l in lines
            if l not in cat_names
            and not l.startswith("商品名称")
            and not l.startswith("价格")
            and "涨跌幅" not in l
        ]

        # Group into (name, price, change) triples
        i = 0
        while i + 2 < len(data_lines):
            name   = data_lines[i]
            price  = data_lines[i + 1]
            change = data_lines[i + 2]

            # Guard: name should look like a commodity, price should be numeric
            if not re.match(r"^[\u4e00-\u9fff\w\(\)（）\-\.]+", name):
                i += 1
                continue
            if not re.match(r"^[\d,\.]+$", price):
                i += 1
                continue

            p = _parse_price(price)
            c = _parse_change_pct(change)

            if p is not None:
                records.append({
                    "日期": effective_date,
                    "品类": cat,
                    "商品名称": name,
                    "价格": p,
                    "单位": "元/吨",
                    "七日涨跌幅(%)": c,
                    "记录时间": scrape_time,
                })
            i += 3

    print(f"  文本解析: {len(records)} 条记录")
    return records


# ── DOM fallback ───────────────────────────────────────

async def _parse_from_dom(page, page_date: str, scrape_time: str) -> list[dict]:
    """Fallback: walk DOM for structured ranking data."""
    records: list[dict] = []

    try:
        rows = await page.evaluate("""
            () => {
                const results = [];
                const cats = ['能源','化工','橡塑','纺织','有色','钢铁','建材','农副'];
                let currentCat = '';
                const allText = document.body.innerText;
                const idx = allText.indexOf('商品涨跌榜');
                if (idx === -1) return results;
                const section = allText.slice(idx, idx + 5000);
                const lines = section.split('\\n').map(l => l.trim()).filter(Boolean);
                for (let i = 0; i < lines.length; i++) {
                    if (cats.includes(lines[i])) { currentCat = lines[i]; continue; }
                    if (!currentCat) continue;
                    const price = parseFloat(lines[i+1]?.replace(/,/g, ''));
                    if (!isNaN(price) && lines[i] && !cats.includes(lines[i])) {
                        const change = lines[i+2] || '';
                        results.push({cat: currentCat, name: lines[i], price: String(price), change});
                        i += 2;
                    }
                }
                return results;
            }
        """)

        for r in rows:
            p = _parse_price(r["price"])
            if p is not None:
                records.append({
                    "日期": page_date,
                    "品类": r["cat"],
                    "商品名称": r["name"],
                    "价格": p,
                    "单位": "元/吨",
                    "七日涨跌幅(%)": _parse_change_pct(r.get("change", "")),
                    "记录时间": scrape_time,
                })

        print(f"  DOM 解析: {len(records)} 条记录")
    except Exception as exc:
        print(f"  DOM 解析异常: {exc}")

    return records


# ── Individual commodity SF-page scraper ──────────────

async def _discover_sf_urls(page) -> dict:
    """
    Level 1 — Dynamic discovery: find commodity sf URLs from the 有色 category page.
    Returns {commodity_name: sf_url}.
    """
    try:
        await page.goto(NONFERROUS_LIST_URL, wait_until="networkidle",
                        timeout=PAGE_LOAD_TIMEOUT)
        await asyncio.sleep(1)

        links = await page.evaluate("""
            () => {
                const as = document.querySelectorAll('a[href*="/sf/"]');
                const result = {};
                as.forEach(a => {
                    const name = a.textContent.trim();
                    const href = a.href;
                    if (name && href.includes('/sf/')) {
                        result[name] = href;
                    }
                });
                return result;
            }
        """)
        return links
    except Exception as exc:
        print(f"   动态发现 URL 失败: {exc}")
        return {}


async def scrape_commodity_sf(commodity_info: dict):  # -> dict | None
    """
    Scrape a single commodity's latest spot price from its /sf/{id}.html page.

    3-level fallback:
      1. Discover URL dynamically from the 有色 category page
      2. Use the known /sf/{id}.html pattern
      3. If both fail, print warning and return None

    Returns a record dict (same format as homepage scraper), or None.
    """
    name     = commodity_info["name"]
    sf_id    = commodity_info["sf_id"]
    category = commodity_info["category"]
    unit     = commodity_info["unit"]
    scrape_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
        )
        page = await context.new_page()
        page.set_default_timeout(PAGE_LOAD_TIMEOUT)

        target_url = None

        try:
            # ── Level 1: dynamic discovery ──
            discovered = await _discover_sf_urls(page)
            if name in discovered:
                target_url = discovered[name]
                print(f"  {name}: 动态发现 → {target_url}")
            else:
                # ── Level 2: known pattern fallback ──
                target_url = f"{BASE_URL}/sf/{sf_id}.html"
                print(f"  {name}: 未动态发现, 使用已知模式 → {target_url}")

            # ── Navigate to the sf page ──
            await page.goto(target_url, wait_until="networkidle",
                            timeout=PAGE_LOAD_TIMEOUT)
            await asyncio.sleep(1)

            # ── Parse the table via DOM ──
            data = await page.evaluate("""
                () => {
                    const tables = document.querySelectorAll('table');
                    for (const table of tables) {
                        const rows = table.querySelectorAll('tr');
                        let dateRow = null;
                        let priceRow = null;
                        for (const row of rows) {
                            const cells = row.querySelectorAll('td, th');
                            const firstCell = (cells[0]?.textContent || '').trim();
                            if (firstCell === '日期' || firstCell.includes('日期')) {
                                dateRow = row;
                            }
                            if (firstCell === '现货价格' || firstCell.includes('现货价格')) {
                                priceRow = row;
                            }
                        }
                        if (dateRow && priceRow) {
                            const dateCells = dateRow.querySelectorAll('td, th');
                            const priceCells = priceRow.querySelectorAll('td, th');
                            const dates = [];
                            const prices = [];
                            for (let i = 1; i < dateCells.length; i++) {
                                const d = (dateCells[i].textContent || '').trim();
                                const p = (priceCells[i]?.textContent || '').trim();
                                if (d && p) {
                                    dates.push(d);
                                    prices.push(p);
                                }
                            }
                            return {dates, prices};
                        }
                    }
                    return null;
                }
            """)

            if not data or not data.get("dates") or not data.get("prices"):
                print(f"  {name}: 无法解析价格数据")
                await browser.close()
                return None

            dates  = data["dates"]
            prices = data["prices"]

            if len(dates) != len(prices) or len(dates) == 0:
                print(f"  {name}: 日期/价格数量不匹配")
                await browser.close()
                return None

            # Latest = last entry
            latest_date_raw = dates[-1]
            latest_price_raw = prices[-1]

            # Parse date: MM-DD → YYYY-MM-DD
            today = datetime.now()
            try:
                parts = latest_date_raw.strip().split("-")
                month, day = int(parts[0]), int(parts[1])
                year = today.year
                record_date = f"{year}-{month:02d}-{day:02d}"
            except (ValueError, IndexError):
                record_date = today.strftime("%Y-%m-%d")

            price = _parse_price(latest_price_raw)
            if price is None:
                print(f"  {name}: 价格解析失败: {latest_price_raw}")
                await browser.close()
                return None

            print(f"  {name}: {record_date} | {price} {unit}")

            await browser.close()
            return {
                "日期": record_date,
                "品类": category,
                "商品名称": name,
                "价格": price,
                "单位": unit,
                "七日涨跌幅(%)": None,
                "记录时间": scrape_time,
            }

        except PlaywrightTimeout:
            print(f"  ⚠️ {name}: 页面加载超时 → {target_url}")
        except Exception as exc:
            print(f"  ⚠️ {name}: 采集异常: {exc}")

        await browser.close()
        return None


async def scrape_tracked_commodities() -> list:
    """Scrape all tracked individual commodities from their sf pages."""
    records = []
    for comm in TRACKED_COMMODITIES:
        rec = await scrape_commodity_sf(comm)
        if rec:
            records.append(rec)
    return records


# ── Standalone test ────────────────────────────────────
if __name__ == "__main__":
    data = asyncio.run(scrape_homepage())
    print(f"\n=== 首页采集: {len(data)} 条 ===")
    for r in data:
        chg = f"{r['七日涨跌幅(%)']:+.2f}%" if r["七日涨跌幅(%)"] is not None else "-"
        print(f"  {r['品类']:6s} | {r['商品名称']:12s} | {r['价格']:>10.2f} | {chg}")

    tracked = asyncio.run(scrape_tracked_commodities())
    print(f"\n=== 跟踪商品采集: {len(tracked)} 条 ===")
    for r in tracked:
        print(f"  {r['品类']:6s} | {r['商品名称']:12s} | {r['价格']:>10.2f} | {r['单位']}")
    for r in data:
        chg = f"{r['七日涨跌幅(%)']:+.2f}%" if r["七日涨跌幅(%)"] is not None else "-"
        print(f"  {r['品类']:6s} | {r['商品名称']:12s} | {r['价格']:>10.2f} | {chg}")
