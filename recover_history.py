"""One-shot: recover historical data for tracked commodities from sf pages."""
import asyncio, sys, os
sys.path.insert(0, os.path.dirname(__file__))

from datetime import datetime
from playwright.async_api import async_playwright
from config import TRACKED_COMMODITIES, BASE_URL

async def recover_one(page, comm: dict) -> list[dict]:
    """Scrape ALL historical spot prices for one commodity from its sf page."""
    name = comm["name"]
    sf_id = comm["sf_id"]
    cat = comm["category"]
    unit = comm["unit"]
    scrape_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    url = f"{BASE_URL}/sf/{sf_id}.html"
    print(f"  {name}: {url}")
    await page.goto(url, wait_until="networkidle", timeout=60000)
    await asyncio.sleep(1)

    data = await page.evaluate("""
        () => {
            const tables = document.querySelectorAll('table');
            for (const t of tables) {
                const rows = t.querySelectorAll('tr');
                let dateRow = null, priceRow = null;
                for (const r of rows) {
                    const first = (r.cells[0]?.textContent || '').trim();
                    if (first === '日期' || first.includes('日期')) dateRow = r;
                    if (first === '现货价格' || first.includes('现货价格')) priceRow = r;
                }
                if (dateRow && priceRow) {
                    const dates = [], prices = [];
                    for (let i = 1; i < dateRow.cells.length; i++) {
                        const d = (dateRow.cells[i].textContent || '').trim();
                        const p = (priceRow.cells[i]?.textContent || '').trim();
                        if (d && p) dates.push(d); prices.push(p);
                    }
                    return {dates, prices};
                }
            }
            return null;
        }
    """)

    if not data:
        print(f"    ❌ 解析失败")
        return []

    records = []
    today = datetime.now()
    for d, p in zip(data["dates"], data["prices"]):
        try:
            parts = d.strip().split("-")
            month, day = int(parts[0]), int(parts[1])
            year = today.year
            date_str = f"{year}-{month:02d}-{day:02d}"
        except (ValueError, IndexError):
            continue

        price = None
        try:
            price = float(p.replace(",", "").strip())
        except ValueError:
            continue

        if price:
            records.append({
                "日期": date_str,
                "品类": cat,
                "商品名称": name,
                "价格": price,
                "单位": unit,
                "七日涨跌幅(%)": None,
                "记录时间": scrape_time,
            })

    print(f"    ✅ {len(records)} 条 ({records[0]['日期'] if records else 'none'} ~ {records[-1]['日期'] if records else 'none'})")
    return records


async def main():
    print("=" * 60)
    print("  恢复历史数据 — 从 sf 页面提取完整现货价格记录")
    print("=" * 60)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        all_records = []
        for comm in TRACKED_COMMODITIES:
            recs = await recover_one(page, comm)
            all_records.extend(recs)

        await browser.close()

    if all_records:
        from storage import append_records
        new = append_records(all_records)
        print(f"\n💾 共恢复 {len(all_records)} 条，新增 {new} 条")
    else:
        print("\n❌ 未恢复任何数据")


if __name__ == "__main__":
    asyncio.run(main())
