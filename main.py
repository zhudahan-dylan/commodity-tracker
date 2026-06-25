#!/usr/bin/env python3
"""
Commodity Price Tracker — CLI entry point.

Usage:
    python3 main.py run      采集 + 保存 + 统计 (默认)
    python3 main.py stats    查看数据统计
"""

import sys
import os
import asyncio
import argparse
import subprocess
from pathlib import Path


# ═══════════════════════════════════════════════════════
#  Chromium bootstrap (Windows / fresh env)
# ═══════════════════════════════════════════════════════

def _chromium_installed() -> bool:
    """Check whether Playwright Chromium browser is available."""
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            browser.close()
            return True
    except Exception:
        return False


def _install_chromium() -> bool:
    """
    Attempt to download Chromium for Playwright.
    Uses the bundled driver if available (PyInstaller), otherwise falls back
    to the system playwright CLI.
    """
    print("╔══════════════════════════════════════════════╗")
    print("║  首次运行 — 正在安装浏览器组件 (~130MB)...   ║")
    print("╚══════════════════════════════════════════════╝")

    methods = []

    # Method 1: bundled driver (PyInstaller)
    try:
        from playwright._impl._driver import compute_driver_executable
        driver = str(compute_driver_executable())
        if os.path.exists(driver):
            methods.append([driver, "install", "chromium"])
    except Exception:
        pass

    # Method 2: system playwright CLI
    methods.append([sys.executable, "-m", "playwright", "install", "chromium"])

    # Method 3: npx (if node is available)
    methods.append(["npx", "playwright", "install", "chromium"])

    for cmd in methods:
        try:
            print(f"  尝试: {' '.join(cmd)}")
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
            if result.returncode == 0 and _chromium_installed():
                print("  ✅ 浏览器组件安装完成")
                return True
        except Exception:
            continue

    return False


def _ensure_browser():
    """Make sure Chromium is ready before scraping."""
    if _chromium_installed():
        return True
    print("⚠️  未检测到浏览器组件")
    if _install_chromium():
        return True
    print("❌ 浏览器安装失败。请手动安装: pip install playwright && playwright install chromium")
    return False


# ═══════════════════════════════════════════════════════
#  Commands
# ═══════════════════════════════════════════════════════

def cmd_scrape() -> int:
    """Scrape homepage and save to Excel."""
    print("=" * 60)
    print("  大宗商品价格采集工具 — 数据来源: 生意社 (100ppi.com)")
    print("=" * 60)

    # Ensure browser is ready
    if not _ensure_browser():
        return 1

    from scraper import scrape_homepage, scrape_tracked_commodities
    from storage import append_records

    records = asyncio.run(scrape_homepage())

    # Also scrape tracked individual commodities (金/铜/铝 etc.)
    tracked = asyncio.run(scrape_tracked_commodities())
    records.extend(tracked)

    if not records:
        print("\n❌ 未采集到任何数据。请检查网络连接或稍后重试。")
        return 1

    print(f"\n📊 共采集 {len(records)} 条记录:")
    for r in records:
        chg = f"{r['七日涨跌幅(%)']:+.2f}%" if r["七日涨跌幅(%)"] is not None else "-"
        print(f"  {r['品类']:6s} | {r['商品名称']:12s} | {r['价格']:>10.2f} | {chg}")

    new_count = append_records(records)
    if new_count == 0:
        print("  所有记录已存在, 无新增。")
    return 0


def cmd_stats() -> int:
    """Print summary statistics."""
    from storage import get_stats

    stats = get_stats()
    print("=" * 60)
    print("  大宗商品价格数据统计")
    print("=" * 60)
    print(f"  总记录数: {stats['total_records']}")
    if stats["date_range"]:
        print(f"  日期范围: {stats['date_range'][0]} ~ {stats['date_range'][1]}")
    print(f"  数据文件: {stats['file']}")
    print("\n  品类分布:")
    for cat, count in sorted(stats["categories"].items(), key=lambda x: -x[1]):
        bar = "█" * min(count // 2, 40)
        print(f"    {cat:8s} | {count:4d} {bar}")
    return 0


def cmd_run() -> int:
    """Full pipeline: scrape + save + stats."""
    ret = cmd_scrape()
    if ret != 0:
        return ret
    print()
    return cmd_stats()


def main():
    parser = argparse.ArgumentParser(
        description="大宗商品价格采集工具 — 从生意社(100ppi.com)采集每日价格",
    )
    parser.add_argument(
        "command",
        nargs="?",
        default="run",
        choices=["run", "scrape", "stats"],
        help="操作: run (采集+统计), scrape (仅采集), stats (仅统计)",
    )
    args = parser.parse_args()

    cmds = {"run": cmd_run, "scrape": cmd_scrape, "stats": cmd_stats}
    return cmds[args.command]()


if __name__ == "__main__":
    sys.exit(main())
