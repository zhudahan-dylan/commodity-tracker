"""
Windows 构建脚本 — 在 Windows 上运行此脚本，一键打包为 .exe

前置条件: Windows 系统、Python 3.9+

运行:
    pip install pyinstaller
    python build_windows.py

输出:
    dist/commodity_tracker_windows.zip
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist" / "pkg"
NAME = "commodity_tracker"


def run(cmd: list[str], **kwargs):
    print(f"  > {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=str(ROOT), **kwargs)
    if result.returncode != 0:
        print(f"  ❌ 命令失败 (exit {result.returncode})")
        sys.exit(1)


def step(msg: str):
    print(f"\n{'='*60}")
    print(f"  {msg}")
    print(f"{'='*60}")


# ── Step 1: Ensure playwright Chromium is installed ──
def ensure_playwright_browsers():
    step("Step 1: 确保 Playwright + Chromium 已安装")
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            p.chromium.launch(headless=True).close()
        print("  ✅ Chromium 已就绪")
    except Exception:
        print("  安装 playwright browsers...")
        run([sys.executable, "-m", "playwright", "install", "chromium"])


# ── Step 2: PyInstaller ──
def build_exe():
    step("Step 2: PyInstaller 打包")

    # Clean
    for d in ["build", "dist"]:
        shutil.rmtree(ROOT / d, ignore_errors=True)

    hidden_imports = [
        "playwright",
        "playwright._impl",
        "playwright._impl._api_structures",
        "playwright._impl._browser",
        "playwright._impl._browser_context",
        "playwright._impl._browser_type",
        "playwright._impl._connection",
        "playwright._impl._driver",
        "playwright._impl._event_context_manager",
        "playwright._impl._frame",
        "playwright._impl._helper",
        "playwright._impl._js_handle",
        "playwright._impl._network",
        "playwright._impl._object_factory",
        "playwright._impl._page",
        "playwright._impl._transport",
        "playwright.async_api",
        "playwright.sync_api",
        "openpyxl",
        "openpyxl.chart",
        "openpyxl.chart.axis",
        "openpyxl.styles",
        "openpyxl.utils",
        "asyncio",
        "json",
        "re",
        "datetime",
        "argparse",
    ]

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--console",
        f"--name={NAME}",
        "--clean",
        "--noconfirm",
    ]

    for mod in hidden_imports:
        cmd.extend(["--hidden-import", mod])

    # Collect playwright package data (driver, etc.)
    cmd.extend(["--collect-all", "playwright"])

    # Main script
    cmd.append(str(ROOT / "main.py"))

    run(cmd)
    print("  ✅ exe 构建完成")


# ── Step 3: Assemble distribution package ──
def assemble_package():
    step("Step 3: 组装发布包")

    DIST.mkdir(parents=True, exist_ok=True)

    # Copy exe
    exe_src = ROOT / "dist" / f"{NAME}.exe"
    exe_dst = DIST / f"{NAME}.exe"
    shutil.copy(exe_src, exe_dst)
    print(f"  ✅ {exe_dst.name}")

    # Copy config
    shutil.copy(ROOT / "config.py", DIST / "config.py")
    print("  ✅ config.py")

    # Copy other modules (they get imported at runtime)
    for mod in ["scraper.py", "storage.py", "chart_builder.py"]:
        shutil.copy(ROOT / mod, DIST / mod)
        print(f"  ✅ {mod}")

    # Create data dir
    (DIST / "data").mkdir(exist_ok=True)

    # Write setup.bat
    _write_setup_bat(DIST)
    print("  ✅ setup.bat (右键→以管理员身份运行)")

    # Write 使用说明.txt
    _write_readme(DIST)
    print("  ✅ 使用说明.txt")

    # Zip it
    zip_path = ROOT / "dist" / f"{NAME}_windows.zip"
    shutil.make_archive(
        str(zip_path.with_suffix("")),
        "zip",
        str(DIST.parent),
        "pkg",
    )
    print(f"\n  📦 发布包: {zip_path}")


def _write_setup_bat(pkg_dir: Path):
    """Write Windows setup batch script."""
    bat = pkg_dir / "setup.bat"
    bat.write_text(r"""@echo off
chcp 65001 >nul
echo ============================================
echo   大宗商品价格采集工具 — 安装定时任务
echo ============================================
echo.
echo 本脚本将注册 Windows 任务计划程序，每天早上 9:30 自动采集。
echo.

set EXE_PATH=%~dp0commodity_tracker.exe
set TASK_NAME=CommodityTracker

echo 任务名称: %TASK_NAME%
echo 程序路径: %EXE_PATH%
echo 执行时间: 每天 09:30
echo.

schtasks /create /tn "%TASK_NAME%" /tr "\"%EXE_PATH%\" run" /sc daily /st 09:30 /f
if %ERRORLEVEL% EQU 0 (
    echo ✅ 定时任务已注册成功!
    echo.
    echo 管理命令:
    echo   查看任务: schtasks /query /tn "%TASK_NAME%"
    echo   手动运行: schtasks /run /tn "%TASK_NAME%"
    echo   删除任务: schtasks /delete /tn "%TASK_NAME%" /f
) else (
    echo ❌ 注册失败, 请右键本文件 → 以管理员身份运行
)

echo.
pause
""", encoding="utf-8")


def _write_readme(pkg_dir: Path):
    """Write usage instructions."""
    readme = pkg_dir / "使用说明.txt"
    readme.write_text("""大宗商品价格采集工具 — 使用说明
====================================

📦 文件说明
  commodity_tracker.exe   主程序 (双击运行)
  setup.bat               注册每日定时任务 (右键→以管理员身份运行)
  config.py               配置文件 (可修改跟踪商品列表)
  data/                   数据目录 (Excel 文件存放处)

🚀 快速开始
  1. 双击 commodity_tracker.exe → 首次自动下载浏览器组件 → 开始采集
  2. 右键 setup.bat → 以管理员身份运行 → 注册每天 9:30 自动采集
  3. 打开 data/commodity_prices.xlsx 查看数据

⌨️ 命令行用法
  commodity_tracker.exe run      采集 + 保存 (默认)
  commodity_tracker.exe stats    查看统计

🔧 添加跟踪商品
  编辑 config.py, 在 TRACKED_COMMODITIES 列表中添加:
    {"sf_id": 商品ID, "name": "商品名", "category": "品类", "unit": "单位"}

  商品 sf_id 查找方法:
    访问 https://www.100ppi.com → 搜索商品名 → 点击「现期图」
    → URL 中 /sf/ 后面的数字即为 sf_id

📊 数据说明
  每日数据追加到 data/commodity_prices.xlsx
  - 汇总 Sheet: 品类一览 (价格、涨跌、走势)
  - 品类 Sheet: 折线图 + 原始数据表

🛑 停止定时任务
  命令行执行: schtasks /delete /tn "CommodityTracker" /f

❓ 常见问题
  Q: 首次运行报错 "找不到浏览器"
  A: 等待浏览器组件下载完成 (~130MB), 或手动运行:
     pip install playwright && playwright install chromium

  Q: 定时任务没执行
  A: 检查 setup.bat 是否以管理员身份运行成功
     检查: schtasks /query /tn "CommodityTracker"

数据来源: 生意社 (www.100ppi.com)
""", encoding="utf-8")


# ═══════════════════════════════════════════════════════
if __name__ == "__main__":
    if sys.platform != "win32":
        print("⚠️  此脚本需要在 Windows 上运行 (PyInstaller 不支持跨平台编译)。")
        print("   请在 Windows 机器上执行:")
        print("     pip install pyinstaller")
        print(f"     python {Path(__file__).name}")
        sys.exit(0)

    ensure_playwright_browsers()
    build_exe()
    assemble_package()
    print("\n✅ 构建完成!")
    print(f"   发布包: {ROOT / 'dist' / f'{NAME}_windows.zip'}")
    print("   将 zip 发给用户 → 解压 → 双击 exe 即可使用")
