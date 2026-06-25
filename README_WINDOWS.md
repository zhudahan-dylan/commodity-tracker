# Windows 移植指南

## 构建流程

### 前提

1. 一台 Windows 10/11 电脑
2. Python 3.9+（[python.org](https://www.python.org/downloads/) 下载）

### 一键构建

```powershell
# 1. 把整个 commodityTracker 目录复制到 Windows

# 2. 安装依赖
pip install pyinstaller playwright openpyxl
playwright install chromium

# 3. 运行构建脚本
python build_windows.py
```

构建完成后，`dist/commodity_tracker_windows.zip` 即为发布包。

### 构建产物

```
dist/commodity_tracker_windows.zip
    ├── commodity_tracker.exe    ← 主程序 (~50MB)
    ├── config.py                ← 可编辑的配置
    ├── scraper.py               ← 采集模块
    ├── storage.py               ← 存储模块
    ├── chart_builder.py         ← 图表模块
    ├── data/                    ← Excel 输出目录
    ├── setup.bat                ← 注册定时任务
    └── 使用说明.txt
```

## 架构说明（裸机兼容）

```
用户双击 commodity_tracker.exe
    │
    ├─ Chromium 已安装? ── 是 ──► 开始采集
    │
    └─ 否 ──► 弹提示 "正在下载浏览器组件..."
               │
               ├─ 1. 尝试 bundled playwright driver (PyInstaller 内)
               ├─ 2. 尝试 python -m playwright install chromium
               └─ 3. 尝试 npx playwright install chromium
               │
               └─ 成功 ──► 开始采集
               失败 ──► 提示手动安装
```

### PyInstaller 打包要点

| 关键点 | 做法 |
|--------|------|
| playwright 驱动 | `--collect-all playwright` 包含 Node.js CLI |
| openpyxl 图表 | `--hidden-import openpyxl.chart.axis` 等 |
| asyncio | `--hidden-import asyncio` |
| 运行时模块 | scraper.py / storage.py / chart_builder.py 保持 .py 文件, exe 运行时 import |

### 为什么不把 scraper.py 也编译进 exe？

保持 `.py` 文件独立 → 用户可以编辑跟踪商品配置、甚至可以改采集逻辑而不用重新构建 exe。config.py 同样可编辑。

## 定时调度 vs launchd

| | macOS | Windows |
|------|------|------|
| 调度系统 | launchd | 任务计划程序 (Task Scheduler) |
| 注册方式 | `launchctl load plist` | `schtasks /create` |
| 管理命令 | `launchctl list` | `schtasks /query` |
| 配置文件 | `.plist` (XML) | 命令行参数 |

setup.bat 封装了 schtasks 命令:
```batch
schtasks /create /tn "CommodityTracker" /tr "...\commodity_tracker.exe run" /sc daily /st 09:30 /f
```

## 注意事项

1. **首次运行需要联网** — Chromium 下载 (~130MB) 只发生一次
2. **setup.bat 需管理员权限** — Windows 任务计划程序要求
3. **exe 不要放中文路径** — PyInstaller 对中文路径支持不好
4. **Chromium 缓存位置** — `%LOCALAPPDATA%\ms-playwright\`
