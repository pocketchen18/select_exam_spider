---
title: "select_exam_spider - 教务系统成绩自动化监控工具"
summary: "基于 Playwright 的教务系统成绩自动监控工具，支持邮件和桌面通知，助力学生实时掌握成绩录入动态。"
cover_url: "/content/projects/image/select_exam.png"
start_date: "2026-01-01"
end_date: "2026-01-21"
demo_url: ""
repo_url: "https://github.com/pocketchen18/select_exam_spider"
is_featured: true
tags: ["Python", "Playwright", "Automation", "Education"]
---
# 项目详情

**select_exam_spider** 是一个专为在校学生设计的教务系统成绩自动监控工具。该项目通过 Playwright 模拟浏览器操作，实时监测教务系统中的成绩录入情况。一旦发现新成绩发布，程序将第一时间通过桌面弹窗和电子邮件通知用户，免去了频繁手动刷新的烦恼。

## 核心功能

- **自动化定时巡检**：支持自定义检查间隔，全天候自动监控成绩状态。
- **多维度即时通知**：集成 Windows 桌面气泡通知与 SMTP 邮件推送服务，确保重要信息不遗漏。
- **Cookie 智能维护**：支持从浏览器快速导入 Cookie，并具备过期自动识别与提醒功能。
- **可视化运行状态**：程序在检查过程中可生成截图，方便用户随时核实系统界面状态。
- **Playwright 驱动**：利用高性能的浏览器自动化库，提供比传统爬虫更稳定的交互体验。

## 快速开始

1. 确保已安装 Python 环境。
2. 安装依赖库：`pip install playwright` 且 `playwright install msedge`。
3. 根据 `config.json.example` 配置您的 `config.json` 文件（包含教务系统 URL、Cookie 及邮箱授权信息）。
4. 运行 `spider.py` 启动监控程序。

感谢关注本项目，祝大家都能取得理想的成绩！
