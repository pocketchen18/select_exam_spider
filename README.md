# 成绩监控自动化脚本说明文档

本脚本基于 Playwright 开发，旨在实现对教务系统成绩的自动监控。当检测到新成绩录入或成绩更新时，会通过 Windows 桌面弹窗和邮件通知用户，并展示总评与分项明细。

## 1. 环境准备

- **Python 环境**：使用 `uv` 创建虚拟环境 `.venv`
- **浏览器**：需安装 **Microsoft Edge**
- **依赖库**：`playwright`

推荐流程：

```powershell
uv venv
uv pip install playwright
.venv\Scripts\python.exe -m playwright install msedge
```

## 2. 核心配置指南

脚本使用 `config.json` 管理固定配置（URL 与选择器）。敏感信息（账号、邮箱授权码、Cookie）通过本地网页输入后保存到 `user_secrets.json`，仅存于本机。

### 2.1 获取成绩查询 URL 与 Cookie

由于教务系统有登录保护，请从浏览器获取登录后的网址和 Cookie：

1. **登录系统**：使用 Edge 浏览器打开教务系统，进入“学生成绩查询”页面。
2. **确认网址**：复制成绩查询页 URL，填写到 `config.json` 的 `"url"` 字段中。
3. **获取 Cookie**：打开开发者工具（F12）-> 网络 (Network) -> 刷新页面。
4. **找到请求记录**：点击 `cjcx_cxDgXscj.html...` 请求，在 Request Headers 中找到 `Cookie`。
5. **在网页中填写**：脚本启动后会打开 `http://127.0.0.1:8000`，将 `JSESSIONID`、`wengine_new_ticket`、`route` 填入表单。

### 2.2 获取网易邮箱授权码

发送邮件需要“授权码”，不是登录密码：

1. 登录网页版网易邮箱 (163.com)。
2. 设置 -> POP3/SMTP/IMAP -> 开启 POP3/SMTP 服务。
3. 新增授权码并保存。
4. 脚本启动后在本地网页输入邮箱地址与授权码。

### 2.3 配置文件 `config.json` 详解

```json
{
    "url": "成绩查询的网址",
    "check_interval_seconds": 1800,
    "cookies": [
        {
            "name": "JSESSIONID",
            "value": "",
            "domain": "jwglxt.gpnu.edu.cn",
            "path": "/"
        },
        { "name": "wengine_new_ticket", "value": "", "domain": "jwglxt.gpnu.edu.cn", "path": "/" },
        { "name": "route", "value": "", "domain": "jwglxt.gpnu.edu.cn", "path": "/" }
    ],
    "email_config": {
        "smtp_server": "smtp.163.com",
        "smtp_port": 465
    },
    "xpath": {
        "search_button": "/html/body/div[2]/div/div/div[3]/div[2]/button",
        "course_row": "tr.jqgrow",
        "course_name_cell": "td[aria-describedby$='_kcmc']",
        "total_score_cell": "td[aria-describedby$='_cj']",
        "detail_button": "a[title='查看成绩详情'], a:has-text('查看成绩详情')",
        "detail_modal": "div[role='dialog']:has-text('查看成绩详情')",
        "detail_rows": "table tbody tr",
        "detail_item_cell": "td:nth-child(1)",
        "detail_ratio_cell": "td:nth-child(2)",
        "detail_score_cell": "td:nth-child(3)",
        "detail_close_button": "button:has-text('关闭')"
    },
    "login": {
        "username_input": "#yhm",
        "password_input": "#mm",
        "submit_button": "#dl"
    }
}
```

## 3. 运行脚本

```powershell
.venv\Scripts\python.exe spider.py
```

启动后浏览器会自动打开本地输入页 `http://127.0.0.1:8000`，填写账号、邮箱、Cookie 后脚本开始运行。

## 4. 功能逻辑

1. **自动登录**：优先注入 Cookie；若输入账号密码则尝试自动登录。
2. **自动查询**：定位并点击“查询”按钮。
3. **成绩对比**：获取课程总评与分项明细，对比历史记录，发现变化即提醒。
4. **即时提醒**：
   - **桌面弹窗**：显示课程总评与分项成绩。
   - **邮件提醒**：发送详细成绩明细到指定邮箱。
5. **定时任务**：脚本持续运行，每隔指定间隔检查一次。

## 5. 注意事项

- **Cookie 时效**：Cookie 有有效期，过期需重新抓取并在本地网页更新。
- **邮箱拦截**：网易等邮箱对自动化发信审查较严，若发送失败可更换发件邮箱。
- **安全建议**：`user_secrets.json` 仅用于本机保存，已加入 `.gitignore`，请勿提交到仓库。
