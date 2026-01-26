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

脚本使用 `config.json` 管理固定配置（URL 与选择器）。敏感信息（账号、邮箱授权码）通过本地网页输入后保存到 `user_secrets.json`，仅存于本机。登录状态会保存在 `pw_profile`，只需登录一次即可长期复用。

### 2.1 获取登录与成绩查询 URL

1. **登录入口**：打开统一认证入口页面，复制登录页地址作为 `login_url`。
2. **成绩查询**：进入“学生成绩查询”页面，复制该地址作为 `grades_url`。

### 2.2 获取网易邮箱授权码

发送邮件需要“授权码”，不是登录密码：

1. 登录网页版网易邮箱 (163.com)。
2. 设置 -> POP3/SMTP/IMAP -> 开启 POP3/SMTP 服务。
3. 新增授权码并保存。
4. 脚本启动后在本地网页输入邮箱地址与授权码。

### 2.3 验证码 OCR 配置（OpenAI 兼容）

在本地输入页填写以下字段：

- `base_url`：OpenAI 兼容接口地址（例如 `https://api.openai.com/v1`）
- `model`：模型名称
- `api_key`：接口密钥

脚本会自动识别算术验证码并输入结果，识别失败会自动刷新并重试 3 次，仍失败则保持浏览器打开并等待手动输入。提交登录时会通过回车键触发，无需点击登录按钮。

### 2.4 配置文件 `config.json` 详解

```json
{
    "login_url": "统一认证登录入口",
    "grades_url": "成绩查询页面",
    "url": "(兼容旧版本) 成绩查询网址",
    "check_interval_seconds": 1800,
    "user_data_dir": "pw_profile",
    "email_config": {
        "smtp_server": "smtp.163.com",
        "smtp_port": 465
    },
    "ocr": {
        "base_url": "",
        "model": "",
        "timeout_seconds": 30,
        "max_retries": 3
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
        "username_input": "#userName",
        "password_input": "#password",
        "submit_button": "button.index-submit-1UOCo.index-logining-HO9Db",
        "captcha_input": "#captcha",
        "captcha_image": ".index-captcha-2FKeU img",
        "captcha_image_fallback": "div[class^='index-captcha-'] img",
        "captcha_refresh": ".index-codeMask-20jm4",
        "switch_to_password": ".index-qr_btn-3JpGS.index-sj_btn-11Xsa",
        "switch_account_btn": "/html/body/div[1]/span/div[3]/div/div[1]/div/div[2]/div[6]/span"
    }
}
```

## 3. 核心功能特性

- **多轮登录支持**：脚本支持检测并处理主页面及 iframe 嵌套内的多轮登录界面（最高 5 轮）。
- **账号切换自动化**：针对部分需要点击“切换账号登录”才能显示表单的页面，可通过 `switch_account_btn` 的 XPath 路径实现自动点击。
- **CAS 统一认证跳转**：自动检测 CAS 统一身份认证提示，执行授权跳转，并验证登录成功状态。
- **验证码跨框架识别**：支持识别主页面及 iframe 内的验证码，具备自动刷新与重试机制。
- **配置驱动**：所有核心选择器、URL 及 OCR 参数均可通过 `config.json` 或本地输入界面灵活配置。

## 4. 运行脚本

```powershell
.venv\Scripts\python.exe spider.py
```

启动后浏览器会自动打开本地输入页 `http://127.0.0.1:8000`，填写登录入口 URL、成绩查询 URL、账号、邮箱及 OCR 配置后脚本开始运行。

## 4. 功能逻辑

1. **自动登录与复用**：优先复用 `pw_profile` 登录态。若失效，脚本会自动尝试填写账号密码、识别验证码并处理 CAS 跳转；若遇到复杂校验（如滑块），则进入手动登录等待模式。
2. **多层级界面适配**：支持 iframe 内嵌的登录表单，并能自动点击“切换账号登录”按钮以显示输入框。
3. **自动查询**：定位并点击“查询”按钮，支持自定义查询页面 URL。
4. **成绩对比**：获取课程总评与分项明细，对比历史记录，发现变化即提醒。
5. **即时提醒**：
   - **桌面弹窗**：显示课程总评与分项成绩。
   - **邮件提醒**：发送详细成绩明细到指定邮箱。
6. **定时任务**：脚本持续运行，每隔指定间隔检查一次。

## 5. 注意事项

- **登录失效**：若账号被迫重新登录，删除 `pw_profile` 后再运行并手动登录一次。
- **验证码识别**：需提供 OpenAI 兼容接口，OCR 失败会自动重试后提示手动输入。
- **邮箱拦截**：网易等邮箱对自动化发信审查较严，若发送失败可更换发件邮箱。
- **安全建议**：`user_secrets.json` 仅用于本机保存，已加入 `.gitignore`，请勿透露给他人。
