# 成绩监控自动化脚本说明文档

本脚本基于 Playwright 开发，旨在实现对教务系统成绩的自动监控。当检测到新成绩录入时，会通过 Windows 桌面弹窗和邮件通知用户。

## 1. 环境准备

- **Python 环境**：建议使用虚拟环境 `D:\code\Python\Findwork\.venv`
- **浏览器**：需安装 **Microsoft Edge**
- **依赖库**：`playwright`, `asyncio`

## 2. 核心配置指南

脚本使用 `config.json` 进行统一管理。请按照以下详细步骤获取并填写参数。

### 2.1 如何获取 网址 (URL) 与 Cookie

由于教务系统有登录保护，您需要手动从浏览器获取登录后的网址和 Cookie 信息：

1. **登录系统**：使用 Edge 浏览器打开教务系统，点击进入“学生成绩查询”页面。
2. **确认网址**：请确保您当前页面的网址与[截图](file:///d:/code/Python/exam_spider/%E5%B1%8F%E5%B9%95%E6%88%AA%E5%9B%BE%202026-01-20%20003255.png)中顶部的地址栏一致（通常为 `https://jwglxt.gpnu.edu.cn/jwglxt/cjcx/cjcx_cxDgXscj.html?gnmkdm=N305005&layout=default`）。
   - **将该网址完整复制**，填写到 `config.json` 的 `"url"` 字段中。
     网址是成绩查询界面的网址
3. **打开开发者工具**：按下键盘上的 **F12** 键（或在页面点击右键选择“检查”）。
4. **进入网络面板**：在弹出的工具栏顶部点击 **“网络” (Network)** 选项卡。
5. **刷新页面**：按 **F5** 刷新页面，此时下方会刷新出很多记录。
6. **找到请求记录**：
   - 在左侧列表中找到名为 `cjcx_cxDgXscj.html...` 的记录并点击它。
   - 在右侧出现的面板中，点击 **“标头” (Headers)** 选项卡。
   - 向下滚动找到 **“请求标头” (Request Headers)** 区域。
   - 寻找 `Cookie:` 这一行，你会看到类似 `JSESSIONID=...; wengine_new_ticket=...; route=...` 的字符串。
7. **填写到 config.json**：
   - 将对应的 `JSESSIONID`、`wengine_new_ticket` 和 `route` 的值分别复制并粘贴到 `config.json` 的 `value` 字段中。

### 2.2 如何获取网易邮箱授权码

脚本发送邮件需要专门的“授权码”，而不是您的登录密码：

1. **登录网页版网易邮箱** (163.com)。
2. **进入设置**：点击顶部菜单栏的 **“设置”** -> **“POP3/SMTP/IMAP”**。
3. **开启服务**：确保 “POP3/SMTP服务” 已勾选为 **“开启”**。
4. **新增授权码**：点击页面下方的 **“新增授权码”** 按钮。
5. **验证身份**：根据提示发送短信验证码。
6. **获取并保存**：页面会显示一串 16 位的字母（如 `DQ2YbcQkCM9J4Kgo`），**请立即复制并保存**，它只显示一次。
7. **填写到 config.json**：将这串字母填写到 `sender_password` 字段。

### 2.3 配置文件 config.json 详解

打开 `config.json`，根据以下说明填入信息：

```json
{
    "url": "成绩查询的网址",
    "check_interval_seconds": 1800, // 每隔多少秒检查一次（1800秒=30分钟）
    "cookies": [
        {
            "name": "JSESSIONID",
            "value": "这里填入从浏览器抓取的 JSESSIONID 值",
            "domain": "jwglxt.gpnu.edu.cn",
            "path": "/"
        },
        // ... 其他两个 cookie 同理填入 value ...
    ],
    "email_config": {
        "smtp_server": "smtp.163.com",
        "smtp_port": 465,
        "sender_email": "你的发件邮箱@163.com",
        "sender_password": "这里填入 16 位的授权码",
        "receiver_email": "接收通知的邮箱@qq.com"
    }
}
```

## 3. 运行脚本

使用指定虚拟环境运行：

```powershell
D:\code\Python\Findwork\.venv\Scripts\python.exe spider.py
```

## 4. 功能逻辑

...（略）

## 5. 常见问题与注意事项

- **电脑休眠问题**：若电脑进入“睡眠”或“休眠”状态，脚本将停止运行。若需长期监控，请在 Windows 电源设置中将“使计算机进入睡眠状态”设为“从不”，并将“关闭盖子”设为“不采取任何操作”。
- **无头模式运行**：如果您不希望每次检查都弹出浏览器窗口，可以将 `spider.py` 中的 `headless=False` 改为 `headless=True`。
- **垃圾邮件拦截**：网易等邮箱对自动化发信审查较严。若邮件发送失败（报错 554），请检查 `config.json` 中的配置。
- **Cookie 时效**：Cookie 过期后脚本将无法进入页面。判断方法：
  1. 脚本控制台会输出 `[错误] Cookie 已过期`。
  2. 查看项目目录下的 `last_check.png`，如果截图显示的是登录页面，说明已过期。
  3. 请定期按照 2.1 节重新获取。
