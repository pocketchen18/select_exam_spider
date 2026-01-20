import asyncio
import os
import json
import ctypes
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime
from playwright.async_api import async_playwright

# --- 静态常量 ---
SEEN_COURSES_FILE = "seen_courses.json"
CONFIG_FILE = "config.json"

def load_config():
    """从 config.json 加载配置"""
    if not os.path.exists(CONFIG_FILE):
        raise FileNotFoundError(f"未找到配置文件: {CONFIG_FILE}")
    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def load_seen_courses():
    if os.path.exists(SEEN_COURSES_FILE):
        try:
            with open(SEEN_COURSES_FILE, 'r', encoding='utf-8') as f:
                return set(json.load(f))
        except Exception as e:
            print(f"读取已见课程文件失败: {e}")
    return set()

def save_seen_courses(courses):
    try:
        with open(SEEN_COURSES_FILE, 'w', encoding='utf-8') as f:
            json.dump(list(courses), f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"保存已见课程文件失败: {e}")

def send_email(new_courses, email_config):
    """发送邮件通知"""
    if "YOUR_EMAIL" in email_config["sender_email"]:
        print("跳过邮件发送：请先在 config.json 中填写您的网易邮箱地址。")
        return

    message_text = f"您好，系统于 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 检测到以下科目成绩已更新，请及时登录系统查看：\n\n"
    message_text += "更新科目如下：\n"
    message_text += "-" * 20 + "\n"
    message_text += "\n".join([f"· {course}" for course in new_courses])
    message_text += "\n" + "-" * 20 + "\n\n此邮件由系统自动发送，请勿直接回复。"
    
    msg = MIMEText(message_text, 'plain', 'utf-8')
    msg['From'] = email_config["sender_email"]
    msg['To'] = email_config["receiver_email"]
    msg['Subject'] = Header('教务系统成绩更新提醒', 'utf-8')

    try:
        server = smtplib.SMTP_SSL(email_config["smtp_server"], email_config["smtp_port"])
        server.login(email_config["sender_email"], email_config["sender_password"])
        server.sendmail(email_config["sender_email"], [email_config["receiver_email"]], msg.as_string())
        server.quit()
        print(f"邮件已成功发送至: {email_config['receiver_email']}")
    except Exception as e:
        print(f"邮件发送失败 (可能是被拦截): {e}")

def show_notification(new_courses):
    """桌面弹窗通知"""
    message = "发现新成绩出炉：\n" + "\n".join(new_courses)
    ctypes.windll.user32.MessageBoxW(0, message, "新成绩通知", 0x40 | 0x1)

async def check_grades(context, seen_courses, config):
    page = await context.new_page()
    url = config["url"]
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 正在检查成绩...")
    
    try:
        await page.goto(url, wait_until="networkidle")
        
        # 检查是否跳转到了登录页面
        if "login" in page.url or await page.query_selector(".login-box"):
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [错误] Cookie 已过期或无效，请重新获取并更新 config.json！")
            return

        # 点击查询按钮
        search_xpath = config["xpath"]["search_button"]
        await page.wait_for_selector(f"xpath={search_xpath}", state="visible", timeout=10000)
        await page.click(f"xpath={search_xpath}")
        
        # 等待数据加载
        course_selector = config["xpath"]["course_name_cell"]
        await page.wait_for_selector(course_selector, timeout=15000)
        
        # 获取所有课程名称
        course_elements = await page.query_selector_all(course_selector)
        current_courses = { (await el.inner_text()).strip() for el in course_elements if (await el.inner_text()).strip() }
        
        # 找出新出的成绩
        new_courses = current_courses - seen_courses
        
        if new_courses:
            print(f"发现新成绩: {new_courses}")
            send_email(new_courses, config["email_config"])
            show_notification(new_courses)
            seen_courses.update(new_courses)
            save_seen_courses(seen_courses)
        else:
            print("未发现新成绩。")
            
        await page.screenshot(path="last_check.png")
        
    except Exception as e:
        print(f"检查过程中发生错误: {e}")
    finally:
        await page.close()

async def run():
    config = load_config()
    seen_courses = load_seen_courses()
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, channel="msedge")
        context = await browser.new_context()
        await context.add_cookies(config["cookies"])
        
        try:
            while True:
                await check_grades(context, seen_courses, config)
                interval = config.get("check_interval_seconds", 1800)
                print(f"等待 {interval // 60} 分钟后进行下一次检查...")
                await asyncio.sleep(interval)
        except KeyboardInterrupt:
            print("脚本已停止。")
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(run())
