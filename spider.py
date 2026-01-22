import asyncio
import ctypes
import json
import os
import smtplib
import threading
import webbrowser
from datetime import datetime
from email.header import Header
from email.mime.text import MIMEText
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs

from playwright.async_api import async_playwright

SEEN_COURSES_FILE = "seen_courses.json"
CONFIG_FILE = "config.json"
SECRETS_FILE = "user_secrets.json"
INPUT_PORT = 8000
USER_DATA_DIR = "pw_profile"


def load_json_file(path, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as file:
            return json.load(file)
    except Exception as exc:
        print(f"读取文件失败: {path} ({exc})")
        return default


def save_json_file(path, data):
    try:
        with open(path, "w", encoding="utf-8") as file:
            json.dump(data, file, ensure_ascii=False, indent=4)
    except Exception as exc:
        print(f"保存文件失败: {path} ({exc})")


def load_config():
    config = load_json_file(CONFIG_FILE, None)
    if config is None:
        raise FileNotFoundError(f"未找到配置文件: {CONFIG_FILE}")
    return config


def load_seen_courses():
    data = load_json_file(SEEN_COURSES_FILE, {})
    if isinstance(data, list):
        return {name: {"total": "", "components": []} for name in data}
    if isinstance(data, dict):
        return data
    return {}


def save_seen_courses(courses):
    save_json_file(SEEN_COURSES_FILE, courses)


def load_user_secrets():
    data = load_json_file(SECRETS_FILE, {})
    return data if isinstance(data, dict) else {}


def save_user_secrets(secrets):
    save_json_file(SECRETS_FILE, secrets)


def pick_value(*values):
    for value in values:
        if value is not None and str(value).strip() != "":
            return str(value).strip()
    return ""


def render_form_html(defaults):
    cookies = defaults.get("cookies", {})
    email = defaults.get("email", {})
    login = defaults.get("login", {})
    url_value = defaults.get("url", "")
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <title>成绩监控运行配置</title>
  <style>
    body {{ font-family: "Microsoft YaHei", sans-serif; margin: 24px; background: #f7f8fa; }}
    .card {{ background: #fff; padding: 20px 24px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.06); max-width: 820px; margin: 0 auto; }}
    h1 {{ margin-top: 0; font-size: 22px; }}
    fieldset {{ border: 1px solid #e4e6eb; margin-bottom: 16px; border-radius: 8px; }}
    legend {{ padding: 0 8px; color: #1f2d3d; }}
    label {{ display: block; margin: 8px 0 4px; font-size: 14px; }}
    input {{ width: 100%; padding: 8px 10px; border-radius: 6px; border: 1px solid #c9ced6; }}
    button {{ margin-top: 12px; padding: 10px 16px; border: none; border-radius: 6px; background: #3b82f6; color: #fff; font-size: 14px; cursor: pointer; }}
    .hint {{ color: #6b7280; font-size: 12px; margin-top: 4px; }}
  </style>
</head>
<body>
  <div class="card">
    <h1>成绩监控启动配置</h1>
    <form method="post" action="/submit">
      <fieldset>
        <legend>教务系统</legend>
        <label>成绩查询 URL</label>
        <input name="url" value="{url_value}" placeholder="成绩查询页面 URL" />
        <label>JSESSIONID</label>
        <input name="cookie_jsessionid" value="{cookies.get("JSESSIONID", "")}" />
        <label>wengine_new_ticket</label>
        <input name="cookie_wengine" value="{cookies.get("wengine_new_ticket", "")}" />
        <label>route</label>
        <input name="cookie_route" value="{cookies.get("route", "")}" />
        <div class="hint">建议填写 Cookie 以免频繁登录。</div>
        <label>学号/账号（可选）</label>
        <input name="login_username" value="{login.get("username", "")}" />
        <label>登录密码（可选）</label>
        <input name="login_password" type="password" value="{login.get("password", "")}" />
      </fieldset>
      <fieldset>
        <legend>邮件通知</legend>
        <label>发件邮箱</label>
        <input name="sender_email" value="{email.get("sender_email", "")}" />
        <label>授权码</label>
        <input name="sender_password" type="password" value="{email.get("sender_password", "")}" />
        <label>收件邮箱</label>
        <input name="receiver_email" value="{email.get("receiver_email", "")}" />
      </fieldset>
      <button type="submit">保存并开始监控</button>
    </form>
  </div>
</body>
</html>"""


def collect_runtime_secrets(config, stored_secrets):
    defaults = {
        "url": pick_value(stored_secrets.get("url"), config.get("url", "")),
        "cookies": {
            "JSESSIONID": pick_value(
                stored_secrets.get("cookies", {}).get("JSESSIONID"),
                config.get("cookies", [{}])[0].get("value", ""),
            ),
            "wengine_new_ticket": pick_value(
                stored_secrets.get("cookies", {}).get("wengine_new_ticket"),
                config.get("cookies", [{}, {}])[1].get("value", ""),
            ),
            "route": pick_value(
                stored_secrets.get("cookies", {}).get("route"),
                config.get("cookies", [{}, {}, {}])[2].get("value", ""),
            ),
        },
        "login": stored_secrets.get("login", {}),
        "email": stored_secrets.get("email", {}),
    }

    result = {}
    event = threading.Event()

    class SecretHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            if self.path != "/":
                self.send_response(404)
                self.end_headers()
                return
            html = render_form_html(defaults)
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(html.encode("utf-8"))

        def do_POST(self):
            if self.path != "/submit":
                self.send_response(404)
                self.end_headers()
                return
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8")
            fields = {key: value[0].strip() for key, value in parse_qs(body).items()}
            result.update(
                {
                    "url": fields.get("url", ""),
                    "cookies": {
                        "JSESSIONID": fields.get("cookie_jsessionid", ""),
                        "wengine_new_ticket": fields.get("cookie_wengine", ""),
                        "route": fields.get("cookie_route", ""),
                    },
                    "login": {
                        "username": fields.get("login_username", ""),
                        "password": fields.get("login_password", ""),
                    },
                    "email": {
                        "sender_email": fields.get("sender_email", ""),
                        "sender_password": fields.get("sender_password", ""),
                        "receiver_email": fields.get("receiver_email", ""),
                    },
                }
            )
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write("配置已保存，可以关闭此页面。".encode("utf-8"))
            event.set()

        def log_message(self, format, *args):
            return

    server = HTTPServer(("127.0.0.1", INPUT_PORT), SecretHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    webbrowser.open(f"http://127.0.0.1:{INPUT_PORT}")
    print(f"请在浏览器中完成配置: http://127.0.0.1:{INPUT_PORT}")
    event.wait()
    server.shutdown()
    server.server_close()
    return result


def merge_secrets(base, updates):
    merged = json.loads(json.dumps(base or {}))
    for key in ("url", "cookies", "login", "email"):
        if key not in merged:
            merged[key] = {}
    if updates.get("url"):
        merged["url"] = updates["url"]
    for section in ("cookies", "login", "email"):
        merged_section = merged.get(section, {})
        for field_key, field_value in updates.get(section, {}).items():
            if field_value:
                merged_section[field_key] = field_value
        merged[section] = merged_section
    return merged


def build_cookies(config, secrets):
    override = secrets.get("cookies", {})
    cookies = []
    for cookie in config.get("cookies", []):
        value = pick_value(
            override.get(cookie.get("name", "")), cookie.get("value", "")
        )
        if value:
            cookies.append({**cookie, "value": value})
    return cookies


def build_email_config(config, secrets):
    email_config = dict(config.get("email_config", {}))
    email_override = secrets.get("email", {})
    email_config["sender_email"] = pick_value(email_override.get("sender_email"))
    email_config["sender_password"] = pick_value(email_override.get("sender_password"))
    email_config["receiver_email"] = pick_value(email_override.get("receiver_email"))
    return email_config


def format_component(component):
    parts = [component.get("name", "")]
    if component.get("ratio"):
        parts.append(component["ratio"])
    if component.get("score"):
        parts.append(component["score"])
    return " ".join([part for part in parts if part])


def format_course_details(course):
    lines = [f"· {course['name']} | 总评: {course.get('total', '')}"]
    for component in course.get("components", []):
        lines.append(f"    - {format_component(component)}")
    return "\n".join(lines)


def send_email(changed_courses, email_config):
    required = ["sender_email", "sender_password", "receiver_email"]
    if any(not email_config.get(key) for key in required):
        print("跳过邮件发送：请先在网页中填写邮箱信息。")
        return

    message_text = f"您好，系统于 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 检测到以下科目成绩更新：\n\n"
    message_text += "\n\n".join(
        format_course_details(course) for course in changed_courses
    )
    message_text += "\n\n此邮件由系统自动发送，请勿直接回复。"

    msg = MIMEText(message_text, "plain", "utf-8")
    msg["From"] = email_config["sender_email"]
    msg["To"] = email_config["receiver_email"]
    msg["Subject"] = str(Header("教务系统成绩更新提醒", "utf-8"))

    try:
        server = smtplib.SMTP_SSL(
            email_config["smtp_server"], email_config["smtp_port"]
        )
        server.login(email_config["sender_email"], email_config["sender_password"])
        server.sendmail(
            email_config["sender_email"],
            [email_config["receiver_email"]],
            msg.as_string(),
        )
        server.quit()
        print(f"邮件已成功发送至: {email_config['receiver_email']}")
    except Exception as exc:
        print(f"邮件发送失败 (可能是被拦截): {exc}")


def show_notification(changed_courses):
    message_lines = ["发现成绩更新："]
    for course in changed_courses:
        message_lines.append(f"{course['name']} | 总评: {course.get('total', '')}")
        for component in course.get("components", []):
            message_lines.append(f"  {format_component(component)}")
    message = "\n".join(message_lines)
    ctypes.windll.user32.MessageBoxW(0, message, "新成绩通知", 0x40 | 0x1)


def normalize_text(text):
    return text.strip() if text else ""


def course_changed(previous, current):
    return previous != current


def build_course_snapshot(name, total, components):
    return {
        "name": name,
        "total": total,
        "components": components,
    }


def merge_course_details(snapshot):
    return {
        "total": snapshot.get("total", ""),
        "components": snapshot.get("components", []),
    }


def get_selector(config, key, fallback=""):
    return config.get("xpath", {}).get(key) or fallback


def get_login_selector(config, key, fallback=""):
    return config.get("login", {}).get(key) or fallback


def should_attempt_login(secrets):
    login = secrets.get("login", {})
    return bool(login.get("username") and login.get("password"))


def get_runtime_url(config, secrets):
    return pick_value(secrets.get("url"), config.get("url", ""))


async def attempt_login(page, config, secrets):
    if not should_attempt_login(secrets):
        return
    username_selector = get_login_selector(config, "username_input")
    password_selector = get_login_selector(config, "password_input")
    submit_selector = get_login_selector(config, "submit_button")
    if not username_selector or not password_selector or not submit_selector:
        return
    username_field = await page.query_selector(username_selector)
    password_field = await page.query_selector(password_selector)
    if not username_field or not password_field:
        return
    login = secrets.get("login", {})
    await page.fill(username_selector, login.get("username", ""))
    await page.fill(password_selector, login.get("password", ""))
    await page.click(submit_selector)
    await page.wait_for_load_state("networkidle")


async def fetch_detail_components(page, row, config):
    detail_button = row.locator(get_selector(config, "detail_button"))
    if await detail_button.count() == 0:
        return []
    try:
        await detail_button.first.click()
    except Exception:
        return []

    modal_selector = get_selector(config, "detail_modal")
    modal = page.locator(modal_selector) if modal_selector else page
    try:
        await modal.wait_for(state="visible", timeout=5000)
    except Exception:
        return []

    rows = modal.locator(get_selector(config, "detail_rows"))
    components = []
    count = await rows.count()
    for index in range(count):
        row_item = rows.nth(index)
        name = normalize_text(
            await row_item.locator(
                get_selector(config, "detail_item_cell")
            ).inner_text()
        )
        ratio = normalize_text(
            await row_item.locator(
                get_selector(config, "detail_ratio_cell")
            ).inner_text()
        )
        score = normalize_text(
            await row_item.locator(
                get_selector(config, "detail_score_cell")
            ).inner_text()
        )
        if name or ratio or score:
            components.append({"name": name, "ratio": ratio, "score": score})

    close_button = modal.locator(get_selector(config, "detail_close_button"))
    if await close_button.count() > 0:
        await close_button.first.click()
    return components


async def scrape_courses(page, config):
    row_selector = get_selector(config, "course_row", "tr")
    rows = page.locator(row_selector)
    count = await rows.count()
    courses = []
    for index in range(count):
        row = rows.nth(index)
        name = normalize_text(
            await row.locator(get_selector(config, "course_name_cell")).inner_text()
        )
        if not name:
            continue
        total = normalize_text(
            await row.locator(get_selector(config, "total_score_cell")).inner_text()
        )
        components = await fetch_detail_components(page, row, config)
        courses.append(build_course_snapshot(name, total, components))
    return courses


async def check_grades(context, seen_courses, config, secrets):
    page = await context.new_page()
    url = get_runtime_url(config, secrets)
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 正在检查成绩...")

    try:
        await page.goto(url, wait_until="networkidle")
        await attempt_login(page, config, secrets)

        search_xpath = get_selector(config, "search_button")
        if search_xpath:
            try:
                await page.wait_for_selector(
                    f"xpath={search_xpath}", state="visible", timeout=10000
                )
            except Exception:
                print("未检测到查询按钮，可能需要手动登录，请在浏览器完成登录。")
                await page.wait_for_selector(f"xpath={search_xpath}", timeout=0)
            await page.click(f"xpath={search_xpath}")

        course_selector = get_selector(config, "course_name_cell")
        try:
            await page.wait_for_selector(course_selector, timeout=15000)
        except Exception:
            print("未检测到成绩表格，可能需要手动登录，请在浏览器完成登录。")
            await page.wait_for_selector(course_selector, timeout=0)
        courses = await scrape_courses(page, config)

        current_courses = {}
        changed_courses = []
        for course in courses:
            current_courses[course["name"]] = merge_course_details(course)
            previous = seen_courses.get(course["name"])
            if previous is None or course_changed(
                previous, current_courses[course["name"]]
            ):
                changed_courses.append(course)

        if changed_courses:
            print(f"发现成绩更新: {[course['name'] for course in changed_courses]}")
            send_email(changed_courses, build_email_config(config, secrets))
            show_notification(changed_courses)
            seen_courses.update(current_courses)
            save_seen_courses(seen_courses)
        else:
            print("未发现新成绩。")

        await page.screenshot(path="last_check.png")
    except Exception as exc:
        print(f"检查过程中发生错误: {exc}")
    finally:
        await page.close()


async def run():
    config = load_config()
    stored_secrets = load_user_secrets()
    runtime_secrets = collect_runtime_secrets(config, stored_secrets)
    secrets = merge_secrets(stored_secrets, runtime_secrets)
    save_user_secrets(secrets)

    user_data_dir = config.get("user_data_dir", USER_DATA_DIR)
    user_data_dir = os.path.abspath(user_data_dir)

    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            user_data_dir, headless=False, channel="msedge"
        )
        cookies = build_cookies(config, secrets)
        if cookies:
            await context.add_cookies(cookies)

        seen_courses = load_seen_courses()
        try:
            while True:
                await check_grades(context, seen_courses, config, secrets)
                interval = config.get("check_interval_seconds", 1800)
                print(f"等待 {interval // 60} 分钟后进行下一次检查...")
                await asyncio.sleep(interval)
        except KeyboardInterrupt:
            print("脚本已停止。")
        finally:
            await context.close()


if __name__ == "__main__":
    asyncio.run(run())
