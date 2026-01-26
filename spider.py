import asyncio
import base64
import ctypes
import json
import os
import re
import smtplib
import threading
import urllib.request
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
    email = defaults.get("email", {})
    login = defaults.get("login", {})
    ocr = defaults.get("ocr", {})
    login_url = defaults.get("login_url", "")
    grades_url = defaults.get("grades_url", "")
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
        <label>登录界面 URL</label>
        <input name="login_url" value="{login_url}" placeholder="统一认证/登录入口 URL" />
        <label>成绩查询 URL</label>
        <input name="grades_url" value="{grades_url}" placeholder="成绩查询页面 URL" />

        <label>学号/账号</label>
        <input name="login_username" value="{login.get("username", "")}" />
        <label>登录密码</label>
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
      <fieldset>
        <legend>验证码识别（OCR）</legend>
        <label>base_url</label>
        <input name="ocr_base_url" value="{ocr.get("base_url", "")}" placeholder="https://api.openai.com/v1" />
        <label>model</label>
        <input name="ocr_model" value="{ocr.get("model", "")}" placeholder="gpt-4o-mini" />
        <label>api_key</label>
        <input name="ocr_api_key" type="password" value="{ocr.get("api_key", "")}" />
        <div class="hint">要求 OpenAI 兼容接口（/v1/chat/completions）。</div>
      </fieldset>
      <button type="submit">保存并开始监控</button>
    </form>
  </div>
</body>
</html>"""


def collect_runtime_secrets(config, stored_secrets):
    legacy_url = pick_value(stored_secrets.get("url"), config.get("url", ""))
    defaults = {
        "login_url": pick_value(
            stored_secrets.get("login_url"), config.get("login_url", ""), legacy_url
        ),
        "grades_url": pick_value(
            stored_secrets.get("grades_url"), config.get("grades_url", ""), legacy_url
        ),
        "login": stored_secrets.get("login", {}),
        "email": stored_secrets.get("email", {}),
        "ocr": {
            "base_url": pick_value(
                stored_secrets.get("ocr", {}).get("base_url"),
                config.get("ocr", {}).get("base_url", ""),
            ),
            "model": pick_value(
                stored_secrets.get("ocr", {}).get("model"),
                config.get("ocr", {}).get("model", ""),
            ),
            "api_key": stored_secrets.get("ocr", {}).get("api_key", ""),
        },
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
                    "login_url": fields.get("login_url", ""),
                    "grades_url": fields.get("grades_url", ""),
                    "login": {
                        "username": fields.get("login_username", ""),
                        "password": fields.get("login_password", ""),
                    },
                    "email": {
                        "sender_email": fields.get("sender_email", ""),
                        "sender_password": fields.get("sender_password", ""),
                        "receiver_email": fields.get("receiver_email", ""),
                    },
                    "ocr": {
                        "base_url": fields.get("ocr_base_url", ""),
                        "model": fields.get("ocr_model", ""),
                        "api_key": fields.get("ocr_api_key", ""),
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
    for key in (
        "url",
        "login_url",
        "grades_url",
        "login",
        "email",
        "ocr",
    ):
        if key not in merged:
            merged[key] = {}
    if updates.get("url"):
        merged["url"] = updates["url"]
    if updates.get("login_url"):
        merged["login_url"] = updates["login_url"]
    if updates.get("grades_url"):
        merged["grades_url"] = updates["grades_url"]
    for section in ("login", "email", "ocr"):
        merged_section = merged.get(section, {})
        for field_key, field_value in updates.get(section, {}).items():
            if field_value:
                merged_section[field_key] = field_value
        merged[section] = merged_section
    return merged


def build_email_config(config, secrets):
    email_config = dict(config.get("email_config", {}))
    email_override = secrets.get("email", {})
    email_config["sender_email"] = pick_value(email_override.get("sender_email"))
    email_config["sender_password"] = pick_value(email_override.get("sender_password"))
    email_config["receiver_email"] = pick_value(email_override.get("receiver_email"))
    return email_config


def build_ocr_config(config, secrets):
    config_ocr = config.get("ocr", {})
    secrets_ocr = secrets.get("ocr", {})
    return {
        "base_url": pick_value(
            secrets_ocr.get("base_url"), config_ocr.get("base_url", "")
        ),
        "model": pick_value(secrets_ocr.get("model"), config_ocr.get("model", "")),
        "api_key": pick_value(secrets_ocr.get("api_key")),
        "timeout_seconds": config_ocr.get("timeout_seconds", 30),
        "max_retries": config_ocr.get("max_retries", 3),
    }


def build_openai_endpoint(base_url):
    if not base_url:
        return ""
    base = base_url.rstrip("/")
    if base.endswith("/v1"):
        return f"{base}/chat/completions"
    return f"{base}/v1/chat/completions"


def is_ocr_configured(ocr_config):
    return bool(
        ocr_config.get("base_url")
        and ocr_config.get("model")
        and ocr_config.get("api_key")
    )


def request_ocr_text(ocr_config, image_base64):
    if not is_ocr_configured(ocr_config):
        return ""
    endpoint = build_openai_endpoint(ocr_config["base_url"])
    if not endpoint:
        return ""
    prompt = "请识别图片中的算式，只输出算式，例如 12+8，不要输出其他文字。"
    payload = {
        "model": ocr_config["model"],
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/png;base64,{image_base64}"},
                    },
                ],
            }
        ],
        "temperature": 0,
    }
    data = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        endpoint,
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {ocr_config['api_key']}",
        },
    )
    try:
        with urllib.request.urlopen(
            request, timeout=ocr_config["timeout_seconds"]
        ) as response:
            response_text = response.read().decode("utf-8")
    except Exception as exc:
        print(f"OCR 请求失败: {exc}")
        return ""
    try:
        data = json.loads(response_text)
        return data.get("choices", [{}])[0].get("message", {}).get("content", "")
    except Exception as exc:
        print(f"OCR 响应解析失败: {exc}")
        return ""


def normalize_ocr_text(text):
    return (
        text.replace(" ", "")
        .replace("\n", "")
        .replace("×", "*")
        .replace("x", "*")
        .replace("X", "*")
        .replace("÷", "/")
    )


def format_math_result(value):
    if value is None:
        return ""
    if isinstance(value, float):
        if abs(value - round(value)) < 1e-9:
            return str(int(round(value)))
        return str(round(value, 4)).rstrip("0").rstrip(".")
    return str(value)


def solve_math_from_text(text):
    if not text:
        return ""
    normalized = normalize_ocr_text(text)
    match = re.search(r"(\d+)([+\-*/])(\d+)", normalized)
    if match:
        left = int(match.group(1))
        op = match.group(2)
        right = int(match.group(3))
        if op == "+":
            return format_math_result(left + right)
        if op == "-":
            return format_math_result(left - right)
        if op == "*":
            return format_math_result(left * right)
        if op == "/":
            if right == 0:
                return ""
            return format_math_result(left / right)
    if re.fullmatch(r"\d+(\.\d+)?", normalized):
        return normalized
    return ""


async def call_ocr_text(ocr_config, image_base64):
    return await asyncio.to_thread(request_ocr_text, ocr_config, image_base64)


async def extract_captcha_base64(container, image_selector, fallback_selector):
    selector = image_selector or ""
    locator = container.locator(selector) if selector else container.locator(fallback_selector)
    if await locator.count() == 0 and fallback_selector:
        locator = container.locator(fallback_selector)
    if await locator.count() == 0:
        return ""
    element = locator.first
    src = await element.get_attribute("src")
    if src and src.startswith("data:image"):
        return src.split(",", 1)[1]
    try:
        data = await element.screenshot(type="png")
        return base64.b64encode(data).decode("utf-8")
    except Exception as exc:
        print(f"验证码截图失败: {exc}")
        return ""


async def refresh_captcha(container, refresh_selector, image_selector, fallback_selector):
    if refresh_selector:
        refresh = container.locator(refresh_selector)
        if await refresh.count() > 0:
            await refresh.first.click()
            return
    selector = image_selector or fallback_selector
    if selector:
        image = container.locator(selector)
        if await image.count() > 0:
            await image.first.click()


async def solve_captcha(container, config, secrets, image_selector, fallback_selector):
    ocr_config = build_ocr_config(config, secrets)
    if not is_ocr_configured(ocr_config):
        print("OCR 配置不完整，无法自动识别验证码。")
        return ""
    image_base64 = await extract_captcha_base64(container, image_selector, fallback_selector)
    if not image_base64:
        return ""
    ocr_text = await call_ocr_text(ocr_config, image_base64)
    answer = solve_math_from_text(ocr_text)
    if not answer:
        print(f"OCR 未能解析验证码算式: {ocr_text}")
    return answer


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


def get_runtime_urls(config, secrets):
    legacy_url = pick_value(secrets.get("url"), config.get("url", ""))
    login_url = pick_value(
        secrets.get("login_url"), config.get("login_url", ""), legacy_url
    )
    grades_url = pick_value(
        secrets.get("grades_url"), config.get("grades_url", ""), legacy_url
    )
    return login_url, grades_url


LOGIN_OK = "ok"
LOGIN_MANUAL = "manual"
LOGIN_FAILED = "failed"


async def wait_for_login_success(page, config, timeout=10000):
    selectors = []
    search_xpath = get_selector(config, "search_button")
    if search_xpath:
        selectors.append(f"xpath={search_xpath}")
    course_selector = get_selector(config, "course_name_cell")
    if course_selector:
        selectors.append(course_selector)
    for selector in selectors:
        try:
            await page.wait_for_selector(selector, timeout=timeout)
            return True
        except Exception:
            continue
    return False


async def wait_for_login_success_forever(page, config):
    selectors = []
    search_xpath = get_selector(config, "search_button")
    if search_xpath:
        selectors.append(f"xpath={search_xpath}")
    course_selector = get_selector(config, "course_name_cell")
    if course_selector:
        selectors.append(course_selector)
    if not selectors:
        return False
    tasks = [
        asyncio.create_task(page.wait_for_selector(sel, timeout=0)) for sel in selectors
    ]
    try:
        done, pending = await asyncio.wait(tasks, return_when=asyncio.FIRST_COMPLETED)
        for task in pending:
            task.cancel()
        return len(done) > 0
    except Exception as exc:
        print(f"等待登录完成时发生错误: {exc}")
        return False


async def is_login_form_visible(page, config):
    target = await get_login_target(page, config)
    return target is not None


async def get_login_target(page, config):
    selectors = [
        get_login_selector(config, "switch_to_password"),
        get_login_selector(config, "username_input"),
        get_login_selector(config, "password_input"),
    ]
    
    # 增加切换账号按钮的选择器
    switch_account_sel = get_login_selector(config, "switch_account_btn")
    if switch_account_sel:
        if switch_account_sel.startswith("/"):
            selectors.append(f"xpath={switch_account_sel}")
        else:
            selectors.append(switch_account_sel)

    # 检查主页面
    for selector in selectors:
        if not selector:
            continue
        try:
            locator = page.locator(selector)
            if await locator.count() > 0 and await locator.first.is_visible():
                return page
        except Exception:
            continue

    # 检查所有 iframe
    for frame in page.frames:
        if frame == page.main_frame:
            continue
        for selector in selectors:
            if not selector:
                continue
            try:
                locator = frame.locator(selector)
                if await locator.count() > 0 and await locator.first.is_visible():
                    return frame
            except Exception:
                continue
    return None


async def check_and_handle_cas_jump(page):
    """检测并处理 CAS 统一身份认证跳转"""
    try:
        # 增加一个小的延迟，确保 frame 加载
        await asyncio.sleep(0.5)
        
        # 检查所有 frame 中的内容
        all_frames = page.frames
        cas_found = False
        target_frame = None
        
        for frame in all_frames:
            try:
                content = await frame.content()
                if "CAS统一身份认证登录" in content or "应用认证平台" in content:
                    cas_found = True
                    target_frame = frame
                    break
            except Exception:
                continue

        if cas_found:
            current_url = page.url
            print(f"检测到 CAS 状态 (URL: {current_url})")
            
            # 如果当前 URL 已经是登录成功后的 URL 或者是教务系统主页，就不再跳转
            if "jwglxt.gpnu.edu.cn" in current_url and "cas_login=true" not in current_url:
                content = await page.content()
                if "广东技术师范大学教务系统" in content:
                    print("已经在教务系统主页，无需再次跳转。")
                    return "SUCCESS"

            print("正在执行 CAS 跳转授权...")
            
            # 记录跳转前的状态
            auth_url = "https://webauth.gpnu.edu.cn/wengine-auth/login?cas_login=true"
            
            try:
                # 尝试跳转，如果已经在跳转中，goto 可能会抛出错误，这里捕获它
                await page.goto(auth_url, wait_until="networkidle", timeout=15000)
            except Exception as e:
                if "navigation" in str(e).lower():
                    print(f"跳转过程中检测到并发导航: {e}，尝试等待加载完成")
                    try:
                        await page.wait_for_load_state("networkidle", timeout=10000)
                    except:
                        pass
                else:
                    print(f"跳转 CAS 授权页面失败: {e}，尝试 domcontentloaded 模式")
                    try:
                        await page.goto(auth_url, wait_until="domcontentloaded", timeout=10000)
                    except Exception as e2:
                        print(f"二次尝试跳转也失败: {e2}")

            # 轮询检查是否到达成功页面，最多等待 10 秒
            print("正在验证登录成功状态...")
            for _ in range(20): # 0.5s * 20 = 10s
                await asyncio.sleep(0.5)
                try:
                    content = await page.content()
                    if "广东技术师范大学教务系统" in content:
                        print("检测到“广东技术师范大学教务系统”文本，登录成功！")
                        return "SUCCESS"
                    
                    # 检查是否还在 CAS 页面
                    all_frames = page.frames
                    still_on_cas = False
                    for frame in all_frames:
                        try:
                            f_content = await frame.content()
                            if "CAS统一身份认证登录" in f_content or "应用认证平台" in f_content:
                                still_on_cas = True
                                break
                        except:
                            continue
                    
                    if not still_on_cas and "jwglxt.gpnu.edu.cn" in page.url:
                        # 虽然没看到文本，但 URL 对了，可能还在加载
                        continue
                except:
                    continue
            
            return True # 返回 True 表示处理过 CAS，但没确认最终成功，让外层重试
    except Exception as exc:
        print(f"检查 CAS 文本或跳转时出错: {exc}")
    return False


async def wait_for_login_exit(page, config, timeout=15000):
    deadline = asyncio.get_running_loop().time() + timeout / 1000
    while True:
        if not await is_login_form_visible(page, config):
            return True
        if asyncio.get_running_loop().time() >= deadline:
            return False
        await page.wait_for_timeout(200)


async def wait_for_login_exit_forever(page, config):
    while True:
        if not await is_login_form_visible(page, config):
            return True
        await page.wait_for_timeout(200)


async def wait_for_login_form_ready(page, config, timeout=2000):
    """等待登录表单出现，支持检测 iframe"""
    deadline = asyncio.get_running_loop().time() + timeout / 1000
    while True:
        if await is_login_form_visible(page, config):
            return True
        if asyncio.get_running_loop().time() >= deadline:
            return False
        await asyncio.sleep(0.2)


async def attempt_login(page, config, secrets):
    # 获取包含登录表单的容器（可能是 page 或某一个 frame）
    target = await get_login_target(page, config)
    if not target:
        return LOGIN_FAILED

    # 优先尝试点击“切换账号登录”
    switch_account_sel = get_login_selector(config, "switch_account_btn")
    if switch_account_sel:
        try:
            selector = f"xpath={switch_account_sel}" if switch_account_sel.startswith("/") else switch_account_sel
            switch_account_btn = target.locator(selector)
            if await switch_account_btn.count() > 0:
                # 检查文本是否匹配，确保是“切换账号登录”
                text = await switch_account_btn.inner_text()
                if "切换账号登录" in text:
                    print("检测到“切换账号登录”按钮，正在点击...")
                    await switch_account_btn.click()
                    await asyncio.sleep(1)  # 等待切换动画或加载
                    # 点击后可能 target 发生了变化，重新获取
                    new_target = await get_login_target(page, config)
                    if new_target:
                        target = new_target
        except Exception as exc:
            print(f"尝试点击“切换账号登录”失败: {exc}")

    username_selector = get_login_selector(config, "username_input")
    password_selector = get_login_selector(config, "password_input")
    submit_selector = get_login_selector(config, "submit_button")
    switch_selector = get_login_selector(config, "switch_to_password")
    
    if not username_selector or not password_selector:
        return LOGIN_FAILED

    # 在 target 中查找元素
    username_field = await target.query_selector(username_selector)
    password_field = await target.query_selector(password_selector)

    if (username_field is None or password_field is None) and switch_selector:
        switch_button = target.locator(switch_selector)
        if await switch_button.count() > 0:
            await switch_button.first.click()
            await asyncio.sleep(0.5)
            # 再次检查
            username_field = await target.query_selector(username_selector)
            password_field = await target.query_selector(password_selector)

    username_visible = False
    password_visible = False
    if username_field:
        try:
            username_visible = await target.locator(username_selector).first.is_visible()
        except Exception:
            username_visible = False
    if password_field:
        try:
            password_visible = await target.locator(password_selector).first.is_visible()
        except Exception:
            password_visible = False

    if not should_attempt_login(secrets):
        if username_visible and password_visible:
            return LOGIN_MANUAL
        return LOGIN_OK

    if not username_visible or not password_visible:
        return LOGIN_MANUAL

    captcha_input_selector = get_login_selector(config, "captcha_input")
    captcha_image_selector = get_login_selector(config, "captcha_image")
    captcha_fallback_selector = get_login_selector(config, "captcha_image_fallback")
    captcha_refresh_selector = get_login_selector(config, "captcha_refresh")

    login = secrets.get("login", {})
    ocr_config = build_ocr_config(config, secrets)
    max_retries = max(1, int(ocr_config.get("max_retries", 3)))

    for attempt in range(max_retries):
        await target.fill(username_selector, login.get("username", ""))
        await target.fill(password_selector, login.get("password", ""))

        captcha_required = False
        if captcha_input_selector:
            captcha_input = target.locator(captcha_input_selector)
            if (
                await captcha_input.count() > 0
                and await captcha_input.first.is_visible()
            ):
                captcha_required = True
                captcha_answer = await solve_captcha(
                    target, # solve_captcha 也需要支持 target (page 或 frame)
                    config,
                    secrets,
                    captcha_image_selector,
                    captcha_fallback_selector,
                )
                if captcha_answer:
                    await captcha_input.first.fill(captcha_answer)
                else:
                    return LOGIN_MANUAL

        if captcha_required:
            await target.locator(captcha_input_selector).first.press("Enter")
        else:
            await target.locator(password_selector).first.press("Enter")
        
        # 统一认证可能需要点跳转，等待一下
        await asyncio.sleep(2)

        # 检测是否有 "CAS统一身份认证登录" 文本并处理跳转
        cas_status = await check_and_handle_cas_jump(page)
        if cas_status == "SUCCESS":
            return LOGIN_OK
        
        if await wait_for_login_success(page, config, timeout=5000):
            return LOGIN_OK
        
        if not captcha_required:
            # 如果没有验证码还失败了，可能是账号密码错，或者需要第二轮登录
            # 我们返回 FAILED 让外层循环处理
            return LOGIN_FAILED
            
        await refresh_captcha(
            target,
            captcha_refresh_selector,
            captcha_image_selector,
            captcha_fallback_selector,
        )

    print("验证码识别失败，请手动输入。")
    return LOGIN_MANUAL


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
    login_url, grades_url = get_runtime_urls(config, secrets)
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 正在检查成绩...")

    # 确保成绩查询 URL 正确
    target_grades_url = grades_url
    
    try:
        await page.goto(login_url, wait_until="domcontentloaded")

        # 增加循环处理逻辑，支持多次账号密码登录（应对多次跳转至登录页的情况）
        max_login_rounds = 5
        for round in range(max_login_rounds):
            # 每一轮开始前先检查是否出现了 CAS 提示界面
            cas_status = await check_and_handle_cas_jump(page)
            if cas_status == "SUCCESS":
                # 如果检测到教务系统文本，说明登录成功，直接退出循环
                break
            elif cas_status:
                # 如果发生了跳转，给一点时间让新页面加载，然后重新开始本轮检测
                await asyncio.sleep(1)

            # 检查当前页面是否已经出现了教务系统文本（可能是不经过 CAS 跳转直接进入的情况）
            if "广东技术师范大学教务系统" in (await page.content()):
                print("检测到“广东技术师范大学教务系统”文本，登录成功！")
                break

            await wait_for_login_form_ready(page, config, timeout=1000)
            login_form_visible = await is_login_form_visible(page, config)

            if login_form_visible:
                print(f"检测到登录界面 (第 {round + 1} 轮)，正在执行登录...")
                login_result = await attempt_login(page, config, secrets)
                if login_result == LOGIN_MANUAL:
                    print("需要手动干预，脚本将等待登录成功后继续。")
                    break
                # 如果是 LOGIN_OK，说明已经到达成功页面，下一轮循环会通过 wait_for_login_success 退出
                # 如果是 LOGIN_FAILED，说明可能还没到成功页面，也可能出现了新的登录框，继续下一轮检测
            else:
                # 检查是否已经处于成功状态
                logged_in = await wait_for_login_success(page, config, timeout=2000)
                if logged_in:
                    print("登录成功，进入主界面。")
                    break
                
                # 如果既没看到登录框也没看到成功标志，且是第一轮，尝试再次检测登录框
                if round == 0:
                    continue
                break

        if not await wait_for_login_exit(page, config, timeout=15000):
            await wait_for_login_exit_forever(page, config)

        # 等待页面加载完成后再跳转
        try:
            await page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            pass

        # 登录流程结束，开始成绩查询部分
        print(f"正在转到成绩查询页面: {target_grades_url}")
        await page.goto(target_grades_url, wait_until="domcontentloaded")
        
        # 再次检查是否需要登录（有时跳转到成绩页会重新要求认证）
        cas_status = await check_and_handle_cas_jump(page)
        if cas_status == "SUCCESS":
            # 已经确认登录成功，重新加载成绩页以防万一
            await page.goto(target_grades_url, wait_until="domcontentloaded")
        elif cas_status:
            await asyncio.sleep(1)
        
        if await is_login_form_visible(page, config):
            login_result = await attempt_login(page, config, secrets)
            if login_result != LOGIN_OK:
                print("需要手动完成登录，脚本将等待成功后继续。")

            if not await wait_for_login_exit(page, config, timeout=15000):
                await wait_for_login_exit_forever(page, config)

            # 重新进入成绩页
            await page.goto(target_grades_url, wait_until="domcontentloaded")

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

        seen_courses = load_seen_courses()
        try:
            while True:
                await check_grades(context, seen_courses, config, secrets)
                interval = config.get("check_interval_seconds", 1800)
                print(f"等待 {interval // 60} 分钟后进行下一次检查...")
                await asyncio.sleep(interval)
        except KeyboardInterrupt:
            print("脚本已停止。")#
        finally:
            await context.close()


if __name__ == "__main__":
    asyncio.run(run())
