import smtplib
import uuid
import time
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formatdate, make_msgid
from datetime import datetime

# --- 邮件配置 (请确保填入了正确的账号) ---
EMAIL_CONFIG = {
    "smtp_server": "smtp.163.com",
    "smtp_port": 465,
    "sender_email": "18928617338@163.com",     # <--- 请确保这里是您的网易邮箱
    "sender_password": "DQ2YtxQjCM9J4Kge",    # 授权码
    "receiver_email": "13827330000@163.com"     # <--- 建议先改回您自己的网易邮箱试试
}

def test_send_email():
    print(f"正在尝试连接到 {EMAIL_CONFIG['smtp_server']}...")
    
    if "YOUR_EMAIL" in EMAIL_CONFIG["sender_email"]:
        print("错误: 请先修改脚本中的邮箱地址！")
        return

    # 构建高度伪装的邮件内容
    content = f"""您好：

这是一封系统自动生成的运行状态报告。
当前节点：Node-Windows-{uuid.uuid4().hex[:6]}
状态：Normal
时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

请查收附件相关的系统配置说明（如有）。
祝好。
"""
    
    msg = MIMEText(content, 'plain', 'utf-8')
    
    # --- 核心：添加伪装邮件头 ---
    msg['From'] = EMAIL_CONFIG['sender_email']
    msg['To'] = EMAIL_CONFIG["receiver_email"]
    msg['Subject'] = Header('关于近期系统运行情况的说明', 'utf-8') # 极其正式的主题
    msg['Date'] = formatdate(localtime=True) # 添加标准日期头
    msg['Message-ID'] = make_msgid() # 添加标准消息 ID
    msg['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'

    try:
        print("正在建立安全连接并登录...")
        server = smtplib.SMTP_SSL(EMAIL_CONFIG["smtp_server"], EMAIL_CONFIG["smtp_port"])
        # server.set_debuglevel(1) # 如果还是失败，可以取消此行注释查看详细交互过程
        server.login(EMAIL_CONFIG["sender_email"], EMAIL_CONFIG["sender_password"])
        
        print("正在发送数据包...")
        server.sendmail(EMAIL_CONFIG["sender_email"], [EMAIL_CONFIG["receiver_email"]], msg.as_string())
        server.quit()
        
        print("-" * 30)
        print("恭喜！邮件发送成功。")
        print(f"请检查邮箱: {EMAIL_CONFIG['receiver_email']}")
        print("-" * 30)
    except Exception as e:
        print("-" * 30)
        print(f"测试依然失败，错误信息: {e}")
        print("-" * 30)
        print("最后建议：")
        print("1. 登录网易网页版邮箱，看看是否有一封退信，点击里面的链接申诉。")
        print("2. 尝试将 'receiver_email' 也改为您的网易邮箱（自发自收）。")
        print("3. 如果还是不行，建议换用 QQ 邮箱作为发件人，网易的 SMTP 过滤是目前国内最严的。")

if __name__ == "__main__":
    test_send_email()
