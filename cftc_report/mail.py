"""邮件发送"""

import smtplib
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from .config import (
    EMAIL_FROM, EMAIL_PASSWORD, EMAIL_TO,
    OUTPUT_EXCEL, OUTPUT_FOLDER, SMTP_PORT, SMTP_SERVER,
)


def send_email(current_analysis=""):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_FROM
    msg['To'] = EMAIL_TO
    msg['Subject'] = f'CFTC 数据更新 {datetime.now().strftime("%Y-%m-%d")}'
    body = f"""CFTC 数据更新完成。\n\n=== 当前分析 ===\n{current_analysis}\n祝交易顺利！"""
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    if OUTPUT_EXCEL.exists():
        with open(OUTPUT_EXCEL, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={OUTPUT_EXCEL.name}")
            msg.attach(part)
    for png in OUTPUT_FOLDER.glob("*.png"):
        with open(png, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={png.name}")
            msg.attach(part)

    try:
        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print("邮件已发送")
    except Exception as e:
        print(f"邮件发送失败: {e}")
