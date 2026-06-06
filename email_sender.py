"""
email_sender.py — Gmail SMTP으로 리포트 발송
"""
from __future__ import annotations

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import date

from config import GMAIL_USER, GMAIL_APP_PASSWORD


def send_report_email(
    to_address: str,
    report_text: str,
    gmail_user: str = "",
    gmail_password: str = "",
) -> tuple[bool, str]:
    """
    Gmail SMTP으로 리포트 이메일 발송.
    반환: (성공여부, 메시지)
    """
    sender = gmail_user or GMAIL_USER
    password = gmail_password or GMAIL_APP_PASSWORD

    if not sender or not password:
        return False, "Gmail 계정 정보(이메일/앱 비밀번호)가 설정되지 않았습니다."
    if not to_address or "@" not in to_address:
        return False, "유효하지 않은 수신 이메일 주소입니다."

    today = date.today().isoformat()
    subject = f"📊 주식 투자 브리핑 리포트 [{today}]"

    # HTML 버전 (간단한 포매팅)
    html_body = report_text.replace("\n", "<br>").replace("=", "─")
    html_content = f"""
    <html>
    <body style="font-family: 'Malgun Gothic', Arial, sans-serif; line-height: 1.8; color: #1a1a2e; padding: 20px;">
        <div style="max-width: 800px; margin: 0 auto; background: #f8f9fa; border-radius: 12px; padding: 30px;">
            <h1 style="color: #0f3460; border-bottom: 3px solid #e94560; padding-bottom: 10px;">
                📊 주식 투자 브리핑 리포트
            </h1>
            <p style="color: #666; font-size: 14px;">기준일: {today}</p>
            <div style="background: white; border-radius: 8px; padding: 20px; margin-top: 20px; 
                        box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                <pre style="white-space: pre-wrap; font-family: 'Malgun Gothic', Arial, sans-serif; 
                            font-size: 14px; line-height: 1.8;">{report_text}</pre>
            </div>
            <p style="margin-top: 20px; font-size: 12px; color: #999; text-align: center;">
                본 리포트는 AI 분석 기반이며 투자 결정은 본인 책임입니다.
            </p>
        </div>
    </body>
    </html>
    """

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = sender
        msg["To"] = to_address

        msg.attach(MIMEText(report_text, "plain", "utf-8"))
        msg.attach(MIMEText(html_content, "html", "utf-8"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, to_address, msg.as_string())

        return True, f"✅ 리포트가 {to_address}로 성공적으로 발송되었습니다."
    except smtplib.SMTPAuthenticationError:
        return False, "❌ Gmail 인증 실패. 앱 비밀번호를 확인해주세요. (2단계 인증 필요)"
    except smtplib.SMTPException as e:
        return False, f"❌ 이메일 발송 실패: {str(e)}"
    except Exception as e:
        return False, f"❌ 예상치 못한 오류: {str(e)}"
