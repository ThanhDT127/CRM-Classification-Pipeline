import logging
import smtplib
from datetime import datetime
from email.message import EmailMessage

import requests

import config
from sharepoint import AuthProvider

logger = logging.getLogger("crm-automation")


def send_alert(subject: str, html_body: str) -> bool:
    """Gui mot thu canh bao KHONG can AuthProvider.

    `NotificationService` doi mot AuthProvider vi duong Graph can token Azure.
    Duong SMTP thi khong can gi ca, va `llm.py` khong co AuthProvider trong tay --
    no duoc dung o tang duoi tang pipeline. Ham nay la cua duy nhat cho no.

    KHONG BAO GIO NEM LOI: bao dong la viec phu, xu ly du lieu la viec chinh.
    """
    try:
        svc = NotificationService.__new__(NotificationService)
        return svc._send_email(subject, html_body)
    except Exception as e:
        logger.error("Khong gui duoc thu canh bao: %s", e)
        return False


class NotificationService:
    """Sends email notifications via Microsoft Graph sendMail endpoint using App Credentials."""
    def __init__(self, auth: AuthProvider) -> None:
        self.auth = auth
        self.session = requests.Session()

    def _send_email_smtp(self, subject: str, html_body: str,
                         sender: str, recipients: list) -> bool:
        """Duong gui THUONG. Dung khi chua co thong tin dang nhap Azure.

        Ton tai canh duong Graph chu khong thay the no: Graph gui bang hom thu cong ty
        va la duong cho production. Do 12/09: ca ba bien AZURE_* deu trong, nen duong
        Graph CHUA TUNG gui duoc thu nao.
        """
        if not config.SMTP_HOST or not config.SMTP_PASSWORD:
            logger.warning("Email skipped: SMTP_HOST/SMTP_PASSWORD is not configured.")
            return False

        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = sender
        msg["To"] = ", ".join(recipients)
        msg.set_content(html_body, subtype="html")

        try:
            with smtplib.SMTP(config.SMTP_HOST, config.SMTP_PORT, timeout=20) as s:
                if config.SMTP_USE_TLS:
                    s.starttls()
                s.login(config.SMTP_USERNAME or sender, config.SMTP_PASSWORD)
                s.send_message(msg)
            logger.info("Email sent via SMTP from %s -> %s", sender, ", ".join(recipients))
            return True
        except Exception as e:
            logger.warning("SMTP send failed: %s", e)
            return False

    def _send_email(self, subject: str, html_body: str) -> bool:
        sender = config.NOTIFICATION_SENDER_EMAIL
        recipients = config.NOTIFICATION_RECIPIENTS

        if not sender:
            logger.warning("Email skipped: NOTIFICATION_SENDER_EMAIL is not configured.")
            return False
        if not recipients:
            logger.warning("Email skipped: NOTIFICATION_RECIPIENTS is not configured.")
            return False

        if config.MAIL_TRANSPORT == "smtp":
            return self._send_email_smtp(subject, html_body, sender, recipients)

        url = f"{config.GRAPH_BASE}/users/{sender}/sendMail"
        to_recipients = [{"emailAddress": {"address": addr}} for addr in recipients]
        
        payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML",
                    "content": html_body
                },
                "toRecipients": to_recipients
            },
            "saveToSentItems": False
        }
        
        try:
            response = self.session.post(
                url,
                headers=self.auth.get_headers(),
                json=payload
            )
            if response.status_code in (200, 202):
                logger.info("Email notification sent successfully from %s -> %s", sender, ", ".join(recipients))
                return True
            else:
                logger.warning("Failed to send email notification (%d): %s", response.status_code, response.text[:300])
        except Exception as e:
            logger.warning("Email notification HTTP request failed: %s", e)
            
        return False

    def send_success(self, duration_s: float, new_rows: int, cells_filled: int) -> None:
        subject = f"[CRM Pipeline] Hoàn tất phân loại dữ liệu CRM - {datetime.now().strftime('%Y-%m-%d')}"
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        html_body = f"""
        <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 600px; line-height: 1.6; color: #333;">
            <div style="background: #d4edda; padding: 16px 20px; border-radius: 6px; margin-bottom: 20px; border-left: 5px solid #28a745;">
                <strong style="color: #155724; font-size: 18px;">✓ Pipeline Phân Loại CRM Thành Công</strong>
            </div>
            <p>Hệ thống tự động chạy ngầm đã hoàn tất phân loại dữ liệu và ghi đè an toàn lên SharePoint.</p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px;">
                <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px 0; color: #666; width: 180px;">Thời điểm chạy:</td><td style="padding: 8px 0; font-weight: bold;">{timestamp}</td></tr>
                <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px 0; color: #666;">Số dòng mới (Delta):</td><td style="padding: 8px 0; font-weight: bold; color: #28a745;">{new_rows} dòng</td></tr>
                <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px 0; color: #666;">Số ô được điền:</td><td style="padding: 8px 0; font-weight: bold;">{cells_filled} ô</td></tr>
                <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px 0; color: #666;">Thời gian xử lý:</td><td style="padding: 8px 0;">{duration_s:.1f} giây</td></tr>
                <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px 0; color: #666;">File SharePoint:</td><td style="padding: 8px 0; font-family: monospace;">{config.SHAREPOINT_TARGET_FILE_PATH}</td></tr>
            </table>
            <p style="color: #888; font-size: 12px; margin-top: 25px; border-top: 1px solid #eee; padding-top: 10px;">
                Đây là email tự động từ hệ thống máy ảo CRM Automation Daemon. Vui lòng không trả lời email này.
            </p>
        </div>
        """
        self._send_email(subject, html_body)

    def send_error(self, error_msg: str) -> None:
        subject = f"[CRITICAL ERROR] Pipeline Phân Loại CRM Thất Bại - {datetime.now().strftime('%Y-%m-%d')}"
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Sanitize error message to prevent rendering bugs
        error_display = error_msg.replace("<", "&lt;").replace(">", "&gt;")
        
        html_body = f"""
        <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 600px; line-height: 1.6; color: #333;">
            <div style="background: #f8d7da; padding: 16px 20px; border-radius: 6px; margin-bottom: 20px; border-left: 5px solid #dc3545;">
                <strong style="color: #721c24; font-size: 18px;">❌ Lỗi Chạy Pipeline Phân Loại CRM</strong>
            </div>
            <p>Hệ thống tự động chạy ngầm gặp sự cố nghiêm trọng và phải dừng chạy khẩn cấp. <b>File Excel trên SharePoint giữ nguyên không bị thay đổi.</b></p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px;">
                <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px 0; color: #666; width: 180px;">Thời điểm lỗi:</td><td style="padding: 8px 0; font-weight: bold;">{timestamp}</td></tr>
                <tr style="border-bottom: 1px solid #ddd;"><td style="padding: 8px 0; color: #666;">Tệp mục tiêu:</td><td style="padding: 8px 0; font-family: monospace;">{config.SHAREPOINT_TARGET_FILE_PATH}</td></tr>
            </table>
            <div style="background: #f8f9fa; padding: 15px; border-radius: 4px; border: 1px solid #e9ecef; font-family: monospace; font-size: 13px; color: #c00; overflow-x: auto; white-space: pre-wrap;">
{error_display}
            </div>
            <p style="color: #888; font-size: 12px; margin-top: 25px; border-top: 1px solid #eee; padding-top: 10px;">
                Vui lòng kiểm tra lại log hệ thống trên máy ảo để khắc phục sự cố.
            </p>
        </div>
        """
        self._send_email(subject, html_body)
