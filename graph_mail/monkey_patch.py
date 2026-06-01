import frappe
import requests
import base64
import email as email_lib


def get_graph_token():
    settings = frappe.get_single("Graph Mail Settings")
    url = f"https://login.microsoftonline.com/{settings.tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": settings.client_id,
        "client_secret": settings.get_password("client_secret"),
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(url, data=data)
    response.raise_for_status()
    return response.json()["access_token"]


def graph_sendmail(self, sender, recipients, msg):
    settings = frappe.get_single("Graph Mail Settings")
    token = get_graph_token()

    if isinstance(recipients, str):
        recipients = [recipients]

    to_recipients = [{"emailAddress": {"address": r.strip()}} for r in recipients]

    if isinstance(msg, bytes):
        parsed = email_lib.message_from_bytes(msg)
    else:
        parsed = email_lib.message_from_string(str(msg))

    subject = parsed.get("Subject", "(no subject)")
    html_body = ""
    text_body = ""
    attachments = []

    if parsed.is_multipart():
        for part in parsed.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if ct == "text/html" and "attachment" not in cd:
                html_body = part.get_payload(decode=True).decode("utf-8", errors="replace")
            elif ct == "text/plain" and "attachment" not in cd:
                text_body = part.get_payload(decode=True).decode("utf-8", errors="replace")
            elif "attachment" in cd:
                attachments.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": part.get_filename() or "attachment",
                    "contentBytes": base64.b64encode(
                        part.get_payload(decode=True)
                    ).decode("utf-8"),
                    "contentType": ct
                })
    else:
        content = parsed.get_payload(decode=True)
        if content:
            text_body = content.decode("utf-8", errors="replace")

    body_content = html_body if html_body else text_body
    body_type = "HTML" if html_body else "Text"

    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": body_type, "content": body_content},
            "toRecipients": to_recipients,
        },
        "saveToSentItems": True
    }

    if attachments:
        payload["message"]["attachments"] = attachments

    send_url = f"https://graph.microsoft.com/v1.0/users/{settings.sender_email}/sendMail"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    response = requests.post(send_url, headers=headers, json=payload)

    if response.status_code != 202:
        frappe.log_error(title="Graph Mail Error", message=f"{response.status_code}: {response.text}")
        raise Exception(f"Graph Mail failed: {response.text}")


def apply_patch():
    from frappe.email import smtp as frappe_smtp
    import smtplib

    class FakeSMTP:
        def __init__(self, *args, **kwargs):
            pass
        def sendmail(self, sender, recipients, msg):
            graph_sendmail(None, sender, recipients, msg)
        def quit(self):
            pass
        def login(self, *args, **kwargs):
            pass
        def starttls(self, *args, **kwargs):
            pass

    # Patch smtplib.SMTP itself so no real connection is ever made
    smtplib.SMTP = FakeSMTP
    smtplib.SMTP_SSL = FakeSMTP
