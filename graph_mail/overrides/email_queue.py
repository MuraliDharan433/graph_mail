import frappe
import requests
import base64
import email as email_lib
from frappe.email.doctype.email_queue.email_queue import EmailQueue


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


def send_via_graph(sender, recipients, msg):
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

    # Use sender from the email itself, fallback to settings
    import re
    sender_email = settings.sender_email
    if sender:
        match = re.search(r'[\w\.-]+@[\w\.-]+', str(sender))
        if match:
            extracted = match.group(0)
            # Only use if it's from our domain
            if "infiligence.com" in extracted:
                sender_email = extracted

    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": body_type, "content": body_content},
            "toRecipients": to_recipients,
            "from": {
                "emailAddress": {"address": sender_email}
            }
        },
        "saveToSentItems": True
    }

    if attachments:
        payload["message"]["attachments"] = attachments

    send_url = f"https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    response = requests.post(send_url, headers=headers, json=payload)

    if response.status_code != 202:
        frappe.log_error(title="Graph Mail Error", message=f"{response.status_code}: {response.text}")
        raise Exception(f"Graph Mail failed: {response.text}")

    frappe.logger().info(f"Graph Mail sent from {sender_email} to {recipients}")


class GraphEmailQueue(EmailQueue):
    def send(self, smtp_server_instance=None):
        try:
            self.db_set("status", "Sending")

            recipients = [r.recipient for r in self.recipients]
            msg = self.message
            sender = self.sender

            send_via_graph(sender, recipients, msg)

            self.db_set("status", "Sent")
            frappe.db.commit()

        except Exception as e:
            frappe.log_error(title="Graph Mail Send Error", message=str(e))
            self.db_set("status", "Error", update_modified=False)
            self.db_set("error", str(e), update_modified=False)
            frappe.db.commit()
            raise
