import imaplib
import email
from email.header import decode_header
import json
import os
import re

CONFIG_FILE = "config.json"

class IMAPService:
    def __init__(self):
        self.config = self.load_config()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {
            "email": "",
            "password": "",
            "server": "outlook.office365.com",
            "port": 993
        }

    def save_config(self, email_str, password_str, server_str="outlook.office365.com", port=993):
        self.config = {
            "email": email_str,
            "password": password_str,
            "server": server_str,
            "port": port
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=4)

    def decode_str(self, val):
        if not val:
            return ""
        decoded_list = decode_header(val)
        res = []
        for bytes_or_str, encoding in decoded_list:
            if isinstance(bytes_or_str, bytes):
                try:
                    res.append(bytes_or_str.decode(encoding or "utf-8", errors="ignore"))
                except Exception:
                    res.append(bytes_or_str.decode("latin1", errors="ignore"))
            else:
                res.append(str(bytes_or_str))
        return "".join(res)

    def get_email_body(self, msg):
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    try:
                        payload = part.get_payload(decode=True)
                        charset = part.get_content_charset() or "utf-8"
                        body = payload.decode(charset, errors="ignore")
                        break
                    except Exception:
                        pass
        else:
            try:
                payload = msg.get_payload(decode=True)
                charset = msg.get_content_charset() or "utf-8"
                body = payload.decode(charset, errors="ignore")
            except Exception:
                pass
        return body

    def fetch_alta_emails(self, search_term="alta", max_results=20):
        if not self.config["email"] or not self.config["password"]:
            return {"success": False, "error": "Faltan credenciales en config.json", "emails": []}

        try:
            mail = imaplib.IMAP4_SSL(self.config["server"], self.config["port"])
            mail.login(self.config["email"], self.config["password"])
            mail.select("INBOX")

            # Buscar todos los correos
            status, messages = mail.search(None, 'ALL')
            if status != "OK" or not messages[0]:
                mail.logout()
                return {"success": True, "error": None, "emails": []}

            email_ids = messages[0].split()
            # Ordenar más reciente primero
            email_ids = email_ids[::-1]

            results = []
            for e_id in email_ids:
                if len(results) >= max_results:
                    break

                res, msg_data = mail.fetch(e_id, "(RFC822)")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        subject = self.decode_str(msg.get("Subject", ""))
                        sender = self.decode_str(msg.get("From", ""))
                        date_str = msg.get("Date", "")

                        if not search_term or search_term.lower() in subject.lower():
                            body = self.get_email_body(msg)
                            
                            name_match = re.search(r'alta\s+solicitada\s+(.+)', subject, re.IGNORECASE)
                            extracted_name = name_match.group(1).strip() if name_match else subject

                            results.append({
                                "id": len(results) + 1,
                                "subject": subject,
                                "extracted_name": extracted_name,
                                "sender": sender,
                                "received_time": date_str,
                                "unread": False,
                                "body": body,
                                "_msg_id": e_id.decode('utf-8')
                            })

            mail.logout()
            return {"success": True, "error": None, "emails": results}
        except Exception as e:
            return {"success": False, "error": str(e), "emails": []}

if __name__ == "__main__":
    srv = IMAPService()
    print("Módulo IMAP listo.")
