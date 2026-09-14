import imaplib
import email
from email.header import decode_header

def check_imap(email_user, password, imap_server="outlook.office365.com"):
    try:
        print(f"Conectando a {imap_server} para {email_user}...")
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_user, password)
        mail.select("inbox")
        
        status, messages = mail.search(None, 'ALL')
        email_ids = messages[0].split()
        print(f"Total correos en servidor: {len(email_ids)}")
        
        # Obtener los 5 más recientes
        for e_id in email_ids[-5:]:
            res, msg_data = mail.fetch(e_id, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding or "utf-8", errors="ignore")
                    print(f"  - Asunto: {subject}")
        mail.logout()
        return True
    except Exception as e:
        print(f"Error IMAP: {e}")
        return False

if __name__ == "__main__":
    print("Prueba de módulo IMAP preparada.")
