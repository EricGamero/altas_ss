import re
import win32com.client

class OutlookService:
    def __init__(self):
        self.outlook = None
        self.namespace = None

    def connect(self):
        """Conecta con la aplicación local de Microsoft Outlook."""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            return True
        except Exception as e:
            print(f"Error al conectar con Outlook: {e}")
            return False

    def get_inbox_folders(self):
        """Obtiene las carpetas de Bandeja de Entrada disponibles."""
        if not self.connect():
            return []

        folders = []
        # 1. Carpeta por defecto
        try:
            default_inbox = self.namespace.GetDefaultFolder(6) # 6 = olFolderInbox
            if default_inbox:
                folders.append(default_inbox)
        except Exception:
            pass

        # 2. Otras cuentas / buzones compartidos
        try:
            for store in self.namespace.Folders:
                for subfolder in store.Folders:
                    name_lower = subfolder.Name.lower()
                    if name_lower in ["bandeja de entrada", "inbox", "gestioaltes"]:
                        if subfolder not in folders:
                            folders.append(subfolder)
        except Exception:
            pass

        return folders

    def get_alta_emails(self, search_term="alta", max_results=30):
        """
        Busca correos cuyo asunto contenga el término especificado.
        """
        folders = self.get_inbox_folders()
        if not folders:
            return []

        results = []

        for inbox in folders:
            try:
                items = inbox.Items
                items.Sort("[ReceivedTime]", True) # Ordenar más reciente primero

                # Revisar solo los últimos 100 correos para respuesta ultrarrápida
                idx = 0
                for item in items:
                    idx += 1
                    if idx > 100 or len(results) >= max_results:
                        break

                    try:
                        if getattr(item, "Class", 0) != 43: # MailItem
                            continue

                        subject = item.Subject or ""
                        
                        # Si no hay término o coincide con search_term
                        if not search_term or search_term.lower() in subject.lower():
                            name_match = re.search(r'alta\s+solicitada\s+(.+)', subject, re.IGNORECASE)
                            extracted_name = name_match.group(1).strip() if name_match else subject

                            results.append({
                                "id": len(results) + 1,
                                "folder_name": inbox.Name,
                                "subject": subject,
                                "extracted_name": extracted_name,
                                "sender": getattr(item, "SenderName", "Desconocido"),
                                "received_time": str(getattr(item, "ReceivedTime", "")),
                                "unread": getattr(item, "UnRead", False),
                                "body": getattr(item, "Body", ""),
                                "attachments_count": getattr(getattr(item, "Attachments", None), "Count", 0),
                                "_item": item
                            })
                    except Exception:
                        continue
            except Exception as e:
                print(f"Error accediendo a carpeta {inbox.Name}: {e}")

        return results

if __name__ == "__main__":
    service = OutlookService()
    print("Conectando con Outlook...")
    if service.connect():
        emails = service.get_alta_emails(search_term="alta", max_results=15)
        print(f"--- Correos encontrados ({len(emails)}) ---")
        for m in emails:
            st = "NUEVO" if m["unread"] else "LEÍDO"
            print(f"[{m['id']}] [{st}] Asunto: {m['subject']}")
            print(f"     Persona/Objeto: {m['extracted_name']}")
            print(f"     De: {m['sender']} | Fecha: {m['received_time']}")
            print("-" * 55)
