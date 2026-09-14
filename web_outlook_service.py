import os
import re
import time
from playwright.sync_api import sync_playwright

USER_DATA_DIR = os.path.join(os.getcwd(), "browser_session")

class WebOutlookService:
    def __init__(self, headless=False):
        self.headless = headless

    def get_alta_emails(self, max_results=10):
        results = []
        os.makedirs(USER_DATA_DIR, exist_ok=True)

        with sync_playwright() as p:
            print("[+] Abriendo navegador para acceder a Outlook Web...")
            # Usamos context con estado persistente para mantener la sesión
            context = p.chromium.launch_persistent_context(
                USER_DATA_DIR,
                headless=self.headless,
                channel="chrome" if os.path.exists("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe") else None,
                args=["--start-maximized", "--disable-blink-features=AutomationControlled"]
            )
            
            page = context.pages[0] if context.pages else context.new_page()
            
            print("[+] Navegando a https://outlook.office.com/mail/ ...")
            page.goto("https://outlook.office.com/mail/", wait_until="domcontentloaded")
            
            # Esperar a que cargue la bandeja o pida login
            time.sleep(3)

            # Comprobar si pide iniciar sesión
            if "login.microsoftonline.com" in page.url or "login.live.com" in page.url:
                print("\n" + "!"*60)
                print("⚠️ POR FAVOR, INICIA SESIÓN EN LA VENTANA DEL NAVEGADOR QUE SE HA ABIERTO.")
                print("El navegador guardará tu sesión para que no tengas que volver a hacerlo.")
                print("!"*60 + "\n")
                
                # Esperar hasta que vuelva a la bandeja de entrada de Outlook
                while "outlook.office" not in page.url and "outlook.live" not in page.url:
                    time.sleep(2)
                
                print("[OK] Sesión detectada correctamente. Cargando correos...")
                time.sleep(4)

            # Esperar a que los correos estén en pantalla
            try:
                page.wait_for_selector('div[role="listbox"], div[role="option"]', timeout=15000)
            except Exception:
                pass

            # Extraer elementos de la lista de correos
            mail_items = page.query_selector_all('div[role="option"]')
            print(f"[+] Elementos visuales detectados en la bandeja: {len(mail_items)}")

            for idx, item in enumerate(mail_items[:max_results], 1):
                try:
                    text_content = item.inner_text()
                    lines = [line.strip() for line in text_content.splitlines() if line.strip()]
                    
                    # Buscar asunto en el texto del elemento
                    subject = lines[0] if lines else "Sin asunto"
                    for line in lines:
                        if "alta" in line.lower() or "solicitada" in line.lower():
                            subject = line
                            break

                    name_match = re.search(r'alta\s+solicitada\s+(.+)', subject, re.IGNORECASE)
                    extracted_name = name_match.group(1).strip() if name_match else subject

                    # Al hacer clic en el elemento, se carga su cuerpo en el panel derecho
                    item.click()
                    time.sleep(1)

                    # Obtener cuerpo del correo seleccionado
                    body_el = page.query_selector('div[role="main"]') or page.query_selector('div[aria-label="Cuerpo del mensaje"]')
                    body_text = body_el.inner_text() if body_el else text_content

                    results.append({
                        "id": idx,
                        "subject": subject,
                        "extracted_name": extracted_name,
                        "sender": lines[0] if lines else "",
                        "received_time": "Hoy / Reciente",
                        "unread": "NUEVO" in text_content.upper(),
                        "body": body_text
                    })
                except Exception as e:
                    continue

            context.close()

        return results

if __name__ == "__main__":
    srv = WebOutlookService(headless=False)
    mails = srv.get_alta_emails(max_results=5)
    print(f"\nResultados obtenidos: {len(mails)}")
    for m in mails:
        print(f" - [{m['id']}] {m['subject']}")
