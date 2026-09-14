from outlook_service import OutlookService

service = OutlookService()
if service.connect():
    emails = service.get_alta_emails(search_term="alta", max_results=5)
    for idx, mail in enumerate(emails, 1):
        print(f"=== CORREO {idx}: {mail['subject']} ===")
        print(f"De: {mail['sender']} | Fecha: {mail['received_time']}")
        print("--- INICIO CUERPO ---")
        # Mostrar los primeros 500 caracteres del cuerpo
        print(mail['body'][:500] if mail['body'] else "(Cuerpo vacío)")
        print("--- FIN CUERPO ---\n" + "="*60 + "\n")
