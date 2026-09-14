import sys
import os

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

from outlook_service import OutlookService
from parser_service import EmailParser
from clipboard_service import ClipboardService

def main():
    print("=" * 65)
    print("     ASISTENTE DE ALTAS - OUTLOOK CLASSIC (MAPI ULTRA-RÁPIDO)")
    print("=" * 65)

    print("\n[+] Conectando con Microsoft Outlook Classic en tu ordenador...")
    service = OutlookService()
    
    if not service.connect():
        print("[-] Error: No se pudo conectar con Outlook Classic.")
        print("Asegúrate de que la aplicación Microsoft Outlook esté abierta en Windows.")
        return

    print("[+] Buscando solicitudes de alta recibidas...")
    emails = service.get_alta_emails(search_term="alta", max_results=25)

    if not emails:
        print("[!] No se encontraron correos con la palabra 'alta' en la bandeja.")
        return

    print(f"\n[OK] ¡Conexión con Outlook Classic exitosa! {len(emails)} solicitudes encontradas:\n")
    print(f"{'Nº':<4} | {'ESTADO':<7} | {'PERSONA / ASUNTO':<45} | {'FECHA'}")
    print("-" * 75)

    for mail in emails:
        status = "NUEVO" if mail["unread"] else "LEÍDO"
        display_name = mail["extracted_name"][:44]
        date_str = mail["received_time"].split()[0] if mail["received_time"] else ""
        print(f"[{mail['id']:<2}] | {status:<7} | {display_name:<45} | {date_str}")

    print("-" * 75)

    if len(sys.argv) > 1 and sys.argv[1].isdigit():
        choice = int(sys.argv[1])
    else:
        try:
            user_input = input("\n>> Ingresa el Nº de la persona a procesar (o 'q' para salir): ")
            if user_input.lower() in ['q', 'exit', 'salir']:
                print("Operación cancelada.")
                return
            choice = int(user_input)
        except ValueError:
            print("Entrada no válida.")
            return

    selected_mail = next((m for m in emails if m["id"] == choice), None)
    if not selected_mail:
        print("Número no encontrado en la lista.")
        return

    print("\n" + "="*65)
    print(f"PROCESANDO ALTA DE: {selected_mail['extracted_name'].upper()}")
    print("="*65)

    parsed_data = EmailParser.parse_body(selected_mail["body"])

    print("\n[ DATOS EXTRAÍDOS DEL CORREO ]")
    print(f"  * Nombre completo  : {parsed_data.get('nombre_completo') or 'No detectado'}")
    print(f"  * DNI / NIE        : {parsed_data.get('dni_nie') or 'No detectado'}")
    print(f"  * Nº Seg. Social   : {parsed_data.get('num_seguridad_social') or 'No detectado'}")
    print(f"  * Fecha de Alta    : {parsed_data.get('fecha_alta') or 'No detectado'}")
    print(f"  * Empresa          : {parsed_data.get('empresa_alta') or 'No detectado'}")
    print(f"  * Puesto           : {parsed_data.get('puesto_trabajo') or 'No detectado'}")
    print(f"  * Horas            : {parsed_data.get('num_horas') or 'No detectado'}")
    print(f"  * Centro           : {parsed_data.get('centro_trabajo') or 'No detectado'}")

    sequence = ClipboardService.prepare_sequence(parsed_data)

    print("\n" + "="*65)
    print("MODO ASISTENTE DE PORTAPAPELES PARA SEGURIDAD SOCIAL")
    print("="*65)
    print("Instrucciones: Se irá copiando cada dato al portapapeles.")
    print("En la web del Gobierno, haz clic en el campo y pulsa Ctrl+V.")
    print("Pulsa ENTER en esta consola para pasar al siguiente dato.\n")

    for idx, (label, val) in enumerate(sequence, 1):
        ClipboardService.set_clipboard(val)
        print(f"  --> [{idx}/{len(sequence)}] Copiado al portapapeles: {label} = '{val}'")
        if len(sys.argv) <= 1:
            input("      [Presiona ENTER para copiar el siguiente dato...]")

    print("\n[OK] ¡Todos los datos han sido preparados en secuencia!")

if __name__ == "__main__":
    main()
