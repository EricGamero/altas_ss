import time
import win32clipboard

class ClipboardService:
    @staticmethod
    def set_clipboard(text):
        """Copia el texto al portapapeles de Windows."""
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(str(text), win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
            return True
        except Exception as e:
            print(f"Error al acceder al portapapeles: {e}")
            return False

    @staticmethod
    def prepare_sequence(parsed_data):
        """
        Prepara la lista ordenada de campos para pegar en la Seguridad Social.
        """
        sequence = [
            ("1. DNI / NIE", parsed_data.get("dni_nie", "")),
            ("2. Nº Seguridad Social (NAF)", parsed_data.get("num_seguridad_social", "")),
            ("3. Nombre Completo", parsed_data.get("nombre_completo", "")),
            ("4. Fecha de Alta", parsed_data.get("fecha_alta", "")),
            ("5. Horas / Jornada", parsed_data.get("num_horas", "")),
            ("6. Empresa de Alta", parsed_data.get("empresa_alta", "")),
            ("7. Puesto de Trabajo", parsed_data.get("puesto_trabajo", "")),
            ("8. Centro de Trabajo", parsed_data.get("centro_trabajo", ""))
        ]
        # Filtrar campos que tengan valor
        return [item for item in sequence if item[1]]
