import re

class EmailParser:
    @staticmethod
    def parse_body(body_text):
        """
        Extrae los campos de la plantilla estándar de correos de Gestio Altes.
        """
        data = {
            "nombre_completo": None,
            "dni_nie": None,
            "num_seguridad_social": None,
            "nacionalidad": None,
            "alojamiento": None,
            "puesto_trabajo": None,
            "empresa_alta": None,
            "num_horas": None,
            "fecha_alta": None,
            "zona_trabajo": None,
            "centro_trabajo": None,
            "raw_body": body_text
        }

        if not body_text:
            return data

        # Patrones específicos basados en la plantilla real
        patterns = {
            "nombre_completo": r'Nombre\s+completo\s*:\s*(.+)',
            "dni_nie": r'N[úu]mero\s+DNI/NIE\s*:\s*(.+)',
            "num_seguridad_social": r'N[ºo]?\s*Seguridad\s+social\s*:\s*(.+)',
            "nacionalidad": r'Nacionalidad\s*:\s*(.+)',
            "alojamiento": r'¿?Con\s+alojamiento\?\s*:\s*(.+)',
            "puesto_trabajo": r'Puesto\s+de\s+trabajo\s*:\s*(.+)',
            "empresa_alta": r'Empresa\s+de\s+alta\s*:\s*(.+)',
            "num_horas": r'N[úu]mero\s+de\s+horas\s*:\s*(.+)',
            "fecha_alta": r'Fecha\s+de\s+alta\s+solicitada\s*:\s*(.+)',
            "zona_trabajo": r'Zona\s+de\s+trabajo\s*:\s*(.+)',
            "centro_trabajo": r'Centro\s+de\s+trabajo\s*:\s*(.+)',
        }

        for field, pattern in patterns.items():
            match = re.search(pattern, body_text, re.IGNORECASE)
            if match:
                value = match.group(1).strip()
                data[field] = value

        # Limpieza secundaria en caso de formatos variados
        if data["dni_nie"]:
            data["dni_nie"] = data["dni_nie"].upper()

        if data["num_seguridad_social"]:
            data["num_seguridad_social"] = re.sub(r'\D', '', data["num_seguridad_social"])

        return data
