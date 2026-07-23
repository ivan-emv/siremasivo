# =========================================================
# APP INDEPENDIENTE - CARGA MASIVA DE RETRASOS
# Generada a partir de la herramienta principal EMV SIRE.
# Mantiene únicamente el módulo de carga masiva y un Admin mínimo
# para refrescar catálogos/cache.
# =========================================================
import faulthandler
faulthandler.enable()

print('[BOOT 01] Iniciando importaciones de la app', flush=True)
import streamlit as st
import datetime
import re
print('[BOOT 02] Importaciones básicas completadas', flush=True)

# =========================================================
# CONFIGURACIÓN DE GOOGLE SHEETS
# - SHEET_ID_BD: base de datos de catálogos (Ciudades, Hoteles, Usuarios, etc.)
# - SHEET_ID_REGISTROS: registro/consulta de incidencias (DATOS y COMPLETO)
# =========================================================
SHEET_ID_BD = "1FyWpAjXMkuOW4TM71Z521lFyTX6nUQ8hNE8RGY3cnS4"
SHEET_ID_REGISTROS = "19v6_WKu7dNoRiyRwgP4ZbXF047jd-MZJba6lJGP3iT0"

WS_DATOS = "DATOS"
TAMANO_BLOQUE_MASIVA = 15

RESOLUCIONES = [
    "SELECCIONE", "Reembolso Parcial/Partial Reimbursement", "Reembolso Total/Total Reimbursement",
    "Compensación/Compensation", "Descuento Próximo Viaje/Next Trip Discount",
    "Cambio Itinerario/Itinerary Change", "En Estudio/Pending",
    "Se informa al Pasajero/Passenger Informed", "Se informa al Operador/Operator Informed",
    "Se informa al Minorista/Agency Informed", "Se informa al Guía/Guide Informed",
    "Se informa al Transferista/TSP Informed", "Se informa al Receptivo/Local Provider Informed",
    "Se informa a Departamento/Department Informed"
]


# Headers oficiales del registro (WS_DATOS). Se usan también para lecturas seguras
# evitando que el lock en Z1 interfiera con get_all_records(expected_headers=HEADERS_REGISTRO).
HEADERS_REGISTRO = [
    "fecha_inicio",
    "fecha_registro",
    "momento_viaje",
    "medio_contacto",
    "localizador",
    "nombre_usuario",
    "operador",
    "quien_contacta",
    "ciudad",
    "tipo_contacto",
    "area",
    "area_relacionada",
    "hotel",
    "tipo_traslado",
    "trayecto",
    "guia",
    "tipo_incidencia",
    "comentario",
    "resolucion",
    "monto",
    "resultado",
    "incidencia_id",
]


# =========================================================
# UTILIDADES
# =========================================================
def init_session():
    if "admin_autenticado" not in st.session_state:
        st.session_state.admin_autenticado = False
    if "admin_usuario" not in st.session_state:
        st.session_state.admin_usuario = ""

    # Control de guardado para Carga Masiva Retrasos
    if "guardando_masiva" not in st.session_state:
        st.session_state.guardando_masiva = False
    if "guardar_masiva_pendiente" not in st.session_state:
        st.session_state.guardar_masiva_pendiente = False
    if "masiva_guardado_ok" not in st.session_state:
        st.session_state.masiva_guardado_ok = False
    if "masiva_upload_id" not in st.session_state:
        st.session_state.masiva_upload_id = ""
    if "masiva_registros_eliminados" not in st.session_state:
        st.session_state.masiva_registros_eliminados = {}
    if "masiva_bloque_actual" not in st.session_state:
        st.session_state.masiva_bloque_actual = 0
    if "masiva_bloques_guardados" not in st.session_state:
        st.session_state.masiva_bloques_guardados = []
    if "masiva_resumen_bloques" not in st.session_state:
        st.session_state.masiva_resumen_bloques = []
    if "masiva_df_preparado" not in st.session_state:
        st.session_state.masiva_df_preparado = None
    if "masiva_errores_archivo" not in st.session_state:
        st.session_state.masiva_errores_archivo = []
    if "masiva_df_upload_id" not in st.session_state:
        st.session_state.masiva_df_upload_id = ""


def get_gspread_client():
    """Crea un cliente gspread mediante google-auth y el service account de secrets.toml."""
    print("[BOOT/API] Preparando cliente de Google Sheets", flush=True)
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    if "gcp_service_account" not in st.secrets:
        raise RuntimeError(
            "No se encontró la sección [gcp_service_account] en los Secrets de Streamlit."
        )

    credentials_info = dict(st.secrets["gcp_service_account"])
    credentials = Credentials.from_service_account_info(
        credentials_info,
        scopes=scopes,
    )
    return gspread.authorize(credentials)


def normalizar_fecha_masiva(valor) -> str:
    """Normaliza fechas de carga masiva a DD/MM/AAAA. Soporta AAMMDD/AAAAMMDD, números Excel y fechas pandas."""
    import pandas as pd

    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass

    if isinstance(valor, datetime.datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, datetime.date):
        return valor.strftime("%d/%m/%Y")

    texto = str(valor).strip()
    if not texto or texto.lower() in ["nan", "nat", "none"]:
        return ""

    # Cuando pandas lee 20260418 como 20260418.0
    if re.match(r"^\d+\.0$", texto):
        texto = texto.split(".")[0]

    solo_digitos = re.sub(r"[^0-9]", "", texto)

    # Formato principal del XLS: AAAAMMDD, por ejemplo 20260418
    if re.match(r"^\d{8}$", solo_digitos):
        try:
            return datetime.datetime.strptime(solo_digitos, "%Y%m%d").strftime("%d/%m/%Y")
        except Exception:
            pass

    # Fallback por si viniera DDMMYYYY
    if re.match(r"^\d{8}$", solo_digitos):
        try:
            return datetime.datetime.strptime(solo_digitos, "%d%m%Y").strftime("%d/%m/%Y")
        except Exception:
            pass

    # Fallback para formatos parseables por pandas
    try:
        fecha = pd.to_datetime(texto, dayfirst=True, errors="coerce")
        if pd.notna(fecha):
            return fecha.strftime("%d/%m/%Y")
    except Exception:
        pass

    return ""



def fecha_masiva_a_date(valor):
    """Convierte una fecha de carga masiva a date usando normalizar_fecha_masiva."""
    texto = normalizar_fecha_masiva(valor)
    if not texto:
        return None
    try:
        return datetime.datetime.strptime(texto, "%d/%m/%Y").date()
    except Exception:
        return None


def calcular_momento_viaje_masivo(fecha_inicio, fecha_finalizacion, fecha_evento) -> str:
    """Calcula momento_viaje según fechas del XLS.

    Reglas:
    - FECHA < FECHA INICIO => Pre Viaje/Pre Tour
    - FECHA entre FECHA INICIO y FECHA DE FINALIZACIÓN => En Ruta/On Route
    - FECHA > FECHA DE FINALIZACIÓN => Post Viaje/Post Tour
    """
    fi = fecha_masiva_a_date(fecha_inicio)
    ff = fecha_masiva_a_date(fecha_finalizacion)
    fe = fecha_masiva_a_date(fecha_evento)

    if not fi or not ff or not fe:
        return ""

    if fe < fi:
        return "Pre Viaje/Pre Tour"
    if fi <= fe <= ff:
        return "En Ruta/On Route"
    if fe > ff:
        return "Post Viaje/Post Tour"
    return ""

def limpiar_valor_masivo(valor) -> str:
    """Convierte valores de Excel/HTML a texto limpio, evitando nan."""
    import pandas as pd
    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass
    texto = str(valor or "").strip()
    return "" if texto.lower() in ["nan", "nat", "none"] else texto


def normalizar_nombre_columna_masiva(nombre) -> str:
    """Normaliza nombres de columnas: mayúsculas, sin tildes y espacios únicos."""
    import unicodedata
    texto = str(nombre or "").strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"\s+", " ", texto)
    return texto


def normalizar_texto_match_masivo(valor) -> str:
    """Normaliza texto para búsquedas flexibles: sin tildes, mayúsculas y espacios únicos."""
    return normalizar_nombre_columna_masiva(valor)


def nombre_operador_sin_pais_masivo(operador_catalogo: str) -> str:
    """Extrae el nombre del operador cuando el catálogo viene como 'PAÍS - OPERADOR'."""
    texto = limpiar_valor_masivo(operador_catalogo)
    if " - " in texto:
        return texto.split(" - ", 1)[1].strip()
    return texto


def filtrar_operadores_por_nombre_masivo(operadores_catalogo: list, operador_xls: str) -> list:
    """Devuelve operadores del catálogo cuyo nombre comercial coincide con el valor del XLS.

    Prioridad operativa:
      1. Coincidencia exacta contra el nombre comercial después de "PAÍS - ".
         Ejemplo: XLS='EUROMUNDO' -> ['MEXICO - EUROMUNDO']
      2. Si no hay exacta, coincidencia exacta contra el texto completo del catálogo.
      3. Si no hay exacta, coincidencias parciales como apoyo/manual.
    """
    raw = normalizar_texto_match_masivo(operador_xls)
    if not raw:
        return _deduplicar_preservando_orden(operadores_catalogo or [])

    exactos_nombre = []
    exactos_catalogo = []
    parciales = []
    for op in operadores_catalogo or []:
        op_txt = limpiar_valor_masivo(op)
        nombre_txt = nombre_operador_sin_pais_masivo(op_txt)
        op_norm = normalizar_texto_match_masivo(op_txt)
        nombre_norm = normalizar_texto_match_masivo(nombre_txt)

        if raw == nombre_norm:
            exactos_nombre.append(op_txt)
        elif raw == op_norm:
            exactos_catalogo.append(op_txt)
        elif raw in nombre_norm or raw in op_norm or nombre_norm in raw:
            parciales.append(op_txt)

    if exactos_nombre:
        return _deduplicar_preservando_orden(exactos_nombre)
    if exactos_catalogo:
        return _deduplicar_preservando_orden(exactos_catalogo)
    return _deduplicar_preservando_orden(parciales)


def sugerir_operador_catalogo_masivo(operadores_catalogo: list, operador_xls: str) -> str:
    """Sugiere operador de catálogo solo cuando hay una coincidencia única."""
    coincidencias = filtrar_operadores_por_nombre_masivo(operadores_catalogo, operador_xls)
    return coincidencias[0] if len(coincidencias) == 1 else ""


def obtener_tipos_incidencia_por_area_masiva(area_valor: str) -> list:
    """Devuelve opciones de Tipo de Incidencia según Área/Área Relacionada."""
    area_norm = str(area_valor or "").strip()
    if area_norm == "Hotel":
        return [
            "", "Desayuno/Breakfast", "Limpieza-Bichos/Cleanliness-Bugs", "Comodidad/Comfort",
            "Ubicación/Location", "Mantenimiento General/Overall Maintenance",
            "Habitación/Room", "Robo-Hurto/Theft-Robbery", "Falta Reserva/Reservation Missing",
            "Noches Adicionales/Additional Nights", "Otro/Other"
        ]
    if area_norm == "Guías/Guides":
        return [
            "", "Actitud/Attitude", "Felicitación/Congratulation", "Conocimiento/Knowledge",
            "Idioma/Language", "Guía Local - Mal Servicio/Local Guide - Poor Service",
            "Pérdida Equipaje/Loss of Luggage", "Versiones Contradictorias/Contradictory Versions",
            "Otro/Other"
        ]
    if area_norm == "Traslados/Transfers":
        return [
            "", "TRF - No Show - PAX", "TRF - No Show - Transfer", "TRF - Pendiente Datos/Pending data",
            "TRF - Error EMV/EMV´s error", "TRF - Actitud Chófer/Driver´s Attitude",
            "TRF - Versiones Contradictorias/Contradictory Versions",
            "TRF - No Incluido-Solicitado/Not Included-Requested",
            "TRF - Retraso PAX no notificado/Unnotified PAX Delay",
            "TRF - Felicitación/Congratulation", "TRF - Otro/Other",
            "BUS - Accidente/Accident", "BUS - Mantenimiento-Falla/Breakdown-Maintenance",
            "BUS - Hurto-Robo en Cabina/Theft-Robbery in the Cabin",
            "BUS - Comodidad - AC / Comfort - AC",
            "BUS - Actitud Chofer/Driver's Attitude",
            "BUS - Felicitación/Congratulation", "BUS - Otro/Other"
        ]
    if area_norm == "Generales/General" or area_norm == "Itinerario/Itinerary" or area_norm == "Otros/Other":
        return [
            "", "Aéreos - KANNAK",
            "Itinerario - Fuerza Mayor/Force Majeure", "Itinerario - Muchos Idiomas/Several Languages",
            "Itinerario - Parada en Tiendas/Shop Stops",
            "Itinerario - Itinerario no Seguido/Unfollowed Timetable",
            "Itinerario - Otro/Other", "Asistencia - No relacionado a EMV/No relation to EMV",
            "Bote/Ferry/Crucero - Cambio Itinerario/Itinerary change",
            "Booking - Error Agente/Agent Error (AGT/TTOO)",
            "Seguro-Call Center - Info Incorrecta/Inaccurate Info",
            "Equipaje - Demora-Pérdida-Daño/Delay-Loss-Damage",
            "Comidas - Calidad-Cantidad/Quality-Quantity",
            "Opcionales - No Realizado/Not done",
            "Opcionales - Incidente en Pago/Payment Issue",
            "Opcionales - Fecha Errada/Wrong Date",
            "Opcionales - Otros/Other",
            "Personal - PAX No Show/No Show PAX",
            "Personal - Enfermedad-Lesión/Illness-Injury", "Otros - General"
        ]
    return [""]


def leer_archivo_retrasos_xls(uploaded_file):
    """Lee .xls exportado como HTML o Excel real y devuelve un DataFrame normalizado.

    Nota operativa:
    - El archivo legacy .xls de retrasos suele ser HTML con extensión .xls.
    - Para evitar el error "Excel file format cannot be determined", se lee el contenido
      como bytes y se decide el parser antes de llamar a pandas.
    """
    import pandas as pd
    from io import BytesIO, StringIO

    if uploaded_file is None:
        return pd.DataFrame()

    nombre = (getattr(uploaded_file, "name", "") or "").lower()

    try:
        contenido = uploaded_file.getvalue()
    except Exception:
        uploaded_file.seek(0)
        contenido = uploaded_file.read()

    if not contenido:
        return pd.DataFrame()

    inicio = contenido[:500].lstrip().lower()
    parece_xlsx = nombre.endswith(".xlsx") or contenido[:2] == b"PK"
    parece_html = inicio.startswith(b"<") or b"<html" in inicio or b"<table" in contenido[:5000].lower()

    if parece_xlsx:
        try:
            return pd.read_excel(BytesIO(contenido), engine="openpyxl")
        except ImportError as e:
            raise RuntimeError("Para leer archivos .xlsx debes agregar 'openpyxl' al requirements.txt.") from e

    if parece_html:
        texto = None
        for encoding in ("utf-8-sig", "utf-8", "latin1"):
            try:
                texto = contenido.decode(encoding)
                break
            except Exception:
                continue
        if texto is None:
            raise RuntimeError("No se pudo decodificar el archivo .xls exportado como HTML.")

        try:
            tablas = pd.read_html(StringIO(texto))
        except ImportError as e:
            raise RuntimeError("Para leer este .xls exportado como HTML debes agregar 'lxml' al requirements.txt.") from e
        except Exception as e:
            raise RuntimeError(f"El archivo parece HTML, pero no se pudo interpretar la tabla: {e}") from e

        if not tablas:
            raise RuntimeError("El archivo HTML no contiene tablas reconocibles.")
        return tablas[0]

    # Fallback para .xls binario real. Solo aplica si realmente no es HTML.
    if nombre.endswith(".xls"):
        try:
            return pd.read_excel(BytesIO(contenido), engine="xlrd")
        except ImportError as e:
            raise RuntimeError("Este .xls parece binario real. Para leerlo debes agregar 'xlrd' al requirements.txt, o convertirlo a .xlsx.") from e

    raise RuntimeError("Formato no reconocido. Carga un .xls HTML exportado desde el sistema o un .xlsx válido.")


def preparar_editor_retrasos_desde_df(df, nombre_usuario: str, operadores_catalogo: list = None) -> tuple[object, list]:
    """Prepara un DataFrame editable: Usuario es global; el resto se completa por fila."""
    import pandas as pd

    errores = []
    columnas_requeridas = ["RESERVA", "OBSERVACION", "OPERADOR", "FECHA INICIO", "FECHA DE FINALIZACION", "FECHA"]
    columnas_actuales = {normalizar_nombre_columna_masiva(c): c for c in df.columns}
    faltantes = [c for c in columnas_requeridas if c not in columnas_actuales]
    if faltantes:
        return pd.DataFrame(), ["Faltan columnas obligatorias en el archivo: " + ", ".join(faltantes)]

    filas = []
    for i, row in df.iterrows():
        localizador = limpiar_valor_masivo(row.get(columnas_actuales["RESERVA"], "")).upper()
        comentario = limpiar_valor_masivo(row.get(columnas_actuales["OBSERVACION"], ""))
        operador_xls = limpiar_valor_masivo(row.get(columnas_actuales["OPERADOR"], ""))
        operador = sugerir_operador_catalogo_masivo(operadores_catalogo or [], operador_xls) or operador_xls
        fecha_inicio = normalizar_fecha_masiva(row.get(columnas_actuales["FECHA INICIO"], ""))
        fecha_finalizacion = normalizar_fecha_masiva(row.get(columnas_actuales["FECHA DE FINALIZACION"], ""))
        fecha_registro = normalizar_fecha_masiva(row.get(columnas_actuales["FECHA"], ""))
        momento_viaje_auto = calcular_momento_viaje_masivo(fecha_inicio, fecha_finalizacion, fecha_registro)

        if not localizador and not comentario:
            continue

        fila_num = i + 2
        if not localizador:
            errores.append(f"Fila {fila_num}: falta RESERVA/localizador.")
        if not comentario:
            errores.append(f"Fila {fila_num}: falta OBSERVACION/comentario.")
        if not operador:
            errores.append(f"Fila {fila_num}: falta OPERADOR.")
        if not fecha_inicio:
            errores.append(f"Fila {fila_num}: FECHA INICIO no tiene formato válido.")
        if not fecha_finalizacion:
            errores.append(f"Fila {fila_num}: FECHA DE FINALIZACIÓN no tiene formato válido.")
        if not fecha_registro:
            errores.append(f"Fila {fila_num}: FECHA no tiene formato válido.")

        filas.append({
            "fecha_inicio": fecha_inicio,
            "fecha_finalizacion": fecha_finalizacion,
            "fecha_registro": fecha_registro,
            "momento_viaje": momento_viaje_auto,
            "localizador": localizador,
            "nombre_usuario": nombre_usuario,
            "operador_xls": operador_xls,
            "operador": operador,
            "comentario": comentario,
            "medio_contacto": "",
            "quien_contacta": "",
            "ciudad": "",
            "tipo_contacto": "Reclamación/Complaint",
            "area": "",
            "area_relacionada": "Traslados/Transfers",
            "hotel": "",
            "tipo_traslado": "",
            "trayecto": "",
            "guia": "",
            "tipo_incidencia": "TRF - Retraso PAX no notificado/Unnotified PAX Delay",
            "resolucion": "",
            "monto": "",
            "resultado": "",
            "incidencia_id": "",
        })

    return pd.DataFrame(filas), errores


def construir_filas_retrasos_desde_editor(df_editor, nombre_usuario: str) -> tuple[list, list]:
    """Convierte la tabla editada a filas oficiales para Google Sheets."""
    import pandas as pd

    filas = []
    errores = []
    if df_editor is None or df_editor.empty:
        return [], ["No hay registros para guardar."]

    requeridos_comunes = [
        "fecha_inicio", "fecha_registro", "localizador", "operador",
        "momento_viaje", "medio_contacto", "quien_contacta", "ciudad", "tipo_contacto",
    ]

    for i, row in df_editor.iterrows():
        fila = {col: "" for col in HEADERS_REGISTRO}
        for col in HEADERS_REGISTRO:
            if col in df_editor.columns:
                fila[col] = limpiar_valor_masivo(row.get(col, ""))

        fila["nombre_usuario"] = nombre_usuario
        fila["localizador"] = limpiar_valor_masivo(fila.get("localizador", "")).upper()
        fila["fecha_inicio"] = normalizar_fecha_masiva(fila.get("fecha_inicio", ""))
        fila["fecha_registro"] = normalizar_fecha_masiva(fila.get("fecha_registro", ""))
        fila["incidencia_id"] = ""

        fila_num = i + 1
        tipo_contacto = limpiar_valor_masivo(fila.get("tipo_contacto", ""))
        area = limpiar_valor_masivo(fila.get("area", ""))
        area_relacionada = limpiar_valor_masivo(fila.get("area_relacionada", ""))
        tipo_incidencia = limpiar_valor_masivo(fila.get("tipo_incidencia", ""))

        for col in requeridos_comunes:
            if not limpiar_valor_masivo(fila.get(col, "")):
                errores.append(f"Registro {fila_num}: falta completar '{col}'.")

        if tipo_contacto == "Información/Information":
            if not area:
                errores.append(f"Registro {fila_num}: falta completar 'area'.")
            if not limpiar_valor_masivo(fila.get("resolucion", "")):
                errores.append(f"Registro {fila_num}: falta completar 'resolucion'.")
            if area == "Hotel" and not limpiar_valor_masivo(fila.get("hotel", "")):
                errores.append(f"Registro {fila_num}: el área Hotel requiere informar 'hotel'.")
            if area == "Traslados/Transfers" and not limpiar_valor_masivo(fila.get("tipo_traslado", "")):
                errores.append(f"Registro {fila_num}: el área Traslados/Transfers requiere informar 'tipo_traslado'.")

        elif tipo_contacto == "Reclamación/Complaint":
            if not area_relacionada:
                errores.append(f"Registro {fila_num}: falta completar 'area_relacionada'.")
            if not tipo_incidencia:
                errores.append(f"Registro {fila_num}: falta completar 'tipo_incidencia'.")
            if not limpiar_valor_masivo(fila.get("resolucion", "")):
                errores.append(f"Registro {fila_num}: falta completar 'resolucion'.")
            if not limpiar_valor_masivo(fila.get("resultado", "")):
                errores.append(f"Registro {fila_num}: falta completar 'resultado'.")

            if area_relacionada == "Hotel" and not limpiar_valor_masivo(fila.get("hotel", "")):
                errores.append(f"Registro {fila_num}: el área Hotel requiere informar 'hotel'.")
            if area_relacionada == "Guías/Guides":
                if not limpiar_valor_masivo(fila.get("trayecto", "")):
                    errores.append(f"Registro {fila_num}: el área Guías/Guides requiere informar 'trayecto'.")
                if not limpiar_valor_masivo(fila.get("guia", "")):
                    errores.append(f"Registro {fila_num}: el área Guías/Guides requiere informar 'guia'.")
            if area_relacionada == "Traslados/Transfers":
                if tipo_incidencia.startswith("TRF") and not limpiar_valor_masivo(fila.get("tipo_traslado", "")):
                    errores.append(f"Registro {fila_num}: las incidencias TRF requieren informar 'tipo_traslado'.")
                if tipo_incidencia.startswith("BUS") and not limpiar_valor_masivo(fila.get("trayecto", "")):
                    errores.append(f"Registro {fila_num}: las incidencias BUS requieren informar 'trayecto'.")
            if area_relacionada == "Generales/General" and tipo_incidencia.startswith("Itinerario"):
                if not limpiar_valor_masivo(fila.get("trayecto", "")):
                    errores.append(f"Registro {fila_num}: las incidencias de Itinerario requieren informar 'trayecto'.")

        elif tipo_contacto == "Otro/Other":
            if not limpiar_valor_masivo(fila.get("resolucion", "")):
                errores.append(f"Registro {fila_num}: falta completar 'resolucion'.")

        elif tipo_contacto == "Cuestionario de Satisfacción":
            if not area_relacionada:
                errores.append(f"Registro {fila_num}: falta completar 'area_relacionada'.")
            if not tipo_incidencia:
                errores.append(f"Registro {fila_num}: falta completar 'tipo_incidencia'.")
            if area_relacionada == "Hotel" and not limpiar_valor_masivo(fila.get("hotel", "")):
                errores.append(f"Registro {fila_num}: el área Hotel requiere informar 'hotel'.")
            if area_relacionada == "Guías/Guides" and not limpiar_valor_masivo(fila.get("guia", "")):
                errores.append(f"Registro {fila_num}: el área Guías/Guides requiere informar 'guia'.")

        if fila.get("resolucion", "").startswith("Reembolso") or fila.get("resolucion", "") == "Compensación/Compensation":
            if not limpiar_valor_masivo(fila.get("monto", "")):
                errores.append(f"Registro {fila_num}: la resolución requiere informar 'monto'.")

        filas.append(fila)

    return filas, errores

def _deduplicar_preservando_orden(valores: list) -> list:
    """Elimina duplicados preservando el orden original."""
    resultado = []
    vistos = set()
    for valor in valores or []:
        v = str(valor or "").strip()
        if not v:
            continue
        clave = v.upper()
        if clave in vistos:
            continue
        vistos.add(clave)
        resultado.append(v)
    return resultado



def filtrar_hoteles_por_ciudad(hoteles: list, ciudad: str) -> list:
    """Filtra hoteles por ciudad usando el prefijo 'CIUDAD - ' y fuerza la opción general 'SIN ESPECIFICAR'."""
    if not hoteles:
        return []
    if not ciudad or ciudad == "SELECCIONE":
        return _deduplicar_preservando_orden(hoteles)

    c = str(ciudad).strip().upper()
    pref = c + " -"
    opciones_generales = []
    filtrados = []

    for h in hoteles:
        hs = str(h or "").strip()
        if not hs:
            continue
        hs_upper = hs.upper()
        if hs_upper == "SIN ESPECIFICAR":
            opciones_generales.append(hs)
        elif hs_upper.startswith(pref):
            filtrados.append(hs)

    return _deduplicar_preservando_orden(opciones_generales + filtrados)



def filtrar_trayectos_por_ciudad(trayectos: list, ciudad: str) -> list:
    """Filtra trayectos por ciudad y fuerza siempre las opciones generales definidas en catálogo."""
    if not trayectos:
        return []
    if not ciudad or ciudad == "SELECCIONE":
        return _deduplicar_preservando_orden(trayectos)

    c = str(ciudad).strip().upper()
    opciones_fijas = [
        "TODOS LOS TRAYECTOS",
        "ITINERARIO GRUPO PRIVADO",
        "ITINERARIO GRUPO CAM",
        "SIN ESPECIFICAR",
    ]
    opciones_generales = []
    filtrados = []

    for t in trayectos:
        ts = str(t or "").strip()
        if not ts:
            continue
        ts_upper = ts.upper()
        if ts_upper in opciones_fijas:
            opciones_generales.append(ts)
        elif c and c in ts_upper:
            filtrados.append(ts)

    return _deduplicar_preservando_orden(opciones_generales + filtrados)





# =========================================================
# PERSISTENCIA (GUARDADO EN GOOGLE SHEETS)
# =========================================================
# =========================================================
# PERSISTENCIA (GUARDADO EN GOOGLE SHEETS)
# =========================================================
def guardar_lote_google_sheets_seguro(lista_filas: list):
    """Guarda un lote de filas (dicts ya consolidados) en WS_DATOS, con lock simple (celda Z1) para minimizar colisiones.
    Optimizaciones:
      - Evita sheet.get_all_values() (lectura completa) para calcular la siguiente fila libre.
      - Evita leer toda la columna 'incidencia_id' en cada guardado; usa un contador en Z2 (bajo lock).
    """
    import pandas as pd
    from gspread_dataframe import set_with_dataframe
    import time
    import datetime

    if not lista_filas:
        st.warning("No hay incidencias para guardar.")
        return False

    client = get_gspread_client()
    sheet = client.open_by_key(SHEET_ID_REGISTROS).worksheet(WS_DATOS)

    def _clean_header_row(row_vals: list) -> list:
        """Limpia el header para que el lock en Z1 ('LOCKED') no se interprete como un header real."""
        if not row_vals:
            return []
        row = list(row_vals)

        # Si Z1 está siendo usado como lock, puede aparecer como 'LOCKED' en la posición 26 (columna Z).
        # Lo reemplazamos por vacío para no extender artificialmente el header.
        if len(row) >= 26 and str(row[25]).strip().upper() == "LOCKED":
            row[25] = ""

        # Recorta a la última columna no vacía
        last = 0
        for i, v in enumerate(row, start=1):
            if str(v).strip() != "":
                last = i
        return row[:last] if last else []

    try:
        # --- Lock ---
        if sheet.acell("Z1").value == "LOCKED":
            st.error("⚠️ El sistema está siendo utilizado por otro usuario. Intenta en unos segundos.")
            return False

        sheet.update_acell("Z1", "LOCKED")

        # --- Header limpio ---
        header_row_raw = sheet.row_values(1) or []
        header_row = _clean_header_row(header_row_raw)

        # --- Asegurar columna incidencia_id ---
        if "incidencia_id" not in header_row:
            col_new = len(header_row) + 1  # primera columna libre real
            sheet.update_cell(1, col_new, "incidencia_id")
            header_row = _clean_header_row(sheet.row_values(1) or [])

        col_incid = header_row.index("incidencia_id") + 1  # 1-indexed

        # --- ID transversal secuencial (EMV-AÑO-XXXXX) usando contador en Z2 ---
        year = datetime.datetime.now().year
        pref = f"EMV-{year}-"

        counter_cell = "Z2"
        counter_val = str(sheet.acell(counter_cell).value or "").strip()

        def _parse_id(s: str) -> int:
            s = str(s or "").strip()
            if s.startswith(pref):
                tail = s[len(pref):]
                if tail.isdigit():
                    return int(tail)
            return 0

        max_n = _parse_id(counter_val)

        # Inicialización: si no hay contador válido en Z2, calculamos UNA VEZ leyendo la columna de IDs
        if max_n == 0:
            try:
                existing_ids = sheet.col_values(col_incid)[1:]  # sin header
                for v in existing_ids:
                    n = _parse_id(v)
                    if n > max_n:
                        max_n = n
            except Exception:
                max_n = 0

        next_n = max_n + 1
        if next_n > 99999:
            raise ValueError(f"Se alcanzó el máximo de IDs para el año {year} (99999).")

        incidencia_id_batch = f"{pref}{str(next_n).zfill(5)}"

        # Persistir el último ID usado en Z2 (bajo lock) para evitar lecturas completas futuras
        try:
            sheet.update_acell(counter_cell, incidencia_id_batch)
        except Exception:
            # Si por permisos o restricciones no se puede, no bloqueamos el guardado
            pass

        # Completa incidencia_id en filas del lote si no viene informado
        for fila in lista_filas:
            if not str(fila.get("incidencia_id", "")).strip():
                fila["incidencia_id"] = incidencia_id_batch

        headers = HEADERS_REGISTRO

        data = []
        for fila in lista_filas:
            row = [fila.get(col, "") for col in headers]
            data.append(row)

        # --- Calcular start_row SIN leer toda la hoja ---
        start_row = 2  # por defecto, después del header
        try:
            # Preferimos usar la columna 'localizador' como referencia (debe estar completa en cada fila)
            if "localizador" in header_row:
                col_loc = header_row.index("localizador") + 1
            else:
                # fallback: si header_row no tiene localizador (raro), usamos el orden esperado
                col_loc = headers.index("localizador") + 1 if "localizador" in headers else 1

            col_vals = sheet.col_values(col_loc)  # incluye header y llega hasta el último no-vacío
            start_row = len(col_vals) + 1 if col_vals else 2
        except Exception:
            # fallback conservador (costoso) solo si algo salió mal
            existing = sheet.get_all_values()
            start_row = len(existing) + 1

        df = pd.DataFrame(data, columns=headers)

        guardado_correcto = False
        for intento in range(3):
            try:
                set_with_dataframe(sheet, df, row=start_row, include_column_header=False)
                guardado_correcto = True
                break
            except Exception as e:
                if intento == 2:
                    st.error(f"❌ Error al guardar tras múltiples intentos: {e}")
                else:
                    time.sleep(1)

        return guardado_correcto

    finally:
        # liberar lock aunque haya error
        try:
            sheet.update_acell("Z1", "")
        except Exception:
            pass


# =========================================================
# INICIO APP


# =========================================================
# INICIO APP INDEPENDIENTE
# =========================================================
print("[BOOT 03] Inicializando session_state", flush=True)
init_session()
print("[BOOT 04] Configurando página Streamlit", flush=True)
st.set_page_config(page_title="Carga Masiva Retrasos - EMV SIRE", layout="wide")
print("[BOOT 05] Configuración de página completada", flush=True)

hide_streamlit_style = """
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header [data-testid="stToolbarActions"] {display: none !important;}
        [data-testid="stAppToolbar"] [data-testid="stToolbarActions"] {display: none !important;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)


# =========================================================
# ADMIN MÍNIMO: SOLO ACTUALIZAR CATÁLOGOS/CACHÉ
# =========================================================
@st.cache_data(show_spinner=False, ttl=3600)
def cargar_admin_users():
    client = get_gspread_client()
    worksheet = client.open_by_key(SHEET_ID_BD).worksheet("ADMIN")
    return worksheet.get_all_records()


def autenticar_admin(usuario, password):
    for admin in cargar_admin_users() or []:
        if str(admin.get("Usuario", "")).strip() == str(usuario or "").strip() and str(admin.get("Password", "")).strip() == str(password or "").strip():
            return True
    return False


with st.sidebar.expander("🔐 Acceso Administrador", expanded=not st.session_state.get("admin_autenticado", False)):
    if not st.session_state.get("admin_autenticado", False):
        admin_user = st.text_input("Usuario", key="admin_user_masiva")
        admin_pass = st.text_input("Contraseña", type="password", key="admin_pass_masiva")
        if st.button("Iniciar Sesión", key="btn_login_admin_masiva"):
            if autenticar_admin(admin_user, admin_pass):
                st.session_state.admin_autenticado = True
                st.session_state.admin_usuario = admin_user
                st.success(f"Acceso concedido. Bienvenido, {admin_user}.")
                st.rerun()
            else:
                st.error("Credenciales incorrectas.")
    else:
        st.success(f"🔓 Admin: {st.session_state.admin_usuario}")
        if st.button("🔄 Actualizar catálogos", key="btn_actualizar_catalogos_admin"):
            st.cache_data.clear()
            st.success("Catálogos actualizados.")
            st.rerun()
        if st.button("🚪 Cerrar sesión", key="btn_logout_admin_masiva"):
            st.session_state.admin_autenticado = False
            st.session_state.admin_usuario = ""
            st.rerun()

# Encabezado principal
col_logo, col_titulo = st.columns([2, 5])
with col_logo:
    try:
        st.image("a1.png", width=320)
    except Exception:
        st.empty()
with col_titulo:
    st.markdown("<h1 style='margin-top: 5px;'>Carga Masiva de Retrasos - EMV SIRE</h1>", unsafe_allow_html=True)
    st.caption("Herramienta independiente para carga masiva de incidencias por retrasos.")

st.caption(
    "Carga pública de incidencias desde archivo .xls/.xlsx. "
    "Solo el Usuario aplica a todo el lote; la clasificación restante se completa por cada registro."
)


@st.cache_data(show_spinner=False, ttl=3600)
def cargar_catalogos_masiva():
    client = get_gspread_client()
    datos = {}
    for nombre in ["Ciudades", "Hoteles", "Guias", "Operadores", "Trayectos", "Usuarios"]:
        worksheet = client.open_by_key(SHEET_ID_BD).worksheet(nombre)
        datos[nombre] = worksheet.get_all_records()
    return datos

print("[BOOT 06] Iniciando carga de catálogos", flush=True)
try:
    datos_bd_masiva = cargar_catalogos_masiva()
except Exception as exc:
    print(f"[BOOT ERROR] Falló la carga de catálogos: {type(exc).__name__}: {exc}", flush=True)
    st.error("No se pudieron cargar los catálogos desde Google Sheets.")
    st.exception(exc)
    st.stop()
print("[BOOT 07] Catálogos cargados correctamente", flush=True)

USUARIOS = [u.get("Nombre") for u in datos_bd_masiva.get("Usuarios", []) if u.get("Nombre")]
CIUDADES = [c.get("Ciudad") for c in datos_bd_masiva.get("Ciudades", []) if c.get("Ciudad")]
HOTELES = [h.get("Nombre Hotel") for h in datos_bd_masiva.get("Hoteles", []) if h.get("Nombre Hotel")]
GUIAS = [g.get("Nombre del Guia") for g in datos_bd_masiva.get("Guias", []) if g.get("Nombre del Guia")]
OPERADORES = [o.get("Nombre del Operador") for o in datos_bd_masiva.get("Operadores", []) if o.get("Nombre del Operador")]
TRAYECTOS = [t.get("Trayecto") for t in datos_bd_masiva.get("Trayectos", []) if t.get("Trayecto")]

st.subheader("Dato común del lote")
nombre_usuario = st.selectbox(
    "Usuario que aplica a todas las reservas cargadas",
    ["SELECCIONE"] + USUARIOS,
    key="masiva_usuario_global"
)

uploaded_file = st.file_uploader(
    "Cargar archivo de retrasos (.xls o .xlsx)",
    type=["xls", "xlsx"],
    key="upload_retrasos_masivo"
)

def limpiar_estado_editor_masivo():
    """Limpia valores editables por fila cuando se cambia o retira el archivo cargado."""
    prefijos = (
        "masiva_com_", "masiva_eliminar_", "masiva_limpiar_com_",
        "masiva_op_", "masiva_momento_", "masiva_medio_", "masiva_quien_",
        "masiva_ciudad_", "masiva_tipo_contacto_", "masiva_area_",
        "masiva_tipo_incidencia_", "masiva_resolucion_", "masiva_resultado_",
        "masiva_monto_", "masiva_hotel_", "masiva_trayecto_",
        "masiva_guia_", "masiva_tipo_traslado_",
    )
    for key in list(st.session_state.keys()):
        if str(key).startswith(prefijos):
            del st.session_state[key]

def limpiar_estado_registro_masivo(idx: int):
    """Elimina del session_state únicamente los widgets correspondientes a una fila."""
    sufijo = f"_{idx}"
    prefijos = (
        "masiva_com_", "masiva_eliminar_", "masiva_limpiar_com_",
        "masiva_op_", "masiva_momento_", "masiva_medio_", "masiva_quien_",
        "masiva_ciudad_", "masiva_tipo_contacto_", "masiva_area_",
        "masiva_tipo_incidencia_", "masiva_resolucion_", "masiva_resultado_",
        "masiva_monto_", "masiva_hotel_", "masiva_trayecto_",
        "masiva_guia_", "masiva_tipo_traslado_", "masiva_fi_", "masiva_ff_",
        "masiva_fr_", "masiva_loc_",
    )
    for key in list(st.session_state.keys()):
        key_text = str(key)
        if key_text.startswith(prefijos) and key_text.endswith(sufijo):
            del st.session_state[key]


# Reset del bloqueo cuando no hay archivo o cuando el usuario carga un archivo diferente.
if uploaded_file is None:
    st.session_state.guardando_masiva = False
    st.session_state.guardar_masiva_pendiente = False
    st.session_state.masiva_guardado_ok = False
    st.session_state.masiva_upload_id = ""
    st.session_state.masiva_registros_eliminados = {}
    st.session_state.masiva_bloque_actual = 0
    st.session_state.masiva_bloques_guardados = []
    st.session_state.masiva_resumen_bloques = []
    st.session_state.masiva_df_preparado = None
    st.session_state.masiva_errores_archivo = []
    st.session_state.masiva_df_upload_id = ""
    limpiar_estado_editor_masivo()
else:
    try:
        upload_size = getattr(uploaded_file, "size", None) or len(uploaded_file.getvalue())
    except Exception:
        upload_size = ""
    upload_id_actual = f"{getattr(uploaded_file, 'name', '')}_{upload_size}"
    if st.session_state.get("masiva_upload_id", "") != upload_id_actual:
        st.session_state.masiva_upload_id = upload_id_actual
        st.session_state.guardando_masiva = False
        st.session_state.guardar_masiva_pendiente = False
        st.session_state.masiva_guardado_ok = False
        st.session_state.masiva_registros_eliminados = {}
        st.session_state.masiva_bloque_actual = 0
        st.session_state.masiva_bloques_guardados = []
        st.session_state.masiva_resumen_bloques = []
        st.session_state.masiva_df_preparado = None
        st.session_state.masiva_errores_archivo = []
        st.session_state.masiva_df_upload_id = ""
        limpiar_estado_editor_masivo()

filas = []
errores_totales = []

if uploaded_file is not None:
    try:
        usuario_final = "" if nombre_usuario == "SELECCIONE" else nombre_usuario

        # El Excel se lee, normaliza y prepara una sola vez por archivo.
        # Los reruns provocados por selectboxes, botones o expansores reutilizan
        # el DataFrame guardado en session_state.
        if (
            st.session_state.get("masiva_df_preparado") is None
            or st.session_state.get("masiva_df_upload_id", "") != upload_id_actual
        ):
            with st.spinner("Procesando el archivo por primera vez..."):
                df_retrasos = leer_archivo_retrasos_xls(uploaded_file)
                df_preparado, errores_preparacion = preparar_editor_retrasos_desde_df(
                    df_retrasos,
                    "",
                    OPERADORES,
                )
                st.session_state.masiva_df_preparado = df_preparado
                st.session_state.masiva_errores_archivo = errores_preparacion
                st.session_state.masiva_df_upload_id = upload_id_actual

        # Se trabaja con una copia ligera para evitar modificar accidentalmente
        # el DataFrame maestro almacenado en session_state.
        df_editor_base = st.session_state.masiva_df_preparado.copy()
        errores_archivo = list(st.session_state.get("masiva_errores_archivo", []))

        # No volcamos aquí los errores por fila del archivo original.
        # Motivo: el usuario puede eliminar visualmente registros antes de guardar;
        # si agregamos estos errores en este punto, seguirían bloqueando el guardado
        # aunque la fila ya no forme parte de la revisión activa.
        # Las validaciones definitivas se ejecutan más abajo únicamente sobre
        # edited_records, es decir, sobre los registros que continúan visibles.
        if errores_archivo and df_editor_base.empty:
            errores_totales.extend(errores_archivo)
        else:
            # Los avisos asociados a filas concretas no se arrastran a la validación
            # del bloque, porque esas filas pueden haber sido eliminadas de la revisión.
            errores_totales = []

        if df_editor_base.empty:
            st.warning("El archivo no contiene registros válidos para preparar la carga.")
        else:
            st.success(f"Archivo leído correctamente. Registros detectados: {len(df_editor_base)}")
            st.info(
                "Completa los campos de clasificación por cada fila. El campo Momento Viaje se calcula automáticamente a partir de FECHA INICIO, FECHA DE FINALIZACIÓN y FECHA. "
                "El comentario es editable y opcional: puedes eliminar textos irrelevantes, depurar datos protegidos o dejarlo vacío antes de guardar. También puedes borrar registros completos de la revisión. El semáforo del encabezado indica si el resto de los datos obligatorios está completo."
            )

            columnas_visibles = [
                "fecha_inicio", "fecha_finalizacion", "fecha_registro", "localizador", "nombre_usuario", "operador_xls", "operador",
                "momento_viaje", "medio_contacto", "quien_contacta", "ciudad",
                "tipo_contacto", "area", "area_relacionada", "hotel", "tipo_traslado",
                "trayecto", "guia", "tipo_incidencia", "comentario", "resolucion",
                "monto", "resultado",
            ]

            df_editor_base = df_editor_base[[c for c in columnas_visibles if c in df_editor_base.columns]]

            total_registros_masiva = len(df_editor_base)
            total_bloques_masiva = max(
                1,
                (total_registros_masiva + TAMANO_BLOQUE_MASIVA - 1) // TAMANO_BLOQUE_MASIVA,
            )
            bloque_actual_masiva = int(st.session_state.get("masiva_bloque_actual", 0) or 0)

            if bloque_actual_masiva >= total_bloques_masiva:
                st.progress(1.0)
                st.success(
                    f"✅ Proceso finalizado. Se administraron {total_registros_masiva} "
                    f"registros en {total_bloques_masiva} bloques."
                )
                for resumen in st.session_state.get("masiva_resumen_bloques", []):
                    st.caption(resumen)
                st.stop()

            inicio_bloque_masiva = bloque_actual_masiva * TAMANO_BLOQUE_MASIVA
            fin_bloque_masiva = min(
                inicio_bloque_masiva + TAMANO_BLOQUE_MASIVA,
                total_registros_masiva,
            )
            df_bloque_masiva = df_editor_base.iloc[inicio_bloque_masiva:fin_bloque_masiva]

            numero_bloque_visible = bloque_actual_masiva + 1
            st.progress(inicio_bloque_masiva / total_registros_masiva)
            st.subheader(
                f"Bloque {numero_bloque_visible} de {total_bloques_masiva} · "
                f"Registros {inicio_bloque_masiva + 1}–{fin_bloque_masiva} "
                f"de {total_registros_masiva}"
            )
            st.caption(
                "Solo se muestran hasta 15 registros simultáneamente. "
                "Al guardar el bloque, la herramienta avanzará automáticamente."
            )

            momento_opts = ["", "Pre Viaje/Pre Tour", "En Ruta/On Route", "Post Viaje/Post Tour"]
            medio_opts = ["", "Email", "Llamada/Call", "WhatsApp", "LINE", "Telegram", "WeChat"]
            quien_opts = [
                "", "Operador/Tour Operator", "Minorista/Tour Agency", "Pasajero/Passenger",
                "Guía/Guide", "Transferista/Transfer", "Receptivo/Local Provider",
                "Depto. Interno/Internal Dept.", "Otro/Other"
            ]
            tipo_contacto_opts = ["", "Información/Information", "Reclamación/Complaint", "Otro/Other", "Cuestionario de Satisfacción"]
            area_opts = ["", "Hotel", "Guías/Guides", "Traslados/Transfers", "Generales/General", "Itinerario/Itinerary", "Otros/Other"]
            area_rel_opts = ["", "Hotel", "Guías/Guides", "Traslados/Transfers", "Generales/General"]
            tipo_traslado_opts = ["", "Llegada/Arrival", "Salida/Departure", "Llegada/Arrival-Pto", "Salida/Departure-Pto", "No Aplica/Does not Apply"]
            resultado_opts = [
                "", "ERROR EMV", "ERROR OPERADOR/AGENTE VIAJES", "ERROR CLIENTE", "ERROR RECEPTIVO",
                "FUERZA MAYOR", "ASISTENCIA / AYUDA", "MOTIVOS COMERCIALES",
                "QUEJA GENERALIZADA", "FELICITACIÓN"
            ]

            # st.data_editor no permite listas dependientes por fila.
            # Para filtrar Tipo de Incidencia según Área/Área Relacionada, usamos selectboxes por registro.
            edited_records = []

            def _campo_completo(valor):
                return bool(limpiar_valor_masivo(valor)) and limpiar_valor_masivo(valor) != "SELECCIONE"

            def _label_estado(label, valor, required=False):
                """Mantiene el label limpio. El semáforo se muestra solo en el encabezado del registro."""
                return label

            def _selectbox_masiva(label, options, default_value, key, required=False):
                opciones = list(options or [""])
                if "" not in opciones:
                    opciones = [""] + opciones
                default_value = limpiar_valor_masivo(default_value)
                if key not in st.session_state:
                    st.session_state[key] = default_value if default_value in opciones else ""
                elif st.session_state.get(key) not in opciones:
                    st.session_state[key] = ""
                label_render = _label_estado(label, st.session_state.get(key, ""), required)
                return st.selectbox(label_render, opciones, key=key)

            def _text_input_masiva(label, default_value, key, required=False):
                if key not in st.session_state:
                    st.session_state[key] = limpiar_valor_masivo(default_value)
                label_render = _label_estado(label, st.session_state.get(key, ""), required)
                return st.text_input(label_render, key=key)

            def _text_area_masiva(label, default_value, key, required=False, max_chars=500):
                if key not in st.session_state:
                    st.session_state[key] = limpiar_valor_masivo(default_value)
                label_render = _label_estado(label, st.session_state.get(key, ""), required)
                return st.text_area(label_render, key=key, max_chars=max_chars)

            def _valor_estado_masiva(idx, row, key_suffix, default_value=""):
                """Lee el valor vigente desde session_state o, si aún no existe el widget, desde la fila base."""
                key = f"masiva_{key_suffix}_{idx}"
                if key in st.session_state:
                    return limpiar_valor_masivo(st.session_state.get(key, ""))
                return limpiar_valor_masivo(default_value)

            def _pendientes_registro_masiva(idx, row):
                """Calcula pendientes del registro para mostrar un único semáforo en el encabezado."""
                pendientes = []

                operador = _valor_estado_masiva(idx, row, "op", row.get("operador", ""))
                momento = _valor_estado_masiva(idx, row, "momento", row.get("momento_viaje", ""))
                medio = _valor_estado_masiva(idx, row, "medio", row.get("medio_contacto", ""))
                quien = _valor_estado_masiva(idx, row, "quien", row.get("quien_contacta", ""))
                ciudad_v = _valor_estado_masiva(idx, row, "ciudad", row.get("ciudad", ""))
                tipo_contacto_v = _valor_estado_masiva(idx, row, "tipo_contacto", row.get("tipo_contacto", "Reclamación/Complaint"))

                for nombre_campo, valor_campo in {
                    "Operador": operador,
                    "Momento Viaje": momento,
                    "Medio de Contacto": medio,
                    "Quién Contacta": quien,
                    "Ciudad": ciudad_v,
                    "Tipo de Contacto": tipo_contacto_v,
                }.items():
                    if not _campo_completo(valor_campo):
                        pendientes.append(nombre_campo)

                resolucion_v = ""
                monto_v = ""

                if tipo_contacto_v == "Información/Information":
                    area_v = _valor_estado_masiva(idx, row, "area_info", row.get("area", ""))
                    resolucion_v = _valor_estado_masiva(idx, row, "resolucion_info", row.get("resolucion", ""))
                    if not _campo_completo(area_v): pendientes.append("Área Relacionada")
                    if not _campo_completo(resolucion_v): pendientes.append("Resolución")
                    if area_v == "Hotel":
                        hotel_v = _valor_estado_masiva(idx, row, "hotel_info", row.get("hotel", ""))
                        if not _campo_completo(hotel_v): pendientes.append("Hotel")
                    if area_v == "Traslados/Transfers":
                        tipo_traslado_v = _valor_estado_masiva(idx, row, "tipo_traslado_info", row.get("tipo_traslado", ""))
                        if not _campo_completo(tipo_traslado_v): pendientes.append("Tipo de Traslado")
                    monto_v = _valor_estado_masiva(idx, row, "monto_info", row.get("monto", ""))

                elif tipo_contacto_v == "Reclamación/Complaint":
                    area_rel_v = _valor_estado_masiva(idx, row, "area_reclamo", row.get("area_relacionada", ""))
                    tipo_inc_v = _valor_estado_masiva(idx, row, "tipo_incidencia_reclamo", row.get("tipo_incidencia", ""))
                    resolucion_v = _valor_estado_masiva(idx, row, "resolucion_reclamo", row.get("resolucion", ""))
                    resultado_v = _valor_estado_masiva(idx, row, "resultado_reclamo", row.get("resultado", ""))
                    if not _campo_completo(area_rel_v): pendientes.append("Área Relacionada")
                    if not _campo_completo(tipo_inc_v): pendientes.append("Tipo de Incidencia")
                    if not _campo_completo(resolucion_v): pendientes.append("Resolución")
                    if not _campo_completo(resultado_v): pendientes.append("Resultado")
                    if area_rel_v == "Hotel":
                        hotel_v = _valor_estado_masiva(idx, row, "hotel_reclamo", row.get("hotel", ""))
                        if not _campo_completo(hotel_v): pendientes.append("Hotel")
                    if area_rel_v == "Guías/Guides":
                        trayecto_v = _valor_estado_masiva(idx, row, "trayecto_guia", row.get("trayecto", ""))
                        guia_v = _valor_estado_masiva(idx, row, "guia_reclamo", row.get("guia", ""))
                        if not _campo_completo(trayecto_v): pendientes.append("Trayecto")
                        if not _campo_completo(guia_v): pendientes.append("Guía")
                    if area_rel_v == "Traslados/Transfers":
                        if tipo_inc_v.startswith("TRF"):
                            tipo_traslado_v = _valor_estado_masiva(idx, row, "tipo_traslado_reclamo", row.get("tipo_traslado", ""))
                            if not _campo_completo(tipo_traslado_v): pendientes.append("Tipo de Traslado")
                        if tipo_inc_v.startswith("BUS"):
                            trayecto_v = _valor_estado_masiva(idx, row, "trayecto_bus", row.get("trayecto", ""))
                            if not _campo_completo(trayecto_v): pendientes.append("Trayecto")
                    if area_rel_v == "Generales/General" and tipo_inc_v.startswith("Itinerario"):
                        trayecto_v = _valor_estado_masiva(idx, row, "trayecto_general", row.get("trayecto", ""))
                        if not _campo_completo(trayecto_v): pendientes.append("Trayecto")
                    monto_v = _valor_estado_masiva(idx, row, "monto_reclamo", row.get("monto", ""))

                elif tipo_contacto_v == "Otro/Other":
                    resolucion_v = _valor_estado_masiva(idx, row, "resolucion_otro", row.get("resolucion", ""))
                    if not _campo_completo(resolucion_v): pendientes.append("Resolución")
                    monto_v = _valor_estado_masiva(idx, row, "monto_otro", row.get("monto", ""))

                elif tipo_contacto_v == "Cuestionario de Satisfacción":
                    area_rel_v = _valor_estado_masiva(idx, row, "area_qs", row.get("area_relacionada", ""))
                    tipo_inc_v = _valor_estado_masiva(idx, row, "tipo_incidencia_qs", row.get("tipo_incidencia", ""))
                    if not _campo_completo(area_rel_v): pendientes.append("Área Relacionada")
                    if not _campo_completo(tipo_inc_v): pendientes.append("Tipo de Incidencia")
                    if area_rel_v == "Hotel":
                        hotel_v = _valor_estado_masiva(idx, row, "hotel_qs", row.get("hotel", ""))
                        if not _campo_completo(hotel_v): pendientes.append("Hotel")
                    if area_rel_v == "Guías/Guides":
                        guia_v = _valor_estado_masiva(idx, row, "guia_qs", row.get("guia", ""))
                        if not _campo_completo(guia_v): pendientes.append("Guía")

                if resolucion_v.startswith("Reembolso") or resolucion_v == "Compensación/Compensation":
                    if not _campo_completo(monto_v): pendientes.append("Monto")

                return pendientes

            def _limpiar_comentario_masivo(comentario_key: str):
                """Limpia el comentario mediante callback, antes del nuevo render de Streamlit."""
                st.session_state[comentario_key] = ""

            for posicion_bloque, (idx_masiva, row) in enumerate(df_bloque_masiva.iterrows()):
                eliminado_key = f"{st.session_state.get('masiva_upload_id', '')}_{idx_masiva}"
                if st.session_state.get("masiva_registros_eliminados", {}).get(eliminado_key):
                    continue

                loc_row = limpiar_valor_masivo(row.get("localizador", ""))
                comentario_row = limpiar_valor_masivo(row.get("comentario", ""))
                pendientes_header = _pendientes_registro_masiva(idx_masiva, row)
                if pendientes_header:
                    titulo_expander = f"🟡 {idx_masiva + 1}. {loc_row}"
                else:
                    titulo_expander = f"🟢 {idx_masiva + 1}. {loc_row}"

                with st.expander(titulo_expander, expanded=(posicion_bloque < 0)):
                    url_reserva_masiva = (
                        "https://www.europamundo-online.com/reservas/"
                        f"buscarreserva2.asp?coreserva={loc_row}"
                    )
                    c_accion1, c_accion2, c_accion3 = st.columns([1, 1, 5])
                    with c_accion1:
                        st.link_button(
                            "🔎 Ver Reserva",
                            url_reserva_masiva,
                            use_container_width=True,
                            disabled=not bool(loc_row),
                        )
                    with c_accion2:
                        if st.button(
                            "🗑️ Eliminar registro",
                            key=f"masiva_eliminar_{idx_masiva}",
                            use_container_width=True,
                            help="Quita este registro de la revisión actual. No se guardará en Google Sheets.",
                        ):
                            st.session_state.setdefault("masiva_registros_eliminados", {})[eliminado_key] = True
                            limpiar_estado_registro_masivo(idx_masiva)
                            st.rerun()


                    c_base1, c_base2, c_base3, c_base4 = st.columns(4)
                    with c_base1:
                        st.text_input("Fecha Inicio", value=row.get("fecha_inicio", ""), disabled=True, key=f"masiva_fi_{idx_masiva}")
                    with c_base2:
                        st.text_input("Fecha Finalización", value=row.get("fecha_finalizacion", ""), disabled=True, key=f"masiva_ff_{idx_masiva}")
                    with c_base3:
                        st.text_input("Fecha Registro", value=row.get("fecha_registro", ""), disabled=True, key=f"masiva_fr_{idx_masiva}")
                    with c_base4:
                        st.text_input("Localizador", value=loc_row, disabled=True, key=f"masiva_loc_{idx_masiva}")

                    c_base5, c_base6 = st.columns([1, 2])
                    operador_xls_row = limpiar_valor_masivo(row.get("operador_xls", row.get("operador", "")))
                    operadores_filtrados = filtrar_operadores_por_nombre_masivo(OPERADORES, operador_xls_row)
                    opciones_operador_masiva = operadores_filtrados if operadores_filtrados else (OPERADORES or [])
                    with c_base5:
                        st.caption(f"Operador XLS: {operador_xls_row or 'Sin informar'}")
                        if not operadores_filtrados and operador_xls_row:
                            st.warning("No se encontró coincidencia exacta por nombre comercial; selecciona manualmente.")
                        operador_catalogo = _selectbox_masiva(
                            "Operador",
                            opciones_operador_masiva,
                            row.get("operador", ""),
                            f"masiva_op_{idx_masiva}",
                            True
                        )
                    with c_base6:
                        comentario_key = f"masiva_com_{idx_masiva}"
                        if comentario_key not in st.session_state:
                            st.session_state[comentario_key] = comentario_row

                        c_com1, c_com2 = st.columns([4, 1])
                        with c_com1:
                            comentario_editado = st.text_area(
                                "Comentario editable",
                                key=comentario_key,
                                height=110,
                                help="Puedes modificar, resumir o eliminar información sensible antes de guardar.",
                            )
                        with c_com2:
                            st.write("")
                            st.write("")
                            st.button(
                                "🧹 Limpiar",
                                key=f"masiva_limpiar_com_{idx_masiva}",
                                on_click=_limpiar_comentario_masivo,
                                args=(comentario_key,),
                            )

                    c1, c2, c3, c4 = st.columns(4)
                    with c1:
                        momento_viaje = _selectbox_masiva("Momento Viaje", momento_opts, row.get("momento_viaje", ""), f"masiva_momento_{idx_masiva}", True)
                    with c2:
                        medio_contacto = _selectbox_masiva("Medio Contacto", medio_opts, row.get("medio_contacto", ""), f"masiva_medio_{idx_masiva}", True)
                    with c3:
                        quien_contacta = _selectbox_masiva("Quién Contacta", quien_opts, row.get("quien_contacta", ""), f"masiva_quien_{idx_masiva}", True)
                    with c4:
                        ciudad = _selectbox_masiva("Ciudad", [""] + CIUDADES, row.get("ciudad", ""), f"masiva_ciudad_{idx_masiva}", True)

                    c5, c6 = st.columns([1, 2])
                    with c5:
                        tipo_contacto = _selectbox_masiva(
                            "Tipo Contacto",
                            tipo_contacto_opts,
                            row.get("tipo_contacto", "Reclamación/Complaint"),
                            f"masiva_tipo_contacto_{idx_masiva}",
                            True
                        )

                    # Valores base. Solo se completan los campos que correspondan a la misma lógica de la carga manual.
                    area = ""
                    area_relacionada = ""
                    hotel = ""
                    tipo_traslado = ""
                    trayecto = ""
                    guia = ""
                    tipo_incidencia = ""
                    resolucion = ""
                    monto = ""
                    resultado = ""

                    hoteles_ciudad = filtrar_hoteles_por_ciudad(HOTELES, ciudad) if ciudad else HOTELES
                    trayectos_ciudad = filtrar_trayectos_por_ciudad(TRAYECTOS, ciudad) if ciudad else TRAYECTOS

                    if tipo_contacto == "Información/Information":
                        area_info_opts = [
                            "", "Traslados/Transfers", "Hotel", "Seguro/Insurance", "Itinerario/Itinerary",
                            "Equipaje/Luggage", "Felicitación Circuito", "Info Guía/Guide Info",
                            "Punto Encuentro/Meeting Point", "Comercial/Commercial", "Enfermedad/Sickness",
                            "Opcionales/Optional Tours", "Otros/Other"
                        ]
                        with c6:
                            area = _selectbox_masiva("Área Relacionada", area_info_opts, row.get("area", ""), f"masiva_area_info_{idx_masiva}", True)

                        if area == "Hotel":
                            hotel = _selectbox_masiva("Hotel", [""] + hoteles_ciudad, row.get("hotel", ""), f"masiva_hotel_info_{idx_masiva}", True)
                        elif area == "Traslados/Transfers":
                            tipo_traslado = _selectbox_masiva(
                                "Tipo de Traslado",
                                ["", "Llegada/Arrival", "Salida/Departure", "Llegada/Arrival-Pto", "Salida/Departure-Pto", "NO APLICA / DOESN´T APPLY"],
                                row.get("tipo_traslado", ""),
                                f"masiva_tipo_traslado_info_{idx_masiva}",
                                True
                            )

                        resolucion = _selectbox_masiva("Resolución", RESOLUCIONES, row.get("resolucion", ""), f"masiva_resolucion_info_{idx_masiva}", True)
                        if resolucion.startswith("Reembolso") or resolucion == "Compensación/Compensation":
                            monto = _text_input_masiva("Monto compensación o tipo de compensación", row.get("monto", ""), f"masiva_monto_info_{idx_masiva}", required=True)

                    elif tipo_contacto == "Reclamación/Complaint":
                        with c6:
                            area_relacionada = _selectbox_masiva(
                                "Área Relacionada",
                                ["", "Hotel", "Guías/Guides", "Traslados/Transfers", "Generales/General"],
                                row.get("area_relacionada", ""),
                                f"masiva_area_reclamo_{idx_masiva}",
                                True
                            )

                        tipo_incidencia_opts = obtener_tipos_incidencia_por_area_masiva(area_relacionada)
                        default_tipo = row.get("tipo_incidencia", "")
                        if limpiar_valor_masivo(default_tipo) not in tipo_incidencia_opts:
                            default_tipo = ""
                        tipo_incidencia = _selectbox_masiva(
                            "Tipo de Incidencia",
                            tipo_incidencia_opts,
                            default_tipo,
                            f"masiva_tipo_incidencia_reclamo_{idx_masiva}",
                            True
                        )

                        if area_relacionada == "Hotel":
                            hotel = _selectbox_masiva("Hotel", [""] + hoteles_ciudad, row.get("hotel", ""), f"masiva_hotel_reclamo_{idx_masiva}", True)
                        elif area_relacionada == "Guías/Guides":
                            c_g1, c_g2 = st.columns(2)
                            with c_g1:
                                trayecto = _selectbox_masiva("Trayecto", [""] + trayectos_ciudad, row.get("trayecto", ""), f"masiva_trayecto_guia_{idx_masiva}", True)
                            with c_g2:
                                guia = _selectbox_masiva("Nombre del Guía", [""] + GUIAS, row.get("guia", ""), f"masiva_guia_reclamo_{idx_masiva}", True)
                        elif area_relacionada == "Traslados/Transfers":
                            if tipo_incidencia.startswith("TRF"):
                                tipo_traslado = _selectbox_masiva(
                                    "Tipo de Traslado",
                                    ["", "Llegada/Arrival", "Salida/Departure", "Llegada/Arrival-Pto", "Salida/Departure-Pto", "No Aplica/Does not Apply"],
                                    row.get("tipo_traslado", ""),
                                    f"masiva_tipo_traslado_reclamo_{idx_masiva}",
                                    True
                                )
                            elif tipo_incidencia.startswith("BUS"):
                                trayecto = _selectbox_masiva("Trayecto", [""] + trayectos_ciudad, row.get("trayecto", ""), f"masiva_trayecto_bus_{idx_masiva}", True)
                        elif area_relacionada == "Generales/General" and tipo_incidencia.startswith("Itinerario"):
                            trayecto = _selectbox_masiva("Trayecto", [""] + trayectos_ciudad, row.get("trayecto", ""), f"masiva_trayecto_general_{idx_masiva}", True)

                        c_r1, c_r2, c_r3 = st.columns(3)
                        with c_r1:
                            resolucion = _selectbox_masiva("Resolución", RESOLUCIONES, row.get("resolucion", ""), f"masiva_resolucion_reclamo_{idx_masiva}", True)
                        with c_r2:
                            if resolucion.startswith("Reembolso") or resolucion == "Compensación/Compensation":
                                monto = _text_input_masiva("Monto compensación o tipo de compensación", row.get("monto", ""), f"masiva_monto_reclamo_{idx_masiva}", required=True)
                        with c_r3:
                            resultado = _selectbox_masiva("Resultado", resultado_opts, row.get("resultado", ""), f"masiva_resultado_reclamo_{idx_masiva}", True)

                    elif tipo_contacto == "Otro/Other":
                        resolucion = _selectbox_masiva("Resolución Otros", RESOLUCIONES, row.get("resolucion", ""), f"masiva_resolucion_otro_{idx_masiva}", True)
                        if resolucion.startswith("Reembolso") or resolucion == "Compensación/Compensation":
                            monto = _text_input_masiva("Monto compensación o tipo de compensación", row.get("monto", ""), f"masiva_monto_otro_{idx_masiva}", required=True)

                    elif tipo_contacto == "Cuestionario de Satisfacción":
                        with c6:
                            area_relacionada = _selectbox_masiva(
                                "Área Relacionada",
                                ["", "Hotel", "Guías/Guides", "Traslados/Transfers", "Generales/General"],
                                row.get("area_relacionada", ""),
                                f"masiva_area_qs_{idx_masiva}",
                                True
                            )

                        tipo_incidencia_opts = obtener_tipos_incidencia_por_area_masiva(area_relacionada)
                        default_tipo = row.get("tipo_incidencia", "")
                        if limpiar_valor_masivo(default_tipo) not in tipo_incidencia_opts:
                            default_tipo = ""
                        tipo_incidencia = _selectbox_masiva(
                            "Tipo de Incidencia",
                            tipo_incidencia_opts,
                            default_tipo,
                            f"masiva_tipo_incidencia_qs_{idx_masiva}",
                            True
                        )

                        if area_relacionada == "Hotel":
                            hotel = _selectbox_masiva("Hotel", [""] + hoteles_ciudad, row.get("hotel", ""), f"masiva_hotel_qs_{idx_masiva}", True)
                        elif area_relacionada == "Guías/Guides":
                            guia = _selectbox_masiva("Nombre del Guía", [""] + GUIAS, row.get("guia", ""), f"masiva_guia_qs_{idx_masiva}", True)

                    edited_records.append({
                        "fecha_inicio": row.get("fecha_inicio", ""),
                        "fecha_registro": row.get("fecha_registro", ""),
                        "momento_viaje": momento_viaje,
                        "localizador": loc_row,
                        "nombre_usuario": usuario_final,
                        "operador": operador_catalogo,
                        "medio_contacto": medio_contacto,
                        "quien_contacta": quien_contacta,
                        "ciudad": ciudad,
                        "tipo_contacto": tipo_contacto,
                        "area": area,
                        "area_relacionada": area_relacionada,
                        "hotel": hotel,
                        "tipo_traslado": tipo_traslado,
                        "trayecto": trayecto,
                        "guia": guia,
                        "tipo_incidencia": tipo_incidencia,
                        "comentario": limpiar_valor_masivo(comentario_editado),
                        "resolucion": resolucion,
                        "monto": monto,
                        "resultado": resultado,
                        "incidencia_id": "",
                    })

            import pandas as pd

            # La validación de cada bloque parte de cero y solo considera
            # los registros que continúan visibles.
            errores_totales = []
            edited_df = pd.DataFrame(edited_records)
            total_visibles_bloque = len(edited_records)
            total_original_bloque = len(df_bloque_masiva)
            total_eliminados_bloque = total_original_bloque - total_visibles_bloque

            if nombre_usuario == "SELECCIONE":
                errores_totales.append("Debe seleccionar el Usuario común del lote.")

            if total_visibles_bloque > 0:
                filas, errores_editor = construir_filas_retrasos_desde_editor(
                    edited_df,
                    usuario_final,
                )
                errores_totales.extend(errores_editor)
            else:
                filas = []
                errores_editor = []

            # Coherencia entre semáforo y validación final:
            # si todos los registros visibles están completos según la misma función,
            # descartamos cualquier error residual que no pertenezca al bloque actual.
            pendientes_visibles_bloque = []
            for idx_visible, row_visible in df_bloque_masiva.iterrows():
                eliminado_visible_key = (
                    f"{st.session_state.get('masiva_upload_id', '')}_{idx_visible}"
                )
                if st.session_state.get("masiva_registros_eliminados", {}).get(
                    eliminado_visible_key
                ):
                    continue
                pendientes_visibles_bloque.extend(
                    _pendientes_registro_masiva(idx_visible, row_visible)
                )

            if not pendientes_visibles_bloque and total_visibles_bloque > 0:
                errores_totales = [
                    error for error in errores_totales
                    if "Usuario común del lote" in str(error)
                ]


    except Exception as e:
        errores_totales.append(f"No se pudo leer el archivo cargado: {e}")
        st.error(f"No se pudo leer el archivo cargado: {e}")

if errores_totales:
    if nombre_usuario == "SELECCIONE":
        st.info(
            "Selecciona el Usuario que aplica al lote antes de guardar o avanzar al siguiente bloque."
        )
    else:
        st.info(
            "Hay campos pendientes en alguno de los registros visibles de este bloque. "
            "Revísalos mediante los indicadores 🟡. Los eliminados no se validan ni se guardan."
        )

def activar_guardado_masivo():
    st.session_state.guardando_masiva = True
    st.session_state.guardar_masiva_pendiente = True

if uploaded_file is not None and "df_bloque_masiva" in locals():
    bloque_sin_filas = len(edited_records) == 0

    if bloque_sin_filas:
        texto_boton_bloque = (
            f"➡️ Continuar al bloque {numero_bloque_visible + 1}"
            if numero_bloque_visible < total_bloques_masiva
            else "✅ Finalizar proceso"
        )
    else:
        texto_boton_bloque = (
            f"✅ Guardar bloque {numero_bloque_visible} y continuar"
            if numero_bloque_visible < total_bloques_masiva
            else f"✅ Guardar bloque {numero_bloque_visible} y finalizar"
        )

    if st.session_state.get("guardando_masiva", False):
        st.info("Procesando el bloque actual. Por favor, espera un momento...")

    st.button(
        texto_boton_bloque,
        disabled=(
            bool(errores_totales)
            or st.session_state.get("guardando_masiva", False)
        ),
        on_click=activar_guardado_masivo,
        key=f"btn_guardar_bloque_{bloque_actual_masiva}",
    )

if st.session_state.get("guardar_masiva_pendiente", False):
    st.session_state.guardar_masiva_pendiente = False

    if nombre_usuario == "SELECCIONE":
        st.session_state.guardando_masiva = False
        st.error("Debe seleccionar el Usuario común del lote antes de guardar o continuar.")
        st.stop()

    if uploaded_file is None or bool(errores_totales):
        st.session_state.guardando_masiva = False
        st.error("No es posible continuar porque el bloque actual tiene validaciones pendientes.")
        st.stop()

    try:
        cantidad_guardada = 0

        if filas:
            with st.spinner(
                f"Guardando bloque {numero_bloque_visible} de {total_bloques_masiva}..."
            ):
                guardado_correcto = guardar_lote_google_sheets_seguro(filas)

            if not guardado_correcto:
                st.session_state.guardando_masiva = False
                st.error(
                    "El bloque no pudo guardarse. No se avanzará al siguiente para evitar pérdida de datos."
                )
                st.stop()

            cantidad_guardada = len(filas)

        resumen_bloque = (
            f"✅ Bloque {numero_bloque_visible}: {cantidad_guardada} guardados · "
            f"{total_eliminados_bloque} eliminados"
        )
        st.session_state.setdefault("masiva_resumen_bloques", []).append(resumen_bloque)
        st.session_state.setdefault("masiva_bloques_guardados", []).append(bloque_actual_masiva)

        for indice_finalizado in range(inicio_bloque_masiva, fin_bloque_masiva):
            limpiar_estado_registro_masivo(indice_finalizado)

        # Los catálogos no cambian al guardar un bloque.
        # Se conserva la caché para que la transición al siguiente sea inmediata.
        st.session_state.guardando_masiva = False
        st.session_state.masiva_bloque_actual = bloque_actual_masiva + 1
        st.rerun()

    except Exception as e:
        st.session_state.guardando_masiva = False
        st.error(f"❌ Error al procesar el bloque actual: {e}")

print("[BOOT 08] Ejecución del script completada", flush=True)
