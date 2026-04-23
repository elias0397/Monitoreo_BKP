cd # -*- coding: utf-8 -*-
"""
extractor.py — Extractor de reportes de backup desde Outlook (MAPI)
Descarga correos de Arcserve y Veeam como .msg filtrados por los últimos 2 días.
"""

import os
import re
import sys
from datetime import datetime, timedelta

import win32com.client
import pywintypes

# ──────────────────────────────────────────────
# Configuración
# ──────────────────────────────────────────────
DESTINO = r"C:\bkp"
DIAS_ATRAS = 2

# Patrones de clasificación (compilados una sola vez)
PATRONES = {
    "Arcserve":        re.compile(r"(?i)Arcserve\s+UDP.*(Copia de Seguridad|Alerta)"),
    "Veeam_Explicito": re.compile(r"(?i)\[(Success|Failed|Warning)\].*Backup"),
    "Veeam_Objetos":   re.compile(r"(?i)\[(Success|Failed|Warning)\].*\(\d+\s+objects\)"),
}

# Regex para extraer estado desde corchetes [ ]
RE_ESTADO = re.compile(r"\[(Success|Failed|Warning)\]", re.IGNORECASE)

# Caracteres prohibidos en nombres de archivo de Windows
RE_CHARS_INVALIDOS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


# ──────────────────────────────────────────────
# Funciones auxiliares
# ──────────────────────────────────────────────
def asegurar_directorio(ruta: str) -> None:
    """Crea el directorio destino si no existe."""
    os.makedirs(ruta, exist_ok=True)


def coincide_patron(asunto: str) -> str | None:
    """Devuelve el nombre del patrón que coincide, o None."""
    for nombre, patron in PATRONES.items():
        if patron.search(asunto):
            return nombre
    return None


def extraer_estado(asunto: str) -> str:
    """Extrae el estado (Success/Failed/Warning) del asunto; 'Desconocido' si no existe."""
    match = RE_ESTADO.search(asunto)
    return match.group(1).capitalize() if match else "Desconocido"


def limpiar_asunto(asunto: str, max_len: int = 120) -> str:
    """Remueve caracteres inválidos y limita longitud para el nombre del archivo."""
    limpio = RE_CHARS_INVALIDOS.sub("_", asunto)
    limpio = re.sub(r"_+", "_", limpio).strip("_ ")
    return limpio[:max_len]


def construir_nombre_archivo(estado: str, fecha: datetime, asunto: str) -> str:
    """Genera el nombre de archivo con formato [ESTADO]_YYYYMMDD_HHMM_AsuntoLimpio.msg"""
    marca = fecha.strftime("%Y%m%d_%H%M")
    asunto_limpio = limpiar_asunto(asunto)
    return f"[{estado}]_{marca}_{asunto_limpio}.msg"


def construir_filtro_fecha(dias: int) -> str:
    """Construye la cadena de filtro DASL/Restrict para los últimos N días."""
    fecha_inicio = (datetime.now() - timedelta(days=dias)).strftime("%m/%d/%Y %H:%M %p")
    return f"[ReceivedTime] >= '{fecha_inicio}'"


# ──────────────────────────────────────────────
# Proceso principal
# ──────────────────────────────────────────────
def main() -> None:
    asegurar_directorio(DESTINO)

    # Conexión a Outlook vía MAPI
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mapi = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"[ERROR] No se pudo conectar a Outlook: {e}")
        sys.exit(1)

    # Acceso al Inbox (olFolderInbox = 6)
    inbox = mapi.GetDefaultFolder(6)
    mensajes = inbox.Items

    # Filtrar por fecha (últimos N días)
    filtro = construir_filtro_fecha(DIAS_ATRAS)
    mensajes_filtrados = mensajes.Restrict(filtro)
    print(f"[INFO] Filtro aplicado: {filtro}")
    print(f"[INFO] Mensajes en rango de fecha: {mensajes_filtrados.Count}")

    guardados = 0
    omitidos = 0

    for i in range(mensajes_filtrados.Count, 0, -1):
        try:
            msg = mensajes_filtrados.Item(i)
        except pywintypes.com_error:
            # Elemento inaccesible (calendario, nota, etc.)
            continue

        # Solo procesar objetos MailItem (Class = 43)
        if msg.Class != 43:
            continue

        asunto = msg.Subject or "(Sin asunto)"
        categoria = coincide_patron(asunto)

        if categoria is None:
            omitidos += 1
            continue

        # Extraer estado y fecha de recepción
        estado = extraer_estado(asunto)
        try:
            fecha_recibido = datetime(
                msg.ReceivedTime.year,
                msg.ReceivedTime.month,
                msg.ReceivedTime.day,
                msg.ReceivedTime.hour,
                msg.ReceivedTime.minute,
            )
        except Exception:
            fecha_recibido = datetime.now()

        nombre_archivo = construir_nombre_archivo(estado, fecha_recibido, asunto)
        ruta_completa = os.path.join(DESTINO, nombre_archivo)

        # Evitar sobrescribir archivos existentes
        if os.path.exists(ruta_completa):
            print(f"[SKIP] Ya existe: {nombre_archivo}")
            continue

        # Guardar como .msg (olMSG = 3)
        try:
            msg.SaveAs(ruta_completa, 3)
            guardados += 1
            print(f"[OK]   [{categoria}] {nombre_archivo}")
        except Exception as e:
            print(f"[ERR]  No se pudo guardar «{asunto[:60]}»: {e}")

    print(f"\n{'='*60}")
    print(f"[RESUMEN] Guardados: {guardados} | Omitidos (sin coincidencia): {omitidos}")
    print(f"[RESUMEN] Destino: {DESTINO}")


if __name__ == "__main__":
    main()
