"""
Sistema de gestión de vencimientos de cartera de pólizas de seguros
====================================================================

Pensado para correr de forma DESATENDIDA (cron, GitHub Actions, etc.):
no pide datos por teclado salvo que falten variables de entorno y se
ejecute en una terminal interactiva.

Variables de entorno usadas (ver .env.example):
    EMAIL_EMISOR      Cuenta de Gmail que envía los correos
    EMAIL_PASSWORD    Contraseña de aplicación de Gmail (16 dígitos)
    CARTERA_PATH      Ruta al CSV/XLSX con la cartera (default: data/cartera.csv)
    DIAS_AVISO        Lista de umbrales separados por coma (default: 30,15,5,1)
"""

import argparse
import logging
import os
import re
import smtplib
import sys
import unicodedata
from datetime import datetime, date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd

# ----------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------

DIAS_AVISO_DEFAULT = [30, 15, 5, 1]

ASUNTO_CORREO = "📋 Recordatorio: Su póliza de seguro vence en {dias} día(s)"

PLANTILLA_CORREO = """
Estimado/a {cliente},

Le recordamos que su póliza de seguro N° {poliza} vence el día {vencimiento}
(en {dias} día(s)).

Detalles de su póliza:
- Número de póliza: {poliza}
- Fecha de vencimiento: {vencimiento}
- Prima anual: ${prima:,.0f}
- Aseguradora: {aseguradora}

Por favor, contacte a su agente para coordinar la renovación.

¡Gracias por confiar en nosotros!

Atentamente,
Su agente de seguros
"""

# Alias aceptados por campo: el sistema detecta solo la columna real,
# sin necesidad de editar el script cada vez que cambie el archivo de origen.
ALIASES_COLUMNAS = {
    "cliente": ["cliente", "nombre", "nombre_cliente", "asegurado", "nombre_asegurado"],
    "email": ["email", "correo", "correo_electronico", "mail", "e_mail", "e-mail"],
    "poliza": ["poliza", "no_poliza", "numero_poliza", "n_poliza", "poliza_no", "póliza"],
    "vencimiento": [
        "vencimiento", "fecha_vencimiento", "fecha_vto", "fecha_de_vencimiento",
        "vence", "fecha_vence",
    ],
    "prima": ["prima", "prima_anual", "valor_prima", "monto", "valor"],
    "aseguradora": ["aseguradora", "compania", "compañia", "aseguradora_nombre"],
}

CAMPOS_OBLIGATORIOS = ["cliente", "email", "poliza", "vencimiento", "prima"]

LOG_DIR = "logs"
HISTORIAL_PATH = os.path.join(LOG_DIR, "historial_envios.csv")
EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


# ----------------------------------------------------------------------
# Logging
# ----------------------------------------------------------------------

def configurar_logging():
    os.makedirs(LOG_DIR, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(os.path.join(LOG_DIR, "sistema.log"), encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger("gestion_cartera")


log = configurar_logging()


# ----------------------------------------------------------------------
# Lectura y normalización de la cartera
# ----------------------------------------------------------------------

def _normalizar(texto):
    """minúsculas, sin tildes, espacios -> guion bajo"""
    texto = str(texto).strip().lower()
    texto = "".join(
        c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"[\s\-]+", "_", texto)


def detectar_columnas(df):
    """
    Detecta automáticamente qué columna del archivo corresponde a cada
    campo del sistema, comparando contra ALIASES_COLUMNAS.
    Lanza ValueError con mensaje claro si falta algún campo obligatorio.
    """
    columnas_normalizadas = {_normalizar(c): c for c in df.columns}
    mapeo = {}

    for campo, alias in ALIASES_COLUMNAS.items():
        encontrado = None
        for posible in alias:
            if posible in columnas_normalizadas:
                encontrado = columnas_normalizadas[posible]
                break
        mapeo[campo] = encontrado

    faltantes = [c for c in CAMPOS_OBLIGATORIOS if mapeo.get(c) is None]
    if faltantes:
        disponibles = ", ".join(df.columns.astype(str))
        raise ValueError(
            f"No se pudieron identificar estas columnas obligatorias: {faltantes}.\n"
            f"Columnas disponibles en el archivo: {disponibles}\n"
            f"Si tu archivo usa nombres distintos, agrégalos a ALIASES_COLUMNAS "
            f"en gestion_cartera.py."
        )
    return mapeo


def leer_cartera(ruta_archivo):
    """
    Lee la cartera desde CSV o Excel (.xlsx/.xls), detecta columnas
    automáticamente y normaliza tipos de datos.
    """
    if not os.path.exists(ruta_archivo):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_archivo}")

    ext = os.path.splitext(ruta_archivo)[1].lower()

    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(ruta_archivo)
    else:
        # Intentos sucesivos: utf-8 -> latin-1 -> separador ';'
        intentos = [
            dict(encoding="utf-8"),
            dict(encoding="latin-1"),
            dict(encoding="utf-8", sep=";"),
            dict(encoding="latin-1", sep=";"),
        ]
        df = None
        ultimo_error = None
        for kwargs in intentos:
            try:
                df = pd.read_csv(ruta_archivo, **kwargs)
                if df.shape[1] > 1:  # si solo hay 1 columna, probablemente el separador está mal
                    break
            except Exception as e:  # noqa: BLE001
                ultimo_error = e
        if df is None:
            raise ValueError(f"No se pudo leer el CSV con ningún encoding/separador conocido: {ultimo_error}")

    log.info("Archivo leído: %s filas, columnas: %s", len(df), list(df.columns))

    mapeo = detectar_columnas(df)
    log.info("Columnas detectadas: %s", mapeo)

    df_norm = pd.DataFrame()
    for campo in CAMPOS_OBLIGATORIOS:
        df_norm[campo] = df[mapeo[campo]]
    df_norm["aseguradora"] = df[mapeo["aseguradora"]] if mapeo.get("aseguradora") else "No especificada"

    # Tipos: primero se intenta detección estándar (cubre ISO yyyy-mm-dd),
    # y solo para lo que falle se reintenta asumiendo formato día/mes/año
    # (evita que "2026-07-15" se interprete mal por dayfirst).
    fechas = pd.to_datetime(df_norm["vencimiento"], errors="coerce")
    pendientes = fechas.isna()
    if pendientes.any():
        fechas.loc[pendientes] = pd.to_datetime(
            df_norm.loc[pendientes, "vencimiento"], errors="coerce", dayfirst=True
        )
    df_norm["vencimiento"] = fechas
    filas_fecha_invalida = df_norm["vencimiento"].isna().sum()
    if filas_fecha_invalida:
        log.warning("%s filas con fecha de vencimiento inválida fueron descartadas", filas_fecha_invalida)
    df_norm = df_norm.dropna(subset=["vencimiento"])

    if df_norm["prima"].dtype == object:
        df_norm["prima"] = (
            df_norm["prima"].astype(str).str.replace(r"[$,\s]", "", regex=True)
        )
    df_norm["prima"] = pd.to_numeric(df_norm["prima"], errors="coerce").fillna(0)

    # Validación de emails
    emails_invalidos = ~df_norm["email"].astype(str).str.match(EMAIL_REGEX)
    if emails_invalidos.any():
        log.warning(
            "%s filas con email inválido serán omitidas: %s",
            emails_invalidos.sum(),
            df_norm.loc[emails_invalidos, "poliza"].tolist(),
        )
    df_norm = df_norm[~emails_invalidos]

    log.info("Cartera válida: %s pólizas", len(df_norm))
    return df_norm


# ----------------------------------------------------------------------
# Historial / anti-duplicados
# ----------------------------------------------------------------------

def cargar_historial():
    if os.path.exists(HISTORIAL_PATH):
        try:
            return pd.read_csv(HISTORIAL_PATH, dtype={"poliza": str})
        except Exception as e:  # noqa: BLE001
            log.warning("No se pudo leer el historial existente (%s); se asume vacío", e)
    return pd.DataFrame(columns=["cliente", "email", "poliza", "vencimiento", "dias_aviso", "fecha_envio", "exito"])


def ya_notificado(historial, poliza, dias_aviso):
    """True si ya se notificó exitosamente este umbral para esta póliza."""
    if historial.empty:
        return False
    coincide = (
        (historial["poliza"].astype(str) == str(poliza))
        & (historial["dias_aviso"] == dias_aviso)
        & (historial["exito"] == "Sí")
    )
    return coincide.any()


def registrar_envio(fila, dias_aviso, exito):
    os.makedirs(LOG_DIR, exist_ok=True)
    registro = pd.DataFrame([{
        "cliente": fila["cliente"],
        "email": fila["email"],
        "poliza": fila["poliza"],
        "vencimiento": fila["vencimiento"].strftime("%Y-%m-%d"),
        "dias_aviso": dias_aviso,
        "fecha_envio": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "exito": "Sí" if exito else "No",
    }])
    header = not os.path.exists(HISTORIAL_PATH)
    registro.to_csv(HISTORIAL_PATH, mode="a", header=header, index=False)


# ----------------------------------------------------------------------
# Filtro de pólizas a notificar
# ----------------------------------------------------------------------

def filtrar_polizas_pendientes(df, historial, umbrales):
    hoy = date.today()
    pendientes = []
    for _, fila in df.iterrows():
        dias_restantes = (fila["vencimiento"].date() - hoy).days
        if dias_restantes in umbrales and not ya_notificado(historial, fila["poliza"], dias_restantes):
            pendientes.append((fila, dias_restantes))
    return pendientes


# ----------------------------------------------------------------------
# Envío de correos
# ----------------------------------------------------------------------

def generar_mensaje(fila, dias_aviso):
    return PLANTILLA_CORREO.format(
        cliente=fila["cliente"],
        poliza=fila["poliza"],
        vencimiento=fila["vencimiento"].strftime("%d/%m/%Y"),
        prima=fila["prima"],
        aseguradora=fila.get("aseguradora", "No especificada"),
        dias=dias_aviso,
    )


def enviar_correo(emisor, password, destinatario, asunto, cuerpo):
    try:
        msg = MIMEMultipart()
        msg["From"] = emisor
        msg["To"] = destinatario
        msg["Subject"] = asunto
        msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as server:
            server.starttls()
            server.login(emisor, password)
            server.send_message(msg)
        return True
    except Exception as e:  # noqa: BLE001
        log.error("Error SMTP enviando a %s: %s", destinatario, e)
        return False


def procesar_vencimientos(df, emisor, password, umbrales):
    historial = cargar_historial()
    pendientes = filtrar_polizas_pendientes(df, historial, umbrales)

    log.info("Pólizas pendientes de notificar: %s", len(pendientes))
    if not pendientes:
        log.info("No hay pólizas que notificar hoy (umbrales: %s)", umbrales)
        return

    enviados, fallidos = 0, 0
    for i, (fila, dias) in enumerate(pendientes, 1):
        log.info("[%s/%s] %s (póliza %s) - vence en %s día(s)", i, len(pendientes), fila["cliente"], fila["poliza"], dias)
        asunto = ASUNTO_CORREO.format(dias=dias)
        cuerpo = generar_mensaje(fila, dias)
        exito = enviar_correo(emisor, password, fila["email"], asunto, cuerpo)
        registrar_envio(fila, dias, exito)
        enviados += exito
        fallidos += not exito

    log.info("Resumen: %s enviados, %s fallidos, %s procesados", enviados, fallidos, len(pendientes))


# ----------------------------------------------------------------------
# Credenciales
# ----------------------------------------------------------------------

def obtener_credenciales():
    emisor = os.environ.get("EMAIL_EMISOR")
    password = os.environ.get("EMAIL_PASSWORD")

    if emisor and password:
        return emisor, password

    if not sys.stdin.isatty():
        raise RuntimeError(
            "Faltan EMAIL_EMISOR / EMAIL_PASSWORD como variables de entorno "
            "y no hay terminal interactiva para pedirlas (¿estás en CI?)."
        )

    # Fallback solo para uso manual local
    import getpass
    log.warning("EMAIL_EMISOR/EMAIL_PASSWORD no están en el entorno; pidiendo por consola.")
    emisor = input("📧 Correo electrónico emisor: ").strip()
    password = getpass.getpass("🔑 Contraseña de aplicación: ")
    return emisor, password


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Sistema de gestión de vencimientos de cartera")
    parser.add_argument(
        "--archivo",
        default=os.environ.get("CARTERA_PATH", "data/cartera.csv"),
        help="Ruta al archivo CSV/XLSX de la cartera",
    )
    parser.add_argument(
        "--dias-aviso",
        default=os.environ.get("DIAS_AVISO", ",".join(map(str, DIAS_AVISO_DEFAULT))),
        help="Umbrales de aviso separados por coma, ej: 30,15,5,1",
    )
    args = parser.parse_args()
    umbrales = [int(x) for x in args.dias_aviso.split(",") if x.strip()]

    log.info("=" * 60)
    log.info("SISTEMA DE GESTIÓN DE VENCIMIENTOS")
    log.info("Archivo: %s | Umbrales: %s", args.archivo, umbrales)
    log.info("=" * 60)

    try:
        df = leer_cartera(args.archivo)
    except Exception as e:  # noqa: BLE001
        log.error("No se pudo cargar la cartera: %s", e)
        sys.exit(1)

    if df.empty:
        log.warning("La cartera quedó vacía tras la validación; nada que procesar.")
        sys.exit(0)

    try:
        emisor, password = obtener_credenciales()
    except Exception as e:  # noqa: BLE001
        log.error("No se pudieron obtener credenciales: %s", e)
        sys.exit(1)

    procesar_vencimientos(df, emisor, password, umbrales)


if __name__ == "__main__":
    main()
