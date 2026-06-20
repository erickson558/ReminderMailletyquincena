"""
email_sender.py
---------------
Backend de envío de correo usando Outlook COM (win32com).
Toda la lógica de email está aquí, completamente separada de la GUI.

FIX Hotmail/Outlook.com:
  - Se llama pythoncom.CoInitialize() en hilos de trabajo para evitar errores COM.
  - Se detalla el error COM con código hexadecimal para facilitar el diagnóstico.
    - Se valida y normaliza la lista de destinatarios antes de enviar.
  - Se registran todos los eventos en el log de la aplicación.
"""
import datetime
import logging
import re
from typing import List, Tuple

logger = logging.getLogger(__name__)

# Nombres de meses en español para los placeholders
_MESES_ES: dict = {
    1: "enero",   2: "febrero", 3: "marzo",    4: "abril",
    5: "mayo",    6: "junio",   7: "julio",    8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}

_PLACEHOLDER_VALUES: dict[str, callable] = {
    "mesactual": lambda month_name, year: month_name,
    "mesanterior": lambda month_name, year: month_name,
    "mesanteriorenletras": lambda month_name, year: month_name,
    "mesdepago": lambda month_name, year: month_name,
    "añoennumero": lambda month_name, year: str(year),
    "anoennumero": lambda month_name, year: str(year),
    "añodelmesdepago": lambda month_name, year: str(year),
    "anodelmesdepago": lambda month_name, year: str(year),
}


def _normalize_placeholder_name(name: str) -> str:
    """Normaliza placeholders para soportar alias con espacios y mayúsculas."""
    normalized = name.strip().lower()
    replacements = {
        "á": "a",
        "é": "e",
        "í": "i",
        "ó": "o",
        "ú": "u",
        "ñ": "n",
    }

    for source, target in replacements.items():
        normalized = normalized.replace(source, target)

    return re.sub(r"[^a-z0-9]", "", normalized)


def _normalize_recipients(recipients: List[str]) -> List[str]:
    """Limpia destinatarios vacíos y elimina duplicados preservando el orden."""
    normalized: List[str] = []
    seen: set[str] = set()

    for recipient in recipients:
        clean_recipient = recipient.strip()
        if not clean_recipient:
            continue

        lookup_key = clean_recipient.lower()
        if lookup_key in seen:
            continue

        seen.add(lookup_key)
        normalized.append(clean_recipient)

    return normalized


def _assign_recipients(mail, recipients: List[str]) -> Tuple[bool, str]:
    """Agrega y resuelve destinatarios en Outlook para evitar pérdidas al parsear mail.To."""
    unresolved: List[str] = []

    for recipient in recipients:
        outlook_recipient = mail.Recipients.Add(recipient)
        resolved = True

        try:
            resolved = bool(outlook_recipient.Resolve())
        except Exception:
            resolved = False

        if not resolved:
            unresolved.append(recipient)

    try:
        if hasattr(mail.Recipients, "ResolveAll") and not mail.Recipients.ResolveAll():
            if not unresolved:
                unresolved = list(recipients)
    except Exception:
        if not unresolved:
            unresolved = list(recipients)

    if unresolved:
        msg = f"No se pudieron resolver estos destinatarios en Outlook: {unresolved}"
        logger.error(msg)
        return False, msg

    logger.info(f"Destinatarios resueltos en Outlook: {recipients}")
    return True, ""


# ---------------------------------------------------------------------------
# Función de sustitución de placeholders
# ---------------------------------------------------------------------------

def replace_placeholders(text: str) -> str:
    """
    Sustituye los marcadores de posición en asunto/cuerpo:
            [Mes Actual] / [Mes anterior en letras] → nombre del mes de pago
            [año en numero]                         → año correspondiente al mes de pago

    Lógica: el pago corresponde al mes que acaba de terminar.
    Si estamos en enero, el mes de pago es diciembre del año anterior.
    """
    now = datetime.datetime.now()
    month, year = now.month, now.year

    if month == 1:
        # Enero → mes de pago = diciembre del año pasado
        prev_month, prev_year = 12, year - 1
    else:
        prev_month, prev_year = month, year

    month_name = _MESES_ES[prev_month].capitalize()

    def replace_match(match: re.Match[str]) -> str:
        placeholder_name = match.group(1)
        normalized_name = _normalize_placeholder_name(placeholder_name)
        value_factory = _PLACEHOLDER_VALUES.get(normalized_name)

        if value_factory is None:
            return match.group(0)

        return value_factory(month_name, prev_year)

    return re.sub(r"\[([^\[\]]+)\]", replace_match, text)


# ---------------------------------------------------------------------------
# Obtención de cuentas Outlook configuradas en el sistema
# ---------------------------------------------------------------------------

def get_outlook_accounts() -> List[str]:
    """
    Devuelve la lista de direcciones SMTP de las cuentas configuradas en Outlook.
    Debe llamarse desde un hilo con pythoncom.CoInitialize() ya aplicado.
    Retorna lista vacía si Outlook no está disponible o hay un error.
    """
    try:
        import win32com.client as win32  # parte de pywin32
        outlook = win32.Dispatch("Outlook.Application")
        accounts = [acct.SmtpAddress for acct in outlook.Session.Accounts]
        logger.info(f"Cuentas Outlook encontradas: {accounts}")
        return accounts
    except ImportError:
        logger.error("pywin32 no está instalado. Outlook COM no disponible.")
        return []
    except Exception as e:
        # Mostrar código hexadecimal del error COM para diagnóstico
        logger.error(f"Error al obtener cuentas Outlook: {e!r}")
        return []


# ---------------------------------------------------------------------------
# Envío de correo mediante Outlook COM
# ---------------------------------------------------------------------------

def send_email_via_outlook(
    recipients: List[str],
    subject: str,
    body: str,
    sender_account: str,
) -> Tuple[bool, str]:
    """
    Envía un correo usando la automatización COM de Outlook.
    Debe llamarse desde un hilo con pythoncom.CoInitialize() ya aplicado.

    Parámetros
    ----------
    recipients      : Lista de correos destinatarios.
    subject         : Asunto del correo (puede tener placeholders).
    body            : Cuerpo del correo (puede tener placeholders).
    sender_account  : SMTP de la cuenta emisora seleccionada.

    Retorna
    -------
    (True, "Email sent successfully")  si se envió correctamente.
    (False, "<descripción del error>") si hubo un fallo.
    """
    try:
        import win32com.client as win32
    except ImportError:
        msg = "pywin32 no instalado. Instala con: pip install pywin32"
        logger.error(msg)
        return False, msg

    try:
        # Crear la aplicación Outlook mediante COM
        outlook = win32.Dispatch("Outlook.Application")

        # Crear un nuevo ítem de correo (0 = olMailItem)
        mail = outlook.CreateItem(0)

        # Aplicar sustitución de placeholders en asunto y cuerpo
        mail.Subject = replace_placeholders(subject)
        mail.Body = replace_placeholders(body)

        # Buscar y asignar la cuenta emisora por SMTP
        account_found = False
        for account in outlook.Session.Accounts:
            if account.SmtpAddress.lower() == sender_account.lower():
                mail.SendUsingAccount = account  # fuerza el envío con esta cuenta
                account_found = True
                logger.info(f"Cuenta emisora seleccionada: {account.SmtpAddress}")
                break

        if not account_found:
            msg = f"Cuenta no encontrada en Outlook: {sender_account}"
            logger.error(msg)
            return False, msg

        filtered = _normalize_recipients(recipients)

        if not filtered:
            msg = (
                "Sin destinatarios válidos. La lista está vacía o solo contiene "
                "entradas en blanco."
            )
            logger.error(msg)
            return False, msg

        recipients_ok, recipients_error = _assign_recipients(mail, filtered)
        if not recipients_ok:
            return False, recipients_error

        logger.info(f"Enviando a: {filtered}")

        # Enviar el correo a través de Outlook (usa la configuración SMTP de la cuenta)
        mail.Send()
        logger.info("Correo enviado exitosamente vía Outlook COM")
        return True, "Correo enviado exitosamente"

    except Exception as e:
        # Construir mensaje de error detallado con código COM si está disponible
        raw = str(e)
        hex_code = ""
        if hasattr(e, "args") and e.args:
            # Los errores COM tienen el código numérico en args[0]
            try:
                hex_code = f" [COM: {hex(e.args[0])}]"
            except (TypeError, ValueError):
                pass

        logger.error(f"Error al enviar correo: {raw}{hex_code}")

        # Mensajes específicos para los errores COM más comunes con Hotmail/Outlook.com
        if "0x80040115" in raw or "-2147221227" in raw:
            return False, (
                "Outlook no tiene conexión a la red. "
                "Verifica tu conexión a Internet.{hex_code}"
            ).format(hex_code=hex_code)
        elif "0x80040154" in raw or "-2147221164" in raw:
            return False, (
                "Outlook no está correctamente instalado o registrado. "
                "Reinstala Microsoft Office.{hex_code}"
            ).format(hex_code=hex_code)
        elif "0x8004010F" in raw or "-2147221233" in raw:
            return False, (
                "No se puede acceder al archivo de datos de Outlook. "
                "Reinicia Outlook e intenta de nuevo.{hex_code}"
            ).format(hex_code=hex_code)
        elif "0x800CCC0F" in raw or "disconnected" in raw.lower():
            return False, (
                "Conexión interrumpida con el servidor de correo. "
                "Verifica que Outlook esté conectado.{hex_code}"
            ).format(hex_code=hex_code)
        elif "authentication" in raw.lower() or "credentials" in raw.lower():
            return False, (
                f"Error de autenticación con Outlook.{hex_code}\n"
                "Para cuentas Hotmail/Outlook.com: verifica que Outlook "
                "esté correctamente autenticado con autenticación moderna (OAuth2).\n"
                f"Detalle: {raw}"
            )
        else:
            return False, f"Error al enviar: {raw}{hex_code}"
