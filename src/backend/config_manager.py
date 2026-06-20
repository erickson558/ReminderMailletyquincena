"""
config_manager.py
-----------------
Maneja la carga y guardado de la configuración de la aplicación en config.json.
Separado de la GUI para mantener la separación de responsabilidades.
"""
import json
import os
import sys
import logging

logger = logging.getLogger(__name__)

# Valores por defecto cuando no existe config.json
DEFAULT_CONFIG: dict = {
    "destinatarios": [],
    "asunto": "",
    "cuerpo": "",
    "auto_send_on_start": True,
    "auto_close": True,
    "auto_close_delay": 60,
    "language": "es",
}


def get_base_path() -> str:
    """
    Devuelve la ruta base de la aplicación.
    - Si está compilado como .exe (frozen), usa la carpeta del ejecutable.
    - Si corre como script Python, usa la carpeta raíz del proyecto (dos niveles
      arriba de este archivo que está en src/backend/).
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def get_config_path() -> str:
    """Devuelve la ruta completa al archivo config.json."""
    return os.path.join(get_base_path(), "config.json")


def load_config() -> dict:
    """
    Carga la configuración desde config.json.
    Si el archivo no existe o tiene errores, usa los valores por defecto.
    Siempre garantiza que todas las claves necesarias estén presentes.
    """
    config_path = get_config_path()
    config: dict = {}

    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            logger.info(f"Configuración cargada desde: {config_path}")
        except (json.JSONDecodeError, IOError) as e:
            # Archivo corrupto o inaccesible → se usan los defaults
            logger.error(f"Error al cargar config.json: {e}. Usando defaults.")
            config = {}
    else:
        logger.info("config.json no encontrado. Usando configuración por defecto.")

    # Completar con valores por defecto cualquier clave faltante
    for key, default_value in DEFAULT_CONFIG.items():
        config.setdefault(key, default_value)

    return config


def save_config(config: dict) -> bool:
    """
    Guarda el diccionario de configuración en config.json.
    Retorna True si tuvo éxito, False si hubo un error.
    """
    config_path = get_config_path()
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        logger.info(f"Configuración guardada en: {config_path}")
        return True
    except IOError as e:
        logger.error(f"Error al guardar config.json: {e}")
        return False
