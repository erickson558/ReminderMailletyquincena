"""
i18n/__init__.py
----------------
Módulo de internacionalización (multi-idioma).
Expone las funciones t() y set_language() usadas en toda la GUI.

Uso:
    from src.i18n import t, set_language
    set_language("en")
    print(t("send_btn"))          # → "Send"
    print(t("status_sent"))       # → "Email sent successfully."
"""
import json
import os
import sys
import logging

logger = logging.getLogger(__name__)

# Idiomas disponibles: código ISO → nombre mostrado en la GUI
SUPPORTED_LANGUAGES: dict = {
    "es": "Español",
    "en": "English",
}
DEFAULT_LANGUAGE = "es"


def _get_i18n_dir() -> str:
    """
    Devuelve la ruta al directorio src/i18n según si la app está compilada o no.
    PyInstaller extrae los datos en una carpeta temporal (_MEIPASS) cuando está
    empaquetado en modo --onefile.
    """
    if getattr(sys, "frozen", False):
        # Ejecutable compilado: los datos se encuentran junto al .exe
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
        return os.path.join(base, "src", "i18n")
    # Modo script: la carpeta i18n está al mismo nivel que este archivo
    return os.path.dirname(os.path.abspath(__file__))


class LanguageManager:
    """
    Administrador de idioma que carga el archivo JSON de traducciones.
    Patrón Singleton: usar get_manager() en lugar de instanciar directamente.
    """

    def __init__(self, language: str = DEFAULT_LANGUAGE) -> None:
        self._language: str = DEFAULT_LANGUAGE
        self._translations: dict = {}
        self.load(language)

    def load(self, language: str) -> bool:
        """
        Carga el archivo <language>.json.
        Regresa True si tuvo éxito, False si falló (queda el idioma anterior).
        """
        if language not in SUPPORTED_LANGUAGES:
            logger.warning(f"Idioma '{language}' no soportado. Usando '{DEFAULT_LANGUAGE}'.")
            language = DEFAULT_LANGUAGE

        file_path = os.path.join(_get_i18n_dir(), f"{language}.json")
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                self._translations = json.load(f)
            self._language = language
            logger.info(f"Idioma cargado: {language}")
            return True
        except FileNotFoundError:
            logger.error(f"Archivo de idioma no encontrado: {file_path}")
            return False
        except json.JSONDecodeError as e:
            logger.error(f"Error al parsear {file_path}: {e}")
            return False

    def get(self, key: str, **kwargs) -> str:
        """
        Devuelve la traducción para 'key'.
        Si hay kwargs, los sustituye con str.format(**kwargs).
        Si la clave no existe, retorna '[key]' como placeholder visible.
        """
        text = self._translations.get(key, f"[{key}]")
        if kwargs:
            try:
                text = text.format(**kwargs)
            except (KeyError, ValueError) as e:
                logger.warning(f"Error al formatear '{key}': {e}")
        return text

    @property
    def language(self) -> str:
        """Código del idioma activo (ej. 'es', 'en')."""
        return self._language


# ---------------------------------------------------------------------------
# Singleton global del LanguageManager
# ---------------------------------------------------------------------------
_manager: LanguageManager | None = None


def get_manager() -> LanguageManager:
    """Retorna la instancia global del LanguageManager (la crea si no existe)."""
    global _manager
    if _manager is None:
        _manager = LanguageManager()
    return _manager


def set_language(language: str) -> None:
    """Cambia el idioma de la aplicación. Crea el manager si aún no existe."""
    global _manager
    if _manager is None:
        _manager = LanguageManager(language)
    else:
        _manager.load(language)


def t(key: str, **kwargs) -> str:
    """
    Función de traducción principal. Ejemplo:
        t("status_send_error", error="timeout")
    """
    return get_manager().get(key, **kwargs)
