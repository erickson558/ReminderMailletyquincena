"""
test_email_sender.py
--------------------
Pruebas unitarias para el módulo email_sender.
No requieren Outlook instalado: se mockea win32com.client.

Ejecutar con:
    python -m pytest tests/ -v
"""
import sys
import os
import unittest
from unittest.mock import MagicMock, patch
import datetime

# Asegurar acceso al paquete src desde la raíz del proyecto
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.backend.email_sender import replace_placeholders, send_email_via_outlook


# ---------------------------------------------------------------------------
# Pruebas de replace_placeholders
# ---------------------------------------------------------------------------

class TestReplacePlaceholders(unittest.TestCase):
    """Verifica la sustitución de marcadores [Mes Actual] y [año en numero]."""

    def test_normal_month(self):
        """En cualquier mes distinto a enero, el mes de pago es el mes actual."""
        # Fijamos la fecha al 15 de marzo de 2025
        with patch("src.backend.email_sender.datetime") as mock_dt:
            mock_dt.datetime.now.return_value = datetime.datetime(2025, 3, 15)
            result = replace_placeholders("Pago de [Mes Actual] [año en numero]")
        # Marzo → mes de pago es Marzo (mismo mes)
        self.assertEqual(result, "Pago de Marzo 2025")

    def test_january_rolls_to_december(self):
        """En enero, el mes de pago debe ser diciembre del año anterior."""
        with patch("src.backend.email_sender.datetime") as mock_dt:
            mock_dt.datetime.now.return_value = datetime.datetime(2025, 1, 1)
            result = replace_placeholders("[Mes Actual] [año en numero]")
        self.assertEqual(result, "Diciembre 2024")

    def test_no_placeholders(self):
        """Si no hay marcadores, el texto debe quedar igual."""
        original = "Texto sin marcadores"
        result = replace_placeholders(original)
        self.assertEqual(result, original)

    def test_multiple_occurrences(self):
        """Ambas instancias de cada marcador deben ser sustituidas."""
        with patch("src.backend.email_sender.datetime") as mock_dt:
            mock_dt.datetime.now.return_value = datetime.datetime(2025, 6, 1)
            text = "[Mes Actual] [Mes Actual] [año en numero]"
            result = replace_placeholders(text)
        self.assertEqual(result, "Junio Junio 2025")


# ---------------------------------------------------------------------------
# Pruebas de send_email_via_outlook (con mock de win32com)
# ---------------------------------------------------------------------------

class TestSendEmailViaOutlook(unittest.TestCase):
    """Pruebas de la función de envío usando un mock de Outlook COM."""

    def _make_mock_outlook(self, smtp_address: str) -> MagicMock:
        """Crea un mock de la aplicación Outlook con una cuenta configurada."""
        mock_account = MagicMock()
        mock_account.SmtpAddress = smtp_address

        mock_session = MagicMock()
        mock_session.Accounts = [mock_account]

        mock_outlook = MagicMock()
        mock_outlook.Session = mock_session
        mock_outlook.CreateItem.return_value = MagicMock()

        return mock_outlook

    @patch("src.backend.email_sender.datetime")
    @patch("src.backend.email_sender.win32com", create=True)
    def test_successful_send(self, mock_win32, mock_dt):
        """Envío exitoso cuando la cuenta existe y hay destinatarios válidos."""
        mock_dt.datetime.now.return_value = datetime.datetime(2025, 5, 15)

        mock_outlook = self._make_mock_outlook("sender@example.com")
        with patch.dict("sys.modules", {"win32com": MagicMock(), "win32com.client": MagicMock()}):
            with patch("src.backend.email_sender.win32") as mock_win32_client:
                mock_win32_client.Dispatch.return_value = mock_outlook
                success, msg = send_email_via_outlook(
                    recipients=["recipient@example.com"],
                    subject="Test [Mes Actual]",
                    body="Body [año en numero]",
                    sender_account="sender@example.com",
                )
        self.assertTrue(success)

    def test_no_pywin32(self):
        """Si pywin32 no está instalado, debe retornar un error claro."""
        with patch.dict("sys.modules", {"win32com": None, "win32com.client": None}):
            # Forzar ImportError al importar win32com.client dentro de la función
            with patch("builtins.__import__", side_effect=ImportError("No module")):
                success, msg = send_email_via_outlook(
                    recipients=["a@b.com"],
                    subject="Test",
                    body="Body",
                    sender_account="sender@example.com",
                )
        self.assertFalse(success)
        self.assertIn("pywin32", msg.lower())

    def test_filters_sender_from_recipients(self):
        """La cuenta emisora debe eliminarse de la lista de destinatarios."""
        mock_outlook = self._make_mock_outlook("sender@example.com")
        mock_mail = mock_outlook.CreateItem.return_value

        with patch.dict("sys.modules", {"win32com": MagicMock(), "win32com.client": MagicMock()}):
            with patch("src.backend.email_sender.win32") as mock_win32_client:
                mock_win32_client.Dispatch.return_value = mock_outlook
                send_email_via_outlook(
                    recipients=["sender@example.com", "other@example.com"],
                    subject="Test",
                    body="Body",
                    sender_account="sender@example.com",
                )

        # mail.To solo debe contener el destinatario que no es la cuenta emisora
        assigned_to = mock_mail.To
        if assigned_to:  # Si el mock capturó la asignación
            self.assertNotIn("sender@example.com", str(assigned_to))

    def test_empty_recipients_after_filter(self):
        """Si todos los destinatarios son la cuenta emisora, debe retornar error."""
        mock_outlook = self._make_mock_outlook("sender@example.com")

        with patch.dict("sys.modules", {"win32com": MagicMock(), "win32com.client": MagicMock()}):
            with patch("src.backend.email_sender.win32") as mock_win32_client:
                mock_win32_client.Dispatch.return_value = mock_outlook
                success, msg = send_email_via_outlook(
                    recipients=["sender@example.com"],  # único y es el emisor
                    subject="Test",
                    body="Body",
                    sender_account="sender@example.com",
                )
        self.assertFalse(success)


# ---------------------------------------------------------------------------
# Pruebas de config_manager
# ---------------------------------------------------------------------------

class TestConfigManager(unittest.TestCase):
    """Pruebas de carga y guardado de configuración."""

    def test_load_defaults_when_no_file(self):
        """Si no existe config.json, deben cargarse los valores por defecto."""
        with patch("os.path.exists", return_value=False):
            from src.backend.config_manager import load_config
            config = load_config()
        self.assertIn("destinatarios", config)
        self.assertIn("auto_close", config)
        self.assertIn("language", config)
        self.assertEqual(config["language"], "es")

    def test_save_and_load_roundtrip(self):
        """Guardar y luego cargar debe producir el mismo diccionario."""
        import json
        import tempfile
        test_config = {
            "destinatarios": ["a@b.com"],
            "asunto": "Test",
            "cuerpo": "Body",
            "auto_close": False,
            "auto_close_delay": 30,
            "language": "en",
        }
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8") as f:
            tmp_path = f.name

        try:
            with patch("src.backend.config_manager.get_config_path", return_value=tmp_path):
                from src.backend.config_manager import save_config, load_config
                save_config(test_config)
                loaded = load_config()
            self.assertEqual(loaded["destinatarios"], ["a@b.com"])
            self.assertEqual(loaded["language"], "en")
            self.assertFalse(loaded["auto_close"])
        finally:
            os.unlink(tmp_path)


if __name__ == "__main__":
    unittest.main(verbosity=2)
