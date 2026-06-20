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

    def test_supports_aliases_with_spaces_and_case(self):
        """Debe reconocer aliases escritos entre corchetes con espacios y mayúsculas."""
        with patch("src.backend.email_sender.datetime") as mock_dt:
            mock_dt.datetime.now.return_value = datetime.datetime(2025, 6, 20)
            text = "Pago [Mes anterior en letras] de [Año en Numero]"
            result = replace_placeholders(text)
        self.assertEqual(result, "Pago Junio de 2025")

    def test_unknown_placeholders_are_left_untouched(self):
        """Los marcadores no soportados deben conservarse para no romper plantillas."""
        with patch("src.backend.email_sender.datetime") as mock_dt:
            mock_dt.datetime.now.return_value = datetime.datetime(2025, 6, 20)
            text = "Pago [Mes de pago] [variable_desconocida]"
            result = replace_placeholders(text)
        self.assertEqual(result, "Pago Junio [variable_desconocida]")


# ---------------------------------------------------------------------------
# Pruebas de send_email_via_outlook (con mock de win32com)
# ---------------------------------------------------------------------------

class TestSendEmailViaOutlook(unittest.TestCase):
    """Pruebas de la función de envío usando un mock de Outlook COM."""

    @staticmethod
    def _patch_win32_modules(mock_outlook):
        """Simula la importación dinámica de win32com.client dentro del backend."""
        mock_client = MagicMock()
        mock_client.Dispatch.return_value = mock_outlook
        mock_win32com = MagicMock()
        mock_win32com.client = mock_client
        return patch.dict(
            "sys.modules",
            {"win32com": mock_win32com, "win32com.client": mock_client},
        )

    def _make_mock_outlook(self, smtp_address: str) -> MagicMock:
        """Crea un mock de la aplicación Outlook con una cuenta configurada."""
        mock_account = MagicMock()
        mock_account.SmtpAddress = smtp_address

        mock_session = MagicMock()
        mock_session.Accounts = [mock_account]

        mock_recipients = MagicMock()
        mock_recipients.ResolveAll.return_value = True

        def add_recipient(address):
            mock_recipient = MagicMock()
            mock_recipient.address = address
            mock_recipient.Resolve.return_value = True
            return mock_recipient

        mock_recipients.Add.side_effect = add_recipient

        mock_outlook = MagicMock()
        mock_outlook.Session = mock_session
        mock_mail = MagicMock()
        mock_mail.Recipients = mock_recipients
        mock_outlook.CreateItem.return_value = mock_mail

        return mock_outlook

    @patch("src.backend.email_sender.datetime")
    @patch("src.backend.email_sender.win32com", create=True)
    def test_successful_send(self, mock_win32, mock_dt):
        """Envío exitoso cuando la cuenta existe y hay destinatarios válidos."""
        mock_dt.datetime.now.return_value = datetime.datetime(2025, 5, 15)

        mock_outlook = self._make_mock_outlook("sender@example.com")
        with self._patch_win32_modules(mock_outlook):
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

    def test_keeps_sender_in_recipients(self):
        """La cuenta emisora también puede figurar como destinatario."""
        mock_outlook = self._make_mock_outlook("sender@example.com")
        mock_mail = mock_outlook.CreateItem.return_value

        with self._patch_win32_modules(mock_outlook):
            send_email_via_outlook(
                recipients=["sender@example.com", "other@example.com"],
                subject="Test",
                body="Body",
                sender_account="sender@example.com",
            )

        added = [call.args[0] for call in mock_mail.Recipients.Add.call_args_list]
        self.assertEqual(added, ["sender@example.com", "other@example.com"])
        mock_mail.Recipients.ResolveAll.assert_called_once()

    def test_single_recipient_matching_sender_is_valid(self):
        """Una cuenta puede enviarse recordatorio a sí misma si está configurada."""
        mock_outlook = self._make_mock_outlook("sender@example.com")
        mock_mail = mock_outlook.CreateItem.return_value

        with self._patch_win32_modules(mock_outlook):
            success, msg = send_email_via_outlook(
                recipients=["sender@example.com"],  # único y es el emisor
                subject="Test",
                body="Body",
                sender_account="sender@example.com",
            )
        self.assertTrue(success)
        mock_mail.Recipients.Add.assert_called_once_with("sender@example.com")

    def test_deduplicates_recipients_case_insensitive(self):
        """Los destinatarios duplicados no deben repetirse en el correo."""
        mock_outlook = self._make_mock_outlook("sender@example.com")
        mock_mail = mock_outlook.CreateItem.return_value

        with self._patch_win32_modules(mock_outlook):
            success, msg = send_email_via_outlook(
                recipients=["recipient@example.com", " Recipient@example.com ", "sender@example.com"],
                subject="Test",
                body="Body",
                sender_account="sender@example.com",
            )

        self.assertTrue(success)
        added = [call.args[0] for call in mock_mail.Recipients.Add.call_args_list]
        self.assertEqual(added, ["recipient@example.com", "sender@example.com"])

    def test_returns_error_when_outlook_cannot_resolve_recipient(self):
        """Si Outlook no resuelve un destinatario, el envío debe fallar antes de Send()."""
        mock_outlook = self._make_mock_outlook("sender@example.com")
        mock_mail = mock_outlook.CreateItem.return_value

        def add_recipient(address):
            mock_recipient = MagicMock()
            mock_recipient.Resolve.return_value = address != "erickson558@hotmail.com"
            return mock_recipient

        mock_mail.Recipients.Add.side_effect = add_recipient
        mock_mail.Recipients.ResolveAll.return_value = False

        with self._patch_win32_modules(mock_outlook):
            success, msg = send_email_via_outlook(
                recipients=["ok@example.com", "erickson558@hotmail.com"],
                subject="Test",
                body="Body",
                sender_account="sender@example.com",
            )

        self.assertFalse(success)
        self.assertIn("erickson558@hotmail.com", msg)
        mock_mail.Send.assert_not_called()


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
