"""
test_main_window.py
-------------------
Pruebas unitarias del flujo de destinatarios en la GUI.

Verifican que agregar, eliminar, guardar y enviar usen la lista visible
en el Listbox, sin depender de lo que haya quedado en config.json.
"""
import os
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.gui.main_window import ReminderApp


class FakeListbox:
    def __init__(self, items=None, selection=None):
        self.items = list(items or [])
        self.selection = tuple(selection or ())

    def get(self, start, end=None):
        return tuple(self.items)

    def insert(self, index, value):
        self.items.append(value)

    def curselection(self):
        return self.selection

    def delete(self, index):
        del self.items[index]


class FakeEntry:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value


class FakeText:
    def __init__(self, value=""):
        self.value = value

    def get(self, start, end):
        return self.value


class FakeVar:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


class FakeLabel:
    def __init__(self, *args, **kwargs):
        self.kwargs = kwargs
        self.pack_calls = []
        self.config_calls = []

    def pack(self, *args, **kwargs):
        self.pack_calls.append((args, kwargs))

    def config(self, **kwargs):
        self.config_calls.append(kwargs)
        self.kwargs.update(kwargs)


class TestReminderAppRecipients(unittest.TestCase):
    def _make_app(self, recipients=None):
        app = ReminderApp.__new__(ReminderApp)
        app.root = MagicMock()
        app.config = {"destinatarios": ["config@example.com"]}
        app._listbox = FakeListbox(recipients or [])
        app._entry_subject = FakeEntry("Asunto")
        app._text_body = FakeText("Cuerpo")
        app._combobox_account = FakeEntry("sender@example.com")
        app._btn_send = MagicMock()
        app._auto_close_var = FakeVar(True)
        app._delay_var = FakeVar("60")
        app._lang_var = FakeVar("es")
        app._update_status = MagicMock()
        return app

    def test_get_current_recipients_uses_gui_values(self):
        app = self._make_app([" user1@example.com ", "", "user2@example.com"])

        self.assertEqual(
            app._get_current_recipients(),
            ["user1@example.com", "user2@example.com"],
        )

    @patch("src.gui.main_window.simpledialog.askstring", return_value=" new@example.com ")
    def test_add_recipient_trims_and_updates_list(self, mock_askstring):
        app = self._make_app(["user1@example.com"])

        app._add_recipient()

        self.assertEqual(app._listbox.items, ["user1@example.com", "new@example.com"])
        app._update_status.assert_called_once()

    def test_remove_recipient_deletes_selected_items(self):
        app = self._make_app(["one@example.com", "two@example.com", "three@example.com"])
        app._listbox.selection = (0, 2)

        app._remove_recipient()

        self.assertEqual(app._listbox.items, ["two@example.com"])

    def test_send_email_uses_current_gui_recipients_not_config(self):
        app = self._make_app(["gui1@example.com", "gui2@example.com"])

        with patch("src.gui.main_window.threading.Thread") as mock_thread:
            app._send_email()

        mock_thread.assert_called_once()
        _, kwargs = mock_thread.call_args
        self.assertEqual(
            kwargs["args"],
            (["gui1@example.com", "gui2@example.com"], "Asunto", "Cuerpo", "sender@example.com"),
        )

    @patch("src.gui.main_window.save_config", return_value=True)
    def test_save_config_persists_current_gui_recipients(self, mock_save_config):
        app = self._make_app([" gui1@example.com ", "gui2@example.com", ""])

        app._save_config()

        saved_config = mock_save_config.call_args.args[0]
        self.assertEqual(saved_config["destinatarios"], ["gui1@example.com", "gui2@example.com"])
        self.assertEqual(app.config["destinatarios"], ["gui1@example.com", "gui2@example.com"])

    def test_update_status_before_status_bar_exists_is_deferred(self):
        app = ReminderApp.__new__(ReminderApp)
        app._status_label = None
        app._pending_status = ("", "black")

        app._update_status("Conectando...", "blue")

        self.assertEqual(app._pending_status, ("Conectando...", "blue"))

    @patch("src.gui.main_window.tk.Label", side_effect=FakeLabel)
    def test_build_status_bar_replays_pending_status(self, mock_label):
        app = ReminderApp.__new__(ReminderApp)
        app.root = MagicMock()
        app._status_label = None
        app._pending_status = ("Conectando...", "blue")

        app._build_status_bar()

        self.assertIsNotNone(app._status_label)
        self.assertEqual(
            app._status_label.config_calls[-1],
            {"text": "Conectando...", "fg": "blue"},
        )
        mock_label.assert_called_once()


if __name__ == "__main__":
    unittest.main(verbosity=2)
