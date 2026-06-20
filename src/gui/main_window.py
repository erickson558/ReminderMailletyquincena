"""
main_window.py
--------------
Capa de presentación (GUI) construida con Tkinter.
No contiene lógica de negocio: delega todo al módulo backend.

Características:
  - Soporte multi-idioma (ES / EN) con cambio dinámico sin reiniciar la app.
  - Botón "Invítame una Cerveza" con enlace a PayPal.
  - Envío de correo en hilo separado para evitar que la GUI se congele.
  - Manejo correcto de COM en hilos (pythoncom.CoInitialize / CoUninitialize).
  - Logging de todos los eventos relevantes.
  - Auto-envío 1 segundo después del inicio (comportamiento del Task Scheduler).
"""
import logging
import os
import sys
import threading
import webbrowser
import tkinter as tk
from tkinter import simpledialog, ttk

from ..backend.config_manager import load_config, save_config, get_base_path
from ..backend.email_sender import get_outlook_accounts, send_email_via_outlook
from ..i18n import t, set_language, get_manager, SUPPORTED_LANGUAGES

logger = logging.getLogger(__name__)

# URL del botón de donación
BEER_PAYPAL_URL = "https://www.paypal.com/donate/?hosted_button_id=ZABFRXC2P3JQN"


def _setup_logging() -> None:
    """
    Configura el sistema de logging.
    Escribe en 'reminder.log' junto al ejecutable/script y también en la consola.
    """
    log_file = os.path.join(get_base_path(), "reminder.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s – %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    logger.info("=== Aplicación iniciada ===")


# ---------------------------------------------------------------------------
# Clase principal de la aplicación
# ---------------------------------------------------------------------------

class ReminderApp:
    """
    Ventana principal de la aplicación.
    Encapsula toda la GUI y coordina las llamadas al backend.
    """

    def __init__(self, root: tk.Tk) -> None:
        """
        Inicializa la aplicación:
          1. Carga la configuración guardada.
          2. Aplica el idioma.
          3. Construye la interfaz.
          4. Pobla los campos con los valores de la configuración.
          5. Programa el auto-envío a 1 segundo del inicio.
        """
        self.root = root
        self.config = load_config()

        # Aplicar idioma guardado en configuración
        set_language(self.config.get("language", "es"))

        self._build_ui()
        self._populate_fields()

        # Auto-envío: preserva el comportamiento original del Task Scheduler.
        # Se dispara 1 segundo después de que la ventana es visible y estable.
        self.root.after(1000, self._send_email)

    # -----------------------------------------------------------------------
    # Construcción de la interfaz
    # -----------------------------------------------------------------------

    def _build_ui(self) -> None:
        """Construye todos los widgets de la ventana."""
        self.root.title(t("app_title"))
        self.root.resizable(False, False)  # Tamaño fijo para consistencia visual

        self._build_top_bar()
        self._build_recipients_section()
        self._build_subject_section()
        self._build_body_section()
        self._build_account_section()
        self._build_auto_close_section()
        self._build_action_buttons()
        self._build_status_bar()

    def _build_top_bar(self) -> None:
        """
        Barra superior con selector de idioma (radio buttons) y
        botón de donación 'Invítame una Cerveza'.
        """
        frame = tk.Frame(self.root)
        frame.pack(fill=tk.X, padx=10, pady=(8, 0))

        # --- Selector de idioma ---
        self._lbl_lang = tk.Label(frame, text=t("language_label"), font=("Segoe UI", 9))
        self._lbl_lang.pack(side=tk.LEFT)

        self._lang_var = tk.StringVar(value=get_manager().language)
        self._radio_buttons: list[tk.Radiobutton] = []
        for code, name in SUPPORTED_LANGUAGES.items():
            rb = tk.Radiobutton(
                frame,
                text=name,
                variable=self._lang_var,
                value=code,
                command=self._on_language_change,
                font=("Segoe UI", 9),
            )
            rb.pack(side=tk.LEFT, padx=4)
            self._radio_buttons.append(rb)

        # --- Botón de donación (derecha) ---
        self._btn_beer = tk.Button(
            frame,
            text=t("buy_beer_btn"),
            fg="white",
            bg="#003087",          # Azul PayPal
            activeforeground="white",
            activebackground="#001f5b",
            relief=tk.FLAT,
            padx=8,
            pady=2,
            cursor="hand2",
            command=lambda: webbrowser.open(BEER_PAYPAL_URL),
            font=("Segoe UI", 9, "bold"),
        )
        self._btn_beer.pack(side=tk.RIGHT, padx=5)

    def _build_recipients_section(self) -> None:
        """Sección de lista de destinatarios con botones Agregar / Eliminar."""
        frame = tk.Frame(self.root)
        frame.pack(pady=(10, 0), fill=tk.X, padx=10)

        self._lbl_recipients = tk.Label(frame, text=t("recipients"), font=("Segoe UI", 9, "bold"))
        self._lbl_recipients.pack(anchor="w")

        self._listbox = tk.Listbox(frame, width=60, height=5, font=("Segoe UI", 9))
        self._listbox.pack(pady=4)

        btn_frame = tk.Frame(frame)
        btn_frame.pack()

        self._btn_add = tk.Button(btn_frame, text=t("add_recipient"), width=15, command=self._add_recipient)
        self._btn_add.pack(side=tk.LEFT, padx=5)

        self._btn_remove = tk.Button(btn_frame, text=t("remove_recipient"), width=15, command=self._remove_recipient)
        self._btn_remove.pack(side=tk.LEFT, padx=5)

    def _build_subject_section(self) -> None:
        """Campo de texto para el asunto del correo."""
        self._lbl_subject = tk.Label(self.root, text=t("subject"), font=("Segoe UI", 9, "bold"))
        self._lbl_subject.pack(pady=(10, 0))

        self._entry_subject = tk.Entry(self.root, width=60, font=("Segoe UI", 9))
        self._entry_subject.pack(pady=4)

    def _build_body_section(self) -> None:
        """Área de texto multilínea para el cuerpo del correo."""
        self._lbl_body = tk.Label(self.root, text=t("body"), font=("Segoe UI", 9, "bold"))
        self._lbl_body.pack(pady=(10, 0))

        self._text_body = tk.Text(self.root, width=60, height=8, font=("Segoe UI", 9))
        self._text_body.pack(pady=4)

    def _build_account_section(self) -> None:
        """
        LabelFrame con el ComboBox para seleccionar la cuenta de envío de Outlook.
        Las cuentas se cargan de forma asíncrona para no bloquear el arranque.
        """
        self._frame_account = tk.LabelFrame(
            self.root, text=t("send_account"), font=("Segoe UI", 9, "bold")
        )
        self._frame_account.pack(padx=10, pady=6, fill="both")

        self._lbl_account = tk.Label(
            self._frame_account, text=t("select_account"), font=("Segoe UI", 9)
        )
        self._lbl_account.pack(anchor="w", padx=10, pady=4)

        self._combobox_account = ttk.Combobox(
            self._frame_account, values=[], state="readonly", width=50
        )
        self._combobox_account.pack(anchor="w", padx=10, pady=4)

        # Cargar cuentas Outlook en un hilo separado para no bloquear la GUI
        self._update_status(t("status_connecting"))
        thread = threading.Thread(target=self._load_accounts_thread, daemon=True)
        thread.start()

    def _build_auto_close_section(self) -> None:
        """Sección de configuración del cierre automático después de enviar."""
        self._frame_autoclose = tk.LabelFrame(
            self.root, text=t("auto_close_config"), font=("Segoe UI", 9, "bold")
        )
        self._frame_autoclose.pack(padx=10, pady=6, fill="both")

        self._auto_close_var = tk.BooleanVar(value=self.config.get("auto_close", True))
        self._chk_autoclose = tk.Checkbutton(
            self._frame_autoclose,
            text=t("auto_close_label"),
            variable=self._auto_close_var,
            font=("Segoe UI", 9),
        )
        self._chk_autoclose.pack(anchor="w", padx=10, pady=4)

        self._lbl_delay = tk.Label(
            self._frame_autoclose, text=t("auto_close_delay_label"), font=("Segoe UI", 9)
        )
        self._lbl_delay.pack(anchor="w", padx=10)

        self._delay_var = tk.StringVar(value=str(self.config.get("auto_close_delay", 60)))
        self._entry_delay = tk.Entry(
            self._frame_autoclose, textvariable=self._delay_var, width=8, font=("Segoe UI", 9)
        )
        self._entry_delay.pack(anchor="w", padx=10, pady=4)

    def _build_action_buttons(self) -> None:
        """Fila de botones: Enviar, Guardar configuración, Salir."""
        frame = tk.Frame(self.root)
        frame.pack(pady=10)

        self._btn_send = tk.Button(
            frame, text=t("send_btn"), width=20, command=self._send_email,
            font=("Segoe UI", 9, "bold"), bg="#28a745", fg="white",
            activebackground="#218838", activeforeground="white", relief=tk.FLAT,
        )
        self._btn_send.pack(side=tk.LEFT, padx=5)

        self._btn_save = tk.Button(
            frame, text=t("save_btn"), width=20, command=self._save_config,
            font=("Segoe UI", 9),
        )
        self._btn_save.pack(side=tk.LEFT, padx=5)

        self._btn_exit = tk.Button(
            frame, text=t("exit_btn"), width=20, command=self._exit,
            font=("Segoe UI", 9), fg="red",
        )
        self._btn_exit.pack(side=tk.LEFT, padx=5)

    def _build_status_bar(self) -> None:
        """Barra de estado en la parte inferior de la ventana."""
        self._status_label = tk.Label(
            self.root, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=("Segoe UI", 9)
        )
        self._status_label.pack(side=tk.BOTTOM, fill=tk.X)

    # -----------------------------------------------------------------------
    # Llenado de campos con datos de configuración
    # -----------------------------------------------------------------------

    def _get_current_recipients(self) -> list[str]:
        """Devuelve la lista actual de destinatarios mostrada en la GUI."""
        return [email.strip() for email in self._listbox.get(0, tk.END) if email.strip()]

    def _populate_fields(self) -> None:
        """Puebla los widgets de la GUI con los valores cargados de config.json."""
        for email in self.config.get("destinatarios", []):
            self._listbox.insert(tk.END, email)

        self._entry_subject.delete(0, tk.END)
        self._entry_subject.insert(0, self.config.get("asunto", ""))

        self._text_body.delete("1.0", tk.END)
        self._text_body.insert("1.0", self.config.get("cuerpo", ""))

    # -----------------------------------------------------------------------
    # Actualización de textos de la GUI (cambio de idioma)
    # -----------------------------------------------------------------------

    def _refresh_ui_text(self) -> None:
        """
        Actualiza el texto de todos los widgets sin destruir la GUI.
        Se llama cuando el usuario cambia el idioma.
        """
        self.root.title(t("app_title"))
        self._lbl_lang.config(text=t("language_label"))
        self._btn_beer.config(text=t("buy_beer_btn"))
        self._lbl_recipients.config(text=t("recipients"))
        self._btn_add.config(text=t("add_recipient"))
        self._btn_remove.config(text=t("remove_recipient"))
        self._lbl_subject.config(text=t("subject"))
        self._lbl_body.config(text=t("body"))
        self._frame_account.config(text=t("send_account"))
        self._lbl_account.config(text=t("select_account"))
        self._frame_autoclose.config(text=t("auto_close_config"))
        self._chk_autoclose.config(text=t("auto_close_label"))
        self._lbl_delay.config(text=t("auto_close_delay_label"))
        self._btn_send.config(text=t("send_btn"))
        self._btn_save.config(text=t("save_btn"))
        self._btn_exit.config(text=t("exit_btn"))

    # -----------------------------------------------------------------------
    # Carga asíncrona de cuentas Outlook
    # -----------------------------------------------------------------------

    def _load_accounts_thread(self) -> None:
        """
        Hilo de trabajo para obtener las cuentas de Outlook.
        Inicializa COM correctamente para el contexto de hilo secundario.
        """
        try:
            import pythoncom
            pythoncom.CoInitialize()  # requerido en hilos secundarios para usar COM
            try:
                cuentas = get_outlook_accounts()
            finally:
                pythoncom.CoUninitialize()
        except ImportError:
            # pythoncom no disponible (entorno sin pywin32)
            cuentas = []

        # Actualizar la GUI en el hilo principal usando after()
        self.root.after(0, self._on_accounts_loaded, cuentas)

    def _on_accounts_loaded(self, cuentas: list) -> None:
        """Callback que se ejecuta en el hilo principal una vez cargadas las cuentas."""
        self._combobox_account.config(values=cuentas)
        if cuentas:
            self._combobox_account.current(0)
            self._update_status("")          # Limpiar "Conectando..."
            logger.info(f"Cuentas cargadas en ComboBox: {cuentas}")
        else:
            self._update_status(
                t("status_no_accounts", error="Verifica que Outlook esté abierto y configurado"),
                "red",
            )

    # -----------------------------------------------------------------------
    # Envío de correo (no bloqueante)
    # -----------------------------------------------------------------------

    def _send_email(self) -> None:
        """
        Valida los campos y lanza el envío en un hilo separado para que
        la GUI no se congele mientras Outlook procesa el correo.
        """
        # Recolectar y validar datos de la GUI
        recipients = self._get_current_recipients()
        if not recipients:
            self._update_status(t("status_no_recipients"), "red")
            return

        subject = self._entry_subject.get().strip()
        body = self._text_body.get("1.0", tk.END).strip()
        sender = self._combobox_account.get()

        if not sender:
            self._update_status(t("status_account_not_found"), "red")
            return

        # Deshabilitar el botón Enviar mientras se procesa para evitar doble envío
        self._btn_send.config(state=tk.DISABLED)
        self._update_status(t("status_sending"))

        # Enviar en hilo secundario para no bloquear la interfaz
        thread = threading.Thread(
            target=self._send_thread,
            args=(recipients, subject, body, sender),
            daemon=True,
        )
        thread.start()

    def _send_thread(self, recipients: list, subject: str, body: str, sender: str) -> None:
        """
        Hilo de trabajo que realiza el envío de correo.
        Inicializa COM para el hilo y retorna el resultado al hilo principal.
        """
        try:
            import pythoncom
            pythoncom.CoInitialize()  # inicializar COM en este hilo
            try:
                success, message = send_email_via_outlook(recipients, subject, body, sender)
            finally:
                pythoncom.CoUninitialize()
        except ImportError:
            success, message = False, "pythoncom no disponible (instala pywin32)"

        # Notificar al hilo principal con el resultado
        self.root.after(0, self._on_send_complete, success, message)

    def _on_send_complete(self, success: bool, message: str) -> None:
        """Callback en el hilo principal con el resultado del envío."""
        self._btn_send.config(state=tk.NORMAL)

        if success:
            self._update_status(t("status_sent"), "green")
            logger.info("Correo enviado exitosamente")
            if self._auto_close_var.get():
                try:
                    delay = int(self._delay_var.get())
                except ValueError:
                    delay = 60  # valor de respaldo si el campo tiene texto inválido
                self._start_countdown(delay)
        else:
            self._update_status(t("status_send_error", error=message), "red")
            logger.error(f"Error al enviar: {message}")

    # -----------------------------------------------------------------------
    # Gestión de destinatarios
    # -----------------------------------------------------------------------

    def _add_recipient(self) -> None:
        """Abre un diálogo para agregar un destinatario a la lista."""
        email = simpledialog.askstring(
            t("dialog_add_recipient"),
            t("dialog_add_recipient_prompt"),
            parent=self.root,
        )
        if email and email.strip():
            self._listbox.insert(tk.END, email.strip())
            self._update_status(t("status_recipient_added"), "green")

    def _remove_recipient(self) -> None:
        """Elimina el/los destinatario(s) seleccionado(s) de la lista."""
        selection = self._listbox.curselection()
        if not selection:
            self._update_status(t("status_select_to_remove"), "red")
            return
        # Eliminar en orden inverso para que los índices no se desplacen
        for index in reversed(selection):
            self._listbox.delete(index)
        self._update_status(t("status_recipient_removed"), "green")

    # -----------------------------------------------------------------------
    # Guardar configuración
    # -----------------------------------------------------------------------

    def _save_config(self) -> None:
        """Lee los valores actuales de la GUI y los persiste en config.json."""
        try:
            delay = int(self._delay_var.get())
        except ValueError:
            delay = 60

        new_config = {
            "destinatarios": self._get_current_recipients(),
            "asunto": self._entry_subject.get(),
            "cuerpo": self._text_body.get("1.0", tk.END).strip(),
            "auto_close": self._auto_close_var.get(),
            "auto_close_delay": delay,
            "language": self._lang_var.get(),
        }
        if save_config(new_config):
            self.config = new_config
            self._update_status(t("status_config_saved"), "green")
        else:
            self._update_status(t("status_config_error", error="Error de escritura"), "red")

    # -----------------------------------------------------------------------
    # Cambio de idioma
    # -----------------------------------------------------------------------

    def _on_language_change(self) -> None:
        """
        Cambia el idioma activo y refresca todos los textos de la GUI.
        Guarda el idioma nuevo en config.json para que persista.
        """
        lang = self._lang_var.get()
        set_language(lang)
        self.config["language"] = lang
        save_config(self.config)
        self._refresh_ui_text()
        logger.info(f"Idioma cambiado a: {lang}")

    # -----------------------------------------------------------------------
    # Cuenta regresiva de cierre automático
    # -----------------------------------------------------------------------

    def _start_countdown(self, seconds: int) -> None:
        """Inicia la cuenta regresiva que cierra la app al llegar a cero."""
        def _tick() -> None:
            nonlocal seconds
            if seconds > 0:
                self._update_status(t("status_closing", seconds=seconds), "green")
                seconds -= 1
                self.root.after(1000, _tick)
            else:
                self._exit()

        _tick()

    # -----------------------------------------------------------------------
    # Utilidades
    # -----------------------------------------------------------------------

    def _update_status(self, message: str, color: str = "black") -> None:
        """Actualiza el texto y color de la barra de estado inferior."""
        self._status_label.config(text=message, fg=color)
        if message:
            logger.debug(f"Status → {message}")

    def _exit(self) -> None:
        """Cierra la ventana y termina la aplicación."""
        logger.info("=== Aplicación cerrada ===")
        self.root.destroy()


# ---------------------------------------------------------------------------
# Función de entrada pública
# ---------------------------------------------------------------------------

def run_app() -> None:
    """
    Crea la ventana raíz de Tkinter, inicializa ReminderApp y arranca el loop.
    Esta función es el único punto de entrada a la GUI.
    """
    _setup_logging()
    root = tk.Tk()
    _app = ReminderApp(root)
    root.mainloop()
