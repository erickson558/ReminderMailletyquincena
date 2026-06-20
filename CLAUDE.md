# ReminderMailletYQuincena — Guía para Claude Code

Aplicación de escritorio Windows que envía un recordatorio de pago quincenal
automáticamente usando Outlook COM (pywin32). Se ejecuta con Windows Task Scheduler.

## Estructura del proyecto

```
reminderpagolety.py          ← Entrada (solo ajusta path y llama run_app)
src/
  backend/
    config_manager.py        ← Carga/guarda config.json
    email_sender.py          ← Lógica de envío Outlook COM + fix Hotmail
  gui/
    main_window.py           ← Tkinter GUI, threading, multi-idioma
  i18n/
    __init__.py              ← LanguageManager, t(), set_language()
    es.json                  ← Traducciones español
    en.json                  ← Traducciones inglés
tests/
  test_email_sender.py       ← pytest con mock de Outlook
.claude/commands/
  github-push.md             ← /github-push skill
  document-code.md           ← /document-code skill
  improve-python.md          ← /improve-python skill
SDD.md                       ← Software Design Document completo
config.json                  ← Configuración persistente (NO hardcodear aquí)
reminder.spec                ← PyInstaller spec
reminderagua.ico             ← Icono del exe
```

## Cómo ejecutar en desarrollo

```bash
pip install pywin32
python reminderpagolety.py
```

## Cómo compilar a .exe

```bash
pip install pyinstaller pywin32
pyinstaller reminder.spec
# Ejecutable en: dist/reminderpagolety.exe
```

## Dependencias

- `pywin32` — Outlook COM automation (única dependencia externa)
- Python stdlib: `tkinter`, `json`, `os`, `sys`, `datetime`, `threading`, `logging`, `webbrowser`

## Reglas para modificar este proyecto

1. **Separación estricta** — la GUI (`src/gui/`) NO importa de `win32com` directamente;
   toda lógica de email va en `src/backend/email_sender.py`.

2. **COM en hilos secundarios** — siempre usar `pythoncom.CoInitialize()` /
   `CoUninitialize()` en cualquier `threading.Thread` que llame a Outlook COM.

3. **i18n obligatorio** — todos los textos visibles en la GUI van en `es.json` y
   `en.json`; nunca strings literales en `main_window.py`.

4. **Logging** — usar `logger.info/error/debug` en lugar de `print()`.

5. **Versioning** — actualizar `SDD.md` sección "Historial de cambios" en cada
   cambio significativo.

## Skills disponibles

| Comando | Archivo | Descripción |
|---------|---------|-------------|
| `/github-push` | `.claude/commands/github-push.md` | Commit + push (erickson558) |
| `/document-code` | `.claude/commands/document-code.md` | Documenta el código |
| `/improve-python` | `.claude/commands/improve-python.md` | Refactoriza como Senior Engineer |

## Fix del error Hotmail/Outlook.com

El error al enviar desde cuentas Hotmail (`@hotmail.com`, `@outlook.com`) se debe a:
- COM no inicializado en el hilo de trabajo → usa `pythoncom.CoInitialize()`.
- Ver `src/backend/email_sender.py` → función `send_email_via_outlook()`.
- El log `reminder.log` incluye el código hexadecimal del error COM para diagnóstico.

## Botón donación
URL PayPal: `https://www.paypal.com/donate/?hosted_button_id=ZABFRXC2P3JQN`
Definido como `BEER_PAYPAL_URL` en `src/gui/main_window.py`.

## Auto-envío al arrancar
`root.after(1000, self._send_email)` en `ReminderApp.__init__()` — preserva el
comportamiento original para Task Scheduler. No eliminar.
