# /document-code — Comentar y documentar el código Python

Soy el skill para agregar docstrings, comentarios y explicaciones a todos los
módulos Python del proyecto ReminderMailletYQuincena.

## Filosofía de documentación

- **Comenta el POR QUÉ, no el QUÉ** — el código ya dice qué hace; el comentario
  explica por qué se decidió así o qué restricción existe.
- **Docstrings en todas las clases y funciones públicas** — formato Google Style.
- **Comentarios inline** solo para lógica no obvia (código COM, threading, placeholders).
- **Nunca** repitas lo que el nombre de la función ya expresa.

## Archivos que proceso

```
src/
├── backend/
│   ├── config_manager.py   ← load_config, save_config, get_base_path
│   └── email_sender.py     ← send_email_via_outlook, get_outlook_accounts,
│                              replace_placeholders, manejo de errores COM
├── gui/
│   └── main_window.py      ← ReminderApp, todos los métodos de GUI
└── i18n/
    └── __init__.py         ← LanguageManager, t(), set_language()
```

## Proceso que ejecuto

1. **Leo** cada archivo Python del proyecto.
2. **Analizo** funciones y clases sin docstring o con documentación incompleta.
3. **Agrego/mejoro** docstrings en formato Google Style:
   ```python
   def send_email_via_outlook(recipients, subject, body, sender_account):
       """
       Envía un correo usando la automatización COM de Outlook.

       Args:
           recipients: Lista de correos destino.
           subject: Asunto (puede contener placeholders).
           body: Cuerpo del correo.
           sender_account: SMTP de la cuenta emisora.

       Returns:
           Tuple (success: bool, message: str).

       Raises:
           No lanza excepciones; todos los errores se capturan y retornan como (False, msg).
       """
   ```
4. **Agrego comentarios inline** donde la lógica sea no obvia:
   - Inicialización COM (`CoInitialize`) y por qué es necesaria en hilos.
   - El filtro que elimina al emisor de la lista de destinatarios.
   - La sustitución de placeholders `[Mes Actual]` / `[año en numero]`.
5. **Reviso** que los archivos sigan compilando (`python -m py_compile <file>`).
6. **Reporto** los cambios realizados.

## Uso
```
/document-code
/document-code src/backend/email_sender.py
```
