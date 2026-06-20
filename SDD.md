# Software Design Document (SDD)
## ReminderMailletYQuincena v2.0

**Proyecto:** Recordatorio automático de pago quincenal (Lety)  
**Autor:** erickson558  
**Fecha:** 2026-06-19  
**Estado:** Implementado

---

## 1. Descripción General

Aplicación de escritorio Windows que envía automáticamente un correo de recordatorio
de pago quincenal usando la cuenta de Outlook configurada en el sistema. Se ejecuta
vía Windows Task Scheduler y se cierra sola después del envío.

### 1.1 Problema Original
- El error al enviar desde Hotmail/Outlook.com se debía a:
  1. Falta de `pythoncom.CoInitialize()` en los hilos de trabajo (COM multi-thread).
  2. Mensajes de error genéricos sin código COM que dificultaban el diagnóstico.
  3. No se validaba si quedaban destinatarios después de filtrar la cuenta emisora.
  4. La GUI se congelaba durante el envío (COM en hilo principal).

### 1.2 Solución Implementada
- Envío de correo en hilo separado con COM inicializado correctamente.
- Mensajes de error específicos con código hexadecimal COM.
- Validación de destinatarios antes de llamar a `mail.Send()`.
- Logging en archivo `reminder.log` para diagnóstico post-mortem.

---

## 2. Arquitectura

### 2.1 Separación Frontend / Backend

```
reminderpagolety.py          ← Entry point (mínimo, solo arranque)
│
├── src/
│   ├── gui/
│   │   └── main_window.py  ← Frontend: Tkinter, NO contiene lógica de negocio
│   │
│   ├── backend/
│   │   ├── config_manager.py  ← I/O de config.json
│   │   └── email_sender.py    ← COM de Outlook, placeholders, envío
│   │
│   └── i18n/
│       ├── __init__.py     ← LanguageManager, t(), set_language()
│       ├── es.json         ← Traducciones en español
│       └── en.json         ← Traducciones en inglés
│
├── tests/
│   └── test_email_sender.py  ← Pruebas unitarias (mock de Outlook)
│
├── config.json              ← Configuración persistente
├── reminder.log             ← Log de ejecución (generado en runtime)
├── reminder.spec            ← Configuración de compilación PyInstaller
└── reminderagua.ico         ← Icono del ejecutable
```

### 2.2 Flujo de Datos Principal

```
[Task Scheduler] → [reminderpagolety.exe]
                          │
                   [main_window.py]
                    ReminderApp.__init__()
                          │
                   ┌──────┴──────┐
                   │             │
             load_config()    after(1000)
                   │             │
             Puebla GUI     _send_email()
                                 │
                        [Thread de envío]
                         pythoncom.CoInitialize()
                                 │
                    send_email_via_outlook()
                         win32.Dispatch("Outlook")
                                 │
                         mail.Send()
                                 │
                    pythoncom.CoUninitialize()
                                 │
                    root.after(0, _on_send_complete)
                                 │
                    [Hilo principal] → GUI update
                                 │
                    auto_close → _start_countdown() → _exit()
```

---

## 3. Componentes

### 3.1 `src/backend/config_manager.py`
| Función | Responsabilidad |
|---------|----------------|
| `get_base_path()` | Devuelve ruta base (exe vs script) |
| `get_config_path()` | Ruta al config.json |
| `load_config()` | Carga y completa con defaults |
| `save_config(config)` | Persiste dict en config.json |

### 3.2 `src/backend/email_sender.py`
| Función | Responsabilidad |
|---------|----------------|
| `replace_placeholders(text)` | Sustituye `[Mes Actual]`/`[año en numero]` |
| `get_outlook_accounts()` | Lista cuentas SMTP en Outlook |
| `send_email_via_outlook(...)` | Envía correo via COM con manejo de errores |

### 3.3 `src/i18n/__init__.py`
| Elemento | Responsabilidad |
|----------|----------------|
| `LanguageManager` | Carga JSON de traducciones |
| `t(key, **kwargs)` | Función de traducción global |
| `set_language(lang)` | Cambia idioma activo |
| `get_manager()` | Singleton del manager |

### 3.4 `src/gui/main_window.py`
| Método | Responsabilidad |
|--------|----------------|
| `_build_ui()` | Construye todos los widgets |
| `_build_top_bar()` | Selector de idioma + botón cerveza |
| `_load_accounts_thread()` | Carga cuentas Outlook async |
| `_send_email()` | Valida y lanza hilo de envío |
| `_send_thread()` | Hilo de trabajo COM |
| `_on_send_complete()` | Callback UI post-envío |
| `_refresh_ui_text()` | Actualiza labels al cambiar idioma |
| `_start_countdown()` | Cuenta regresiva de cierre |

---

## 4. Configuración (`config.json`)

```json
{
    "destinatarios": ["correo1@example.com", "correo2@example.com"],
    "asunto": "Reminder de Pagar a la Lety su quincena del Mes de [Mes Actual] de [año en numero]",
    "cuerpo": "Recordatorio de pagar quincena de Lety de [Mes Actual] de [año en numero]",
    "auto_close": true,
    "auto_close_delay": 60,
    "language": "es"
}
```

### 4.1 Placeholders disponibles
| Placeholder | Valor en runtime |
|-------------|-----------------|
| `[Mes Actual]` | Nombre del mes anterior (mes de pago) |
| `[año en numero]` | Año del mes de pago |

---

## 5. Internacionalización (i18n)

### 5.1 Agregar un nuevo idioma
1. Crear `src/i18n/<codigo>.json` copiando `es.json`.
2. Traducir todos los valores.
3. Agregar el código al diccionario `SUPPORTED_LANGUAGES` en `src/i18n/__init__.py`.
4. Recompilar el `.exe` (el JSON debe incluirse en `reminder.spec` → `datas`).

### 5.2 Idiomas actuales
| Código | Nombre | Archivo |
|--------|--------|---------|
| `es` | Español | `es.json` |
| `en` | English | `en.json` |

---

## 6. Threading y COM

El uso de Outlook COM desde Python requiere inicialización por hilo:

```python
# En cualquier hilo secundario que use win32com:
import pythoncom
pythoncom.CoInitialize()
try:
    # ... código COM ...
finally:
    pythoncom.CoUninitialize()
```

La GUI de Tkinter corre en el hilo principal (sin necesidad de CoInitialize).  
Las operaciones lentas (obtener cuentas, enviar correo) se ejecutan en hilos
separados y retornan resultados al hilo principal via `root.after(0, callback)`.

---

## 7. Logging

La aplicación escribe en `reminder.log` (junto al exe/script):
```
2026-06-19 08:10:01 [INFO] src.gui.main_window – === Aplicación iniciada ===
2026-06-19 08:10:01 [INFO] src.backend.email_sender – Cuentas Outlook encontradas: [...]
2026-06-19 08:10:02 [INFO] src.backend.email_sender – Correo enviado exitosamente vía Outlook COM
2026-06-19 08:11:02 [INFO] src.gui.main_window – === Aplicación cerrada ===
```

Útil para diagnosticar errores en ejecuciones automáticas del Task Scheduler.

---

## 8. Compilación a .exe

### 8.1 Requisitos
```bash
pip install pyinstaller pywin32
```

### 8.2 Compilar
```bash
pyinstaller reminder.spec
```

El ejecutable queda en `dist/reminderpagolety.exe`.

### 8.3 `reminder.spec` — puntos clave
- `scripts`: `['reminderpagolety.py']`
- `datas`: incluye los JSON de i18n y el .ico
- `hiddenimports`: `['win32com.client', 'win32com', 'pythoncom']`
- `icon`: `'reminderagua.ico'`
- `console=False`: sin ventana de consola

---

## 9. Skills de Claude Code

| Skill | Comando | Descripción |
|-------|---------|-------------|
| GitHub Push | `/github-push` | Commit + push a erickson558 |
| Documentar código | `/document-code` | Agrega docstrings y comentarios |
| Mejorar Python | `/improve-python` | Refactorización como Senior Engineer |

Archivos de skills en `.claude/commands/`.

---

## 10. Task Scheduler (Windows)

Archivo de configuración: `matarreminder.xml`

| Campo | Valor |
|-------|-------|
| Días | Lunes, Miércoles, Viernes |
| Hora | 08:10 AM |
| Acción | Ejecuta `reminderpagolety.exe` |
| Comportamiento | Se cierra automáticamente tras el envío |

---

## 11. Errores COM conocidos y sus soluciones

| Código COM | Descripción | Solución |
|-----------|-------------|---------|
| `0x80040115` | Outlook sin red | Verificar conexión a Internet |
| `0x80040154` | COM no registrado | Reinstalar Microsoft Office |
| `0x8004010F` | Archivo de datos inaccesible | Reiniciar Outlook |
| `0x800CCC0F` | Conexión interrumpida | Verificar que Outlook esté online |
| Auth error | OAuth2/Modern Auth requerido | Configurar cuenta en Outlook Desktop |

---

## 12. Historial de cambios

| Versión | Fecha | Cambios |
|---------|-------|---------|
| 1.0 | — | Versión inicial monolítica |
| 2.0 | 2026-06-19 | Separación frontend/backend, multi-idioma, fix COM threading, logging, botón donación, tests |
| 2.1 | 2026-06-20 | Se corrige el envío para conservar destinatarios que coinciden con la cuenta emisora, se normalizan duplicados y se actualizan las pruebas |
| 2.2 | 2026-06-20 | Se centraliza la lectura de destinatarios desde la GUI y se agregan pruebas para agregar, eliminar, guardar y enviar usando la lista visible |
| 2.3 | 2026-06-20 | Se cambia la asignación de destinatarios en Outlook a Recipients.Add con resolución explícita para evitar que se pierdan direcciones al usar mail.To |
