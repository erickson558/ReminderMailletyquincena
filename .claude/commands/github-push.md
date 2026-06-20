# /github-push — Commit y Push a GitHub (cuenta erickson558)

Soy el skill para subir los cambios del proyecto ReminderMailletYQuincena a GitHub
usando la cuenta **erickson558** (ya autenticada via `gh` CLI).

## Proceso que ejecuto

1. **Verifico el estado del repositorio**
   ```bash
   git status
   git diff --stat
   ```

2. **Muestro los cambios pendientes** y pregunto un mensaje de commit si el usuario no lo proporcionó.

3. **Stages y commit** (excluyo archivos sensibles y binarios grandes):
   ```bash
   git add src/ tests/ .claude/ reminderpagolety.py reminder.spec
   git add CLAUDE.md SDD.md requirements.txt config.json
   # NO incluyo: *.exe, build/, dist/, *.log, __pycache__/
   ```

4. **Creo el commit**:
   ```bash
   git commit -m "<mensaje del usuario o auto-generado>"
   ```

5. **Push a origin main**:
   ```bash
   git push origin main
   ```

6. **Confirmo** mostrando el último commit y el URL del repo.

## Cuenta GitHub
- Usuario: **erickson558**
- Autenticación: `gh` CLI con token OAuth (keyring)
- Protocolo: HTTPS
- Scopes: `repo`, `workflow`, `read:org`, `gist`

## Archivos excluidos del commit (.gitignore)
- `dist/` — ejecutables compilados
- `build/` — artefactos de PyInstaller
- `*.exe` — el binario compilado
- `*.log` — logs de ejecución
- `__pycache__/` — bytecode Python
- `.env` — variables de entorno (si existiera)

## Uso
```
/github-push
/github-push "feat: agregar soporte multi-idioma"
```
