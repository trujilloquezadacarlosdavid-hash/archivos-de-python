---
name: git-safe
description: Reglas duras para operaciones git en el repo archivos-de-python de Carlos. Usar SIEMPRE que haya que hacer commit, push, reset, rebase, checkout, clean o cualquier operación git. Cubre: cómo evitar que Secret Scanning bloquee el push (PAT, API keys), convivencia con AutoGitHub cada 2h, qué NO commitear (dashboards que se regeneran, logs, .xlsm pesados), nunca force-push a main, y qué hacer cuando GitHub rechaza el push.
---

# Git Safe - operar sin romper nada

Carlos tiene auto-sync cada 2h con `AutoGitHub_Python`. Esta skill evita colisiones y pérdida de trabajo.

## Repo

- **Nombre**: `trujilloquezadacarlosdavid-hash/archivos-de-python`
- **Rama principal**: `main`
- **Auto-sync**: `AutoGitHub_Python` cada 2h (hace `add .`, `commit`, `push`)
- **PC redundante**: Dorcas Castro (`recepciones@`) con el mismo repo clonado
- **Working dir principal**: `C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON`

## Regla #1 — NUNCA commitear secretos

GitHub tiene **Push Protection** activo y rechaza TODO push que contenga:
- Personal Access Tokens (`ghp_...`, `github_pat_...`)
- API keys (OpenAI, Anthropic, etc.)
- Tokens Slack, Twilio, Supabase
- Connection strings con password

**Incidente 2026-11-02**: Push bloqueado por PAT en `AUTO_GITHUB.ps1:8` commit `e92c9ea`. GitHub lo detectó y bloqueó todos los push posteriores.

### Antes de cada commit

```powershell
# Chequear archivos staged por patrones de secreto
git diff --cached | Select-String -Pattern "ghp_|github_pat_|sk-|xoxb-|AKIA|password\s*=\s*[`"']"
```

Si aparece algo → NO commitear ese archivo. Mover el secreto a variable de entorno:
```powershell
# ❌ En el script
$token = "ghp_abc123..."

# ✅ En el script
$token = $env:GITHUB_PAT
# Y setear $env:GITHUB_PAT en el perfil de Windows (no en el repo)
```

### Si el push ya fue rechazado

```
remote: - Push cannot contain secrets
remote: —— GitHub Personal Access Token ——
remote:    locations:
remote:      - commit: <sha>
remote:        path: ARCHIVO.ps1:LINEA
```

Pasos:
1. **Rotar el token comprometido** en github.com/settings/tokens (está público aunque sea en commit viejo)
2. **Mover el nuevo token a variable de entorno** o archivo `.env` (que esté en `.gitignore`)
3. **Reescribir el commit** que lo contiene (destructivo, requiere OK de Carlos):
   - `git filter-repo --path ARCHIVO.ps1 --invert-paths` (y volver a crear el archivo limpio)
   - O permitir el secreto en el link que GitHub proveyó (menos seguro, útil si el token ya se rotó)
4. **NUNCA usar `git push --force` sin autorización explícita** — borra trabajo de AutoGitHub

## Regla #2 — Convivir con AutoGitHub (cada 2h)

AutoGitHub hace `git add .` + `git commit` + `git push` automáticamente. Esto implica:

### Qué puede pisar

- Si Carlos tiene cambios locales no commiteados cuando AutoGitHub corre → los agarra y los commitea tal cual (incluso código a medio hacer)
- Dashboards que se auto-actualizan generan 1 commit por corrida → la historia se llena de "Auto-sync YYYY-MM-DD HH:MM"

### Patrón correcto para trabajar

```powershell
# ANTES de empezar cambios grandes
git pull --rebase          # agarrar lo que AutoGitHub pusheó

# Hacer tus cambios

# ANTES de commit manual
git status                  # ver qué tocó AutoGitHub mientras tanto
git add <archivos-específicos>  # NO usar git add -A (toma dashboards actualizándose)
git commit -m "..."
git pull --rebase          # por si corrió AutoGitHub mientras commiteabas
git push
```

### Qué poner en .gitignore

```
# Dashboards que se auto-regeneran (ruido en historia)
*/dashboard_*.html
*/refresco_log.txt
*/_*.log
*/_fsd_*.json       # excepto si son plantillas base

# Archivos pesados / binarios
*.xlsm
*.pdf
*.png             # pero sí commitear si son del repo (plantillas)

# Secrets
.env
*token*.txt
auth.json
```

## Regla #3 — Nunca force-push a main

- ❌ `git push --force origin main` — borra commits que AutoGitHub o PC Dorcas subieron
- ❌ `git push -f` — igual de peligroso
- ❌ `git reset --hard origin/main` SIN `git status` previo — borra uncommitted work

Si realmente hay que reescribir historia:
1. Decirle a Carlos QUÉ se va a reescribir y por qué
2. Esperar OK explícito
3. Confirmar que AutoGitHub no corra mientras se hace la reescritura
4. Hacer backup: `git branch backup-pre-reescritura`
5. Ejecutar
6. Verificar con `git log` que lo que quedó es lo esperado
7. Push force-with-lease (más seguro que --force): `git push --force-with-lease`

## Regla #4 — Antes de operaciones destructivas, status primero

Comandos que pueden perder trabajo:
- `git checkout -- <file>` (descarta cambios locales)
- `git restore <file>` (igual)
- `git reset --hard` (borra todo lo no commiteado)
- `git clean -fd` (borra untracked)
- `git rebase -i` (puede perder commits si se cancela mal)
- `rm -rf` en el working dir

ANTES de cualquiera de esos:
```powershell
git status
# Si hay cambios sin commit → stash primero
git stash push -u -m "backup antes de reset"
# Hacer la operación destructiva
# Si todo bien → git stash drop
# Si algo mal → git stash pop
```

## Regla #5 — Commits legibles

Formato recomendado:
```
<tipo>: <qué cambió en una línea>

<por qué, si no es obvio>

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
```

Tipos: `feat`, `fix`, `refactor`, `chore`, `docs`, `data` (para datos/dashboards actualizados).

AutoGitHub usa `Auto-sync YYYY-MM-DD HH:MM` — no competir con ese formato en commits manuales.

## Regla #6 — Confirmar antes de push

```powershell
# Mostrar qué se va a pushear
git log origin/main..HEAD --oneline
# Si es más de 3 commits o toca archivos sensibles → confirmar con Carlos
git push origin main
```

## Checklist antes de git push

- [ ] `git diff --cached` NO muestra secrets (PAT, API keys, passwords)
- [ ] `git status` NO muestra archivos sospechosos de secrets (`.env`, `*token*`, `*cred*`)
- [ ] `git log origin/main..HEAD --oneline` muestra los commits esperados (no basura de AutoGitHub mezclada)
- [ ] NO estoy usando `--force` sin OK explícito de Carlos
- [ ] NO estoy pisando commits que AutoGitHub subió
- [ ] Si falló por secret → rotar el token ANTES de hacer workaround
