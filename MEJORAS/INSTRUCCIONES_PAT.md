# 🔐 Rotar PAT y migrar a variable de entorno

## Problema actual (urgente)

El token `ghp_ytst...22ZTLsE` está hardcoded en `AUTO_GITHUB.ps1:8`, visible en el commit `e92c9ea`.

GitHub Secret Scanning lo detectó y **bloquea TODOS los push desde hace tiempo** (188 commits locales acumulados sin subir).

**Riesgo real**: cualquiera que haya mirado tu repo puede usar ese token para subir código a tu repo como vos.

---

## Pasos (toma 10 minutos)

### 1️⃣ Rotar el token en GitHub

1. Abrí: https://github.com/settings/tokens
2. Buscá el token actual (empieza con `ghp_ytst...`)
3. Click en el token → **Revoke** (lo invalida inmediatamente)
4. Click en **"Generate new token (classic)"**
5. Nombre: `archivos-python-autosync-2026`
6. Expiración: 1 año (o "no expiration" si querés)
7. Scopes mínimos a marcar:
   - ✅ `repo` (todos los sub-permisos: repo:status, repo_deployment, public_repo, repo:invite, security_events)
8. Click **Generate token**
9. **COPIA EL TOKEN** (empieza con `ghp_`) — solo se muestra una vez

### 2️⃣ Setear el token como variable de entorno

Abrí PowerShell como **administrador** y corré:

```powershell
[Environment]::SetEnvironmentVariable("GITHUB_PAT", "ghp_PEGA_AQUI_EL_NUEVO_TOKEN", "User")
```

(Reemplazá `ghp_PEGA_AQUI_EL_NUEVO_TOKEN` por el token que copiaste)

Verificá que quedó bien:
```powershell
# Cerrá y abrí PowerShell de nuevo para que tome la variable
[Environment]::GetEnvironmentVariable("GITHUB_PAT", "User")
```
Debería mostrar el token.

### 3️⃣ Reemplazar AUTO_GITHUB.ps1 por la versión limpia

Ya te preparé la versión nueva en `MEJORAS\AUTO_GITHUB_v2.ps1`. Para activarla:

```powershell
cd "C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON"
Copy-Item .\AUTO_GITHUB.ps1 .\AUTO_GITHUB.ps1.backup
Copy-Item .\MEJORAS\AUTO_GITHUB_v2.ps1 .\AUTO_GITHUB.ps1
```

### 4️⃣ Resolver el commit viejo con el token

**Opción A — Rápida (recomendada ya que el token ya fue rotado y es inservible):**

Visitá el link que te dio GitHub para permitir ese secreto específico:
```
https://github.com/trujilloquezadacarlosdavid-hash/archivos-de-python/security/secret-scanning/unblock-secret/3LXp7HwCe7dYAbTBYCI2oC2bSxJ
```
Click en "**Allow secret**" → Push ya no se bloquea.

Esto es seguro **porque ya rotaste el token** → el que está en el commit viejo no sirve para nada.

**Opción B — Limpia (reescribe la historia):**

Más compleja, requiere `git filter-repo`. Solo si querés que el token desaparezca del historial. Avisame si querés hacer esto.

### 5️⃣ Probar que todo anda

```powershell
cd "C:\Users\trujillo.c\OneDrive - Me Llega, S.A\Escritorio\ARCHIVOS DE PHYTON"
.\AUTO_GITHUB.ps1
```

Si push exitoso, va a subir los **188 commits acumulados** de una vez.

---

## Checklist

- [ ] Rotado el token viejo en github.com/settings/tokens
- [ ] Generado token nuevo y copiado
- [ ] `[Environment]::SetEnvironmentVariable("GITHUB_PAT", "...", "User")` corrido
- [ ] PowerShell cerrado y reabierto
- [ ] `AUTO_GITHUB.ps1` reemplazado por la versión v2
- [ ] Permitido el secreto viejo en el link de GitHub (Opción A)
- [ ] `AUTO_GITHUB.ps1` corrido manualmente — push exitoso
- [ ] AutoGitHub programado sigue andando (próxima corrida se verifica sola)
