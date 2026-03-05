# 🚂 Guía de Despliegue en Railway
## Paso a paso, muy detallado — Para principiantes

---

> **¿Qué es Railway?**  
> Es una página web donde puedes poner a correr tu código Python **24 horas al día, los 7 días de la semana**, en servidores en la nube. Tiene un plan gratuito de 500 horas/mes.  
> Cuando lo despliegues, el agente procesará facturas aunque apagues tu computador.

---

## 🔴 ANTES DE EMPEZAR — verifica que tienes esto

- [ ] Una cuenta en **GitHub** (si no tienes, créala gratis en [github.com](https://github.com))
- [ ] **Git** instalado en tu PC (verifica con: `git --version` en PowerShell)
- [ ] El proyecto funcionando localmente (el servidor arranca con `python start.py`)

---

## PARTE 1 — Subir el código a GitHub

---

### 📌 PASO 1 — Abrir PowerShell en la carpeta del proyecto

1. Abre el **Explorador de Archivos** de Windows
2. Navega hasta: `C:\Users\Stiven I.A\Desktop\AI_LANGCHAIN STPM`
3. Haz clic en la barra de direcciones → escribe `powershell` → Enter

---

### 📌 PASO 2 — Inicializar Git (solo se hace una vez)

```powershell
git init
```

---

### 📌 PASO 3 — Decirle a Git quién eres (solo la primera vez)

```powershell
git config user.email "tu_correo@gmail.com"
git config user.name "Tu Nombre"
```

---

### 📌 PASO 4 — Ver qué archivos se van a subir

```powershell
git status
```

Los archivos en `.gitignore` (`.env`, `service_account.json`) **no aparecerán** — correcto, son secretos.

---

### 📌 PASO 5 — Agregar los archivos al commit

```powershell
git add .
git status
```

---

### 📌 PASO 6 — Hacer el primer commit

```powershell
git commit -m "Agente de facturas v1.0"
```

---

### 📌 PASO 7 — Crear el repositorio en GitHub

1. Ve a [github.com](https://github.com) → **New repository**
2. Nombre: `invoice-agent`
3. Selecciona **Private**
4. ❌ NO marques "Add a README" ni "Add .gitignore"
5. Clic en **Create repository**

---

### 📌 PASO 8 — Conectar tu carpeta local con GitHub

```powershell
git remote add origin https://github.com/TU_USUARIO/invoice-agent.git
git branch -M main
git push -u origin main
```

---

### 📌 PASO 9 — Verificar que subió bien

Ve a `https://github.com/TU_USUARIO/invoice-agent` y verifica:

- ✅ `app.py`, `backend.py`, `start.py`, `requirements.txt`, `dashboard.html`, `railway.toml`
- ❌ `.env` — NO debe aparecer
- ❌ `service_account.json` — NO debe aparecer

---

## PARTE 2 — Crear el proyecto en Railway

---

### 📌 PASO 10 — Crear cuenta en Railway

1. Ve a [railway.app](https://railway.app) → **Login with GitHub**
2. Autoriza a Railway para acceder a tu GitHub

---

### 📌 PASO 11 — Crear nuevo proyecto

1. **New Project** → **Deploy from GitHub repo**
2. Selecciona `invoice-agent`
3. Clic en **Deploy Now**

> Railway empezará a construir. Aún no funcionará porque faltan las variables de entorno.

---

## PARTE 3 — Configurar las variables de entorno

---

### 📌 PASO 12 — Abrir la configuración de variables

Proyecto → Servicio → pestaña **Variables** → **New Variable**

---

### 📌 PASO 13 — Agregar cada variable

> ⚠️ **IMPORTANTE**: Nunca copies claves reales en documentos ni en el repositorio.
> Copia cada valor directamente desde tu archivo `.env` local.

| Variable | Dónde obtenerla |
|----------|----------------|
| `EMAIL_USER` | Tu correo Gmail |
| `EMAIL_PASS` | Contraseña de aplicación de Gmail (16 caracteres) |
| `IMAP_HOST` | `imap.gmail.com` |
| `IMAP_PORT` | `993` |
| `IMAP_FOLDER` | `INBOX` |
| `ANTHROPIC_API_KEY` | [console.anthropic.com](https://console.anthropic.com) → API Keys |
| `GOOGLE_SHEETS_ID` | URL de tu Sheet: la parte entre `/d/` y `/edit` |
| `GOOGLE_CREDENTIALS_JSON` | Contenido completo de `service_account.json` (ver abajo) |

---

### ¿Cómo obtener el valor de GOOGLE_CREDENTIALS_JSON?

```powershell
# En PowerShell, dentro de la carpeta del proyecto:
Get-Content "service_account.json" -Raw
```

Copia todo el output (desde `{` hasta `}`) y pégalo como valor de la variable en Railway.

---

### 📌 PASO 14 — Verificar que están todas las variables

| Variable | ¿Tiene valor? |
|----------|--------------|
| `EMAIL_USER` | ✅ |
| `EMAIL_PASS` | ✅ |
| `IMAP_HOST` | ✅ |
| `IMAP_PORT` | ✅ |
| `IMAP_FOLDER` | ✅ |
| `ANTHROPIC_API_KEY` | ✅ |
| `GOOGLE_SHEETS_ID` | ✅ |
| `GOOGLE_CREDENTIALS_JSON` | ✅ |

---

## PARTE 4 — Verificar el despliegue

---

### 📌 PASO 15 — Esperar el redeploy automático

Al agregar variables, Railway reinicia el build. Espera ~2-3 minutos hasta que el estado sea `Active` (verde).

---

### 📌 PASO 16 — Obtener la URL pública

Servicio → **Settings** → **Networking** → **Generate Domain**

URL resultante: `https://invoice-agent-production.up.railway.app`

---

### 📌 PASO 17 — Probar que funciona

```
https://TU-URL.up.railway.app/api/health
→ {"status": "ok", "timestamp": "..."}

https://TU-URL.up.railway.app/dashboard.html
→ Dashboard visual
```

---

### 📌 PASO 18 — Probar que procesa correos

1. Abre el dashboard
2. Selecciona el mes y año
3. Clic en **▶ Procesar**
4. Observa los logs en tiempo real

---

## PARTE 5 — Ver los logs si algo falla

---

### 📌 PASO 19 — Logs de Railway

Proyecto → **Deployments** → clic en el deployment fallido → ver logs completos

**Errores comunes:**

| Error en logs | Causa | Solución |
|--------------|-------|---------|
| `ModuleNotFoundError` | Falta paquete | Verifica `requirements.txt` |
| `GOOGLE_SHEETS_ID not set` | Variable mal escrita | Revisa nombre exacto en Railway |
| `No such file: service_account.json` | Falta `GOOGLE_CREDENTIALS_JSON` | Agrega la variable con el JSON |
| `Authentication failed` (IMAP) | Contraseña incorrecta | Verifica `EMAIL_PASS` |

---

## PARTE 6 — Mantener el agente actualizado

---

### 📌 PASO 20 — Subir cambios de código

```powershell
git add .
git commit -m "Descripción del cambio"
git push
```

Railway detecta el push y hace redeploy automático. ✅

---

## 📊 Resumen visual del proceso

```
Tu PC                          GitHub                    Railway
──────                         ──────                    ───────
Código local
    │
    │  git add .
    │  git commit
    │  git push
    └──────────────────────► Repositorio ──────────────► Build automático
                             invoice-agent                    │
                                                              │ python start.py
                                                              │
                                                         Servidor 24/7
                                                         https://tu-url.up.railway.app
```

---

## ⏱️ Tiempo total estimado

| Parte | Tiempo |
|-------|--------|
| Subir a GitHub | ~10 min |
| Crear proyecto en Railway | ~5 min |
| Configurar variables | ~10 min |
| Verificar despliegue | ~5 min |
| **TOTAL** | **~30 min** |

---

*Guía actualizada: marzo de 2026*
