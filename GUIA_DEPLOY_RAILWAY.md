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

> GitHub es como Google Drive pero para código. Railway necesita que el código esté ahí para poder descargarlo y correrlo.

---

### 📌 PASO 1 — Abrir PowerShell en la carpeta del proyecto

1. Abre el **Explorador de Archivos** de Windows
2. Navega hasta: `C:\Users\Stiven I.A\Desktop\AI_LANGCHAIN STPM`
3. Haz clic en la barra de direcciones (donde dice la ruta)
4. Escribe `powershell` y presiona Enter

   ✅ Se abrirá una ventana negra (PowerShell) ya dentro de tu carpeta

---

### 📌 PASO 2 — Inicializar Git (solo se hace una vez)

Escribe este comando y presiona Enter:

```powershell
git init
```

Deberías ver algo como:
```
Initialized empty Git repository in C:/Users/Stiven I.A/Desktop/AI_LANGCHAIN STPM/.git/
```

---

### 📌 PASO 3 — Decirle a Git quién eres (solo la primera vez)

```powershell
git config user.email "tu_correo@gmail.com"
git config user.name "Tu Nombre"
```

> Reemplaza con tu correo y nombre reales. Esto se muestra en los commits de GitHub.

---

### 📌 PASO 4 — Ver qué archivos se van a subir

```powershell
git status
```

Verás una lista de archivos. Los que aparecen en **rojo** son los que aún NO están guardados.  
Los archivos en `.gitignore` (como `.env` y `service_account.json`) **no aparecerán** — eso es correcto, son secretos.

---

### 📌 PASO 5 — Agregar los archivos al commit

```powershell
git add .
```

> El punto `.` significa "agrega TODO lo que no esté en el .gitignore"

Verifica de nuevo:
```powershell
git status
```
Ahora los archivos deben aparecer en **verde** (listos para commit).

---

### 📌 PASO 6 — Hacer el primer commit

```powershell
git commit -m "Agente de facturas v1.0"
```

Verás algo como:
```
[main (root-commit) a1b2c3d] Agente de facturas v1.0
 7 files changed, 500 insertions(+)
```

---

### 📌 PASO 7 — Crear el repositorio en GitHub

1. Abre el navegador y ve a [github.com](https://github.com)
2. Inicia sesión con tu cuenta
3. Haz clic en el botón verde **"New"** (o el símbolo **+** arriba a la derecha → "New repository")

   ![Botón New en GitHub]

4. Rellena el formulario:
   - **Repository name:** `invoice-agent` (o el nombre que quieras)
   - **Description:** `Agente IA para procesar facturas desde Gmail`
   - Selecciona **Private** (para que nadie más vea tus credenciales aunque el .env no se suba)
   - ❌ **NO** marques "Add a README file" (ya tienes el tuyo)
   - ❌ **NO** marques "Add .gitignore" (ya tienes el tuyo)

5. Haz clic en **"Create repository"**

---

### 📌 PASO 8 — Conectar tu carpeta local con GitHub

GitHub te mostrará instrucciones. Copia el bloque que dice **"…or push an existing repository from the command line"**. Será algo como:

```powershell
git remote add origin https://github.com/TU_USUARIO/invoice-agent.git
git branch -M main
git push -u origin main
```

> Reemplaza `TU_USUARIO` con tu usuario real de GitHub.

Ejecuta esos 3 comandos **uno por uno** en PowerShell.

Cuando ejecutes `git push`, puede pedirte que inicies sesión en GitHub desde el navegador — simplemente sigue las instrucciones que aparecen.

---

### 📌 PASO 9 — Verificar que subió bien

1. Ve a `https://github.com/TU_USUARIO/invoice-agent`
2. Debes ver los archivos del proyecto:
   - ✅ `app.py`
   - ✅ `backend.py`
   - ✅ `start.py`
   - ✅ `requirements.txt`
   - ✅ `dashboard.html`
   - ✅ `railway.toml`
   - ❌ `.env` — NO debe aparecer (es secreto)
   - ❌ `service_account.json` — NO debe aparecer (es secreto)

---

## PARTE 2 — Crear el proyecto en Railway

---

### 📌 PASO 10 — Crear cuenta en Railway

1. Ve a [railway.app](https://railway.app)
2. Haz clic en **"Start a New Project"** o **"Login"**
3. Selecciona **"Login with GitHub"**
4. Autoriza a Railway para acceder a tu GitHub

   ✅ Quedarás dentro del panel de Railway

---

### 📌 PASO 11 — Crear nuevo proyecto

1. Dentro de Railway, haz clic en **"New Project"**

   ![Botón New Project]

2. Selecciona **"Deploy from GitHub repo"**

   ![Deploy from GitHub]

3. Si es la primera vez, Railway pedirá permiso para ver tus repositorios. Haz clic en **"Configure GitHub App"** y dale acceso al repo `invoice-agent`

4. Selecciona tu repositorio **`invoice-agent`** de la lista

5. Haz clic en **"Deploy Now"**

> Railway empezará a construir el proyecto. Verás logs de construcción aparecer. Esto tarda ~2-3 minutos. **Aún no funcionará** porque faltan las variables de entorno (credenciales). Eso lo configuramos a continuación.

---

## PARTE 3 — Configurar las variables de entorno (credenciales secretas)

> Como el `.env` no se sube a GitHub, tenemos que decirle estas claves a Railway directamente.

---

### 📌 PASO 12 — Abrir la configuración de variables

1. En Railway, haz clic en tu proyecto (el cuadro que aparece)
2. Haz clic en el servicio (un cuadro que dice tu nombre de repo)
3. Arriba verás pestañas: **Deployments**, **Variables**, **Settings**, etc.
4. Haz clic en **"Variables"**

   Verás una pantalla con una tabla vacía o con un botón **"New Variable"**

---

### 📌 PASO 13 — Agregar cada variable (una por una)

Para cada variable: clic en **"New Variable"**, escribe el nombre, escribe el valor, clic en **"Add"**.

---

**Variable 1:**
```
Nombre: EMAIL_USER
Valor:  stiven.arts@gmail.com
```

---

**Variable 2:**
```
Nombre: EMAIL_PASS
Valor:  mgye mnqz sjqg wela
```
> (Con espacios, exactamente como aparece)

---

**Variable 3:**
```
Nombre: IMAP_HOST
Valor:  imap.gmail.com
```

---

**Variable 4:**
```
Nombre: IMAP_PORT
Valor:  993
```

---

**Variable 5:**
```
Nombre: IMAP_FOLDER
Valor:  INBOX
```

---

**Variable 6:**
```
Nombre: ANTHROPIC_API_KEY
Valor:  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
```

---

**Variable 7:**
```
Nombre: GOOGLE_SHEETS_ID
Valor:  1AZcol2DMpn-BafC-zOMRs45LRpImUPu7F_iqc3f9cno
```

---

**Variable 8 — La más importante y complicada:**
```
Nombre: GOOGLE_CREDENTIALS_JSON
Valor:  (ver instrucciones abajo)
```

### ¿Cómo obtener el valor de GOOGLE_CREDENTIALS_JSON?

Esta variable debe contener **todo el contenido** de tu archivo `service_account.json`, pero en **una sola línea** (sin saltos de línea).

**Opción A — Usando PowerShell (recomendado):**

```powershell
# En PowerShell, dentro de la carpeta del proyecto:
Get-Content "service_account.json" -Raw
```

Copia TODO lo que salga (desde el `{` inicial hasta el `}` final) y pégalo como valor de la variable en Railway.

**Opción B — Manual:**
1. Abre `service_account.json` con el Bloc de Notas
2. Selecciona todo (Ctrl+A)
3. Copia (Ctrl+C)
4. Pégalo en el campo de valor de la variable en Railway

> ⚠️ **Railway acepta el JSON con múltiples líneas** en el campo de variables — no necesitas convertirlo a una sola línea. Simplemente pega el contenido completo del archivo.

---

### 📌 PASO 14 — Verificar que están todas las variables

Debes tener exactamente estas 8 variables en Railway:

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

### 📌 PASO 15 — Esperar que Railway haga el redeploy

Al agregar las variables, Railway automáticamente reinicia el build. Verás en la pestaña **"Deployments"** un nuevo build corriendo con estado `Building...`

Espera ~2-3 minutos hasta que cambie a `Active` (color verde).

---

### 📌 PASO 16 — Obtener la URL pública de tu agente

1. En Railway, haz clic en tu servicio
2. Ve a la pestaña **"Settings"**
3. Busca la sección **"Networking"** → **"Public Networking"**
4. Haz clic en **"Generate Domain"**

Railway te dará una URL como:
```
https://invoice-agent-production.up.railway.app
```

---

### 📌 PASO 17 — Probar que funciona

Abre esa URL pública en el navegador:

```
https://TU-URL.up.railway.app/api/health
```

Debe responder:
```json
{"status": "ok", "timestamp": "2026-03-02T..."}
```

Y para ver el dashboard:
```
https://TU-URL.up.railway.app/dashboard.html
```

---

### 📌 PASO 18 — Probar que procesa correos

1. Abre el dashboard: `https://TU-URL.up.railway.app/dashboard.html`
2. Haz clic en el botón **"Procesar Ahora"**
3. Mira el panel de logs — deben aparecer mensajes como:
   ```
   ✅ Conectado a imap.gmail.com
   📬 20 correos para analizar
   📨 Analizando: Factura #001...
   ✅ Factura guardada en Sheets: FAC-001 — Proveedor X
   ✅ Proceso completado — 2 factura(s) extraídas
   ```

---

## PARTE 5 — Ver los logs si algo falla

---

### 📌 PASO 19 — Ver logs de Railway (si hay error)

Si el build falla o el servidor no arranca:

1. En Railway, haz clic en tu servicio
2. Ve a la pestaña **"Deployments"**
3. Haz clic en el deployment que falló (el que tiene ❌ rojo)
4. Verás los logs completos del error

**Errores comunes y su solución:**

| Error que ves en logs | Causa | Solución |
|----------------------|-------|---------|
| `ModuleNotFoundError: No module named 'anthropic'` | Falta el paquete | Verifica `requirements.txt` tiene `anthropic>=0.25.0` |
| `GOOGLE_SHEETS_ID not set` | Variable mal escrita | Revisa el nombre exacto de la variable en Railway |
| `No such file: service_account.json` | No configuraste `GOOGLE_CREDENTIALS_JSON` | Agrega la variable con el contenido del JSON |
| `Authentication failed` (IMAP) | Contraseña incorrecta | Verifica `EMAIL_PASS` en Railway |
| `Port 9000 in use` | No pasa en Railway | El port lo asigna Railway automáticamente con `PORT` |

---

## PARTE 6 — Mantener el agente actualizado

---

### 📌 PASO 20 — Cuando hagas cambios en el código

Cada vez que modifiques un archivo y quieras que Railway lo tome:

```powershell
# En PowerShell dentro de la carpeta del proyecto:
git add .
git commit -m "Descripción del cambio"
git push
```

Railway detectará el push automáticamente y hará un nuevo deploy. ✅

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
                                                         Servidor corriendo 24/7
                                                         https://tu-url.up.railway.app
```

---

## ⏱️ Tiempo total estimado

| Parte | Tiempo |
|-------|--------|
| Subir a GitHub (Pasos 1-9) | ~10 minutos |
| Crear proyecto en Railway (Pasos 10-11) | ~5 minutos |
| Configurar variables (Pasos 12-14) | ~10 minutos |
| Esperar build y verificar (Pasos 15-18) | ~5 minutos |
| **TOTAL** | **~30 minutos** |

---

*Guía creada: 2 de marzo de 2026*
