# 📄 Documentación del Agente de Facturas con IA
### Guía paso a paso — explicada para principiantes

---

## ¿Qué hace este agente?

Este agente se conecta a tu Gmail, lee los correos entrantes, detecta si alguno contiene una factura (en el cuerpo o como adjunto PDF/DOCX/XLSX), extrae automáticamente los datos con Inteligencia Artificial (Claude de Anthropic) y los guarda en Google Sheets. Todo esto con un tablero visual desde el navegador.

```
Gmail ──► Agente ──► Claude AI ──► Google Sheets ──► Dashboard Web
```

---

## 🗂️ Estructura de archivos del proyecto

```
AI_LANGCHAIN STPM/
│
├── app.py                  ← El cerebro del agente (leer correos, llamar IA, guardar)
├── backend.py              ← El servidor web (API, dashboard, WebSocket)
├── start.py                ← Punto de entrada para arrancar todo
├── dashboard.html          ← Interface visual en el navegador
├── requirements.txt        ← Lista de librerías que necesita Python
├── .env                    ← Claves secretas (email, API keys, etc.)
├── service_account.json    ← Credenciales de Google Sheets
├── railway.toml            ← Configuración para desplegar en Railway (nube)
├── .gitignore              ← Archivos que NO se suben a GitHub
└── processed_emails.json   ← Cache de correos ya procesados (se limpia automático)
```

---

## 🧱 PASO 1 — Preparar el entorno de Python

### ¿Qué es un entorno virtual?
Es una carpeta aislada donde instalas las librerías de tu proyecto sin mezclarlas con otras cosas de tu PC.

### Cómo se creó

```powershell
# Crear el entorno virtual
python -m venv AGENT_AI_ENV

# Activarlo (en Windows)
.\AGENT_AI_ENV\Scripts\Activate.ps1

# Instalar las dependencias
pip install -r requirements.txt
```

### Contenido de `requirements.txt` (lo que se instaló)

```
fastapi==0.109.0          ← Servidor web rápido
uvicorn==0.27.0           ← El que "sirve" el servidor
python-dotenv==1.2.1      ← Lee las variables del archivo .env
anthropic>=0.25.0         ← SDK para hablar con la IA de Anthropic (Claude)
gspread>=6.0.0            ← SDK para leer/escribir en Google Sheets
google-auth>=2.0.0        ← Autenticación con Google
pandas>=2.1.0             ← Manejo de tablas de datos
openpyxl>=3.1.0           ← Generar archivos Excel
PyPDF2>=3.0.0             ← Extraer texto de PDFs
python-docx>=1.0.0        ← Extraer texto de Word (.docx)
beautifulsoup4>=4.12.0    ← Limpiar HTML de correos
schedule==1.2.2           ← Ejecutar tareas automáticas cada X minutos
```

---

## 🔑 PASO 2 — Configurar las credenciales (archivo `.env`)

El archivo `.env` guarda todas las claves secretas. **Nunca lo subas a GitHub.**

```env
# Correo Gmail que el agente va a leer
EMAIL_USER=tu_correo@gmail.com
EMAIL_PASS=xxxx xxxx xxxx xxxx      ← Contraseña de aplicación (no tu contraseña normal)

# Servidor de correo (para Gmail siempre es esto)
IMAP_HOST=imap.gmail.com
IMAP_PORT=993
IMAP_FOLDER=INBOX

# Clave de Claude AI (Anthropic)
ANTHROPIC_API_KEY=sk-ant-api03-...

# ID de tu Google Sheet (la parte de la URL entre /d/ y /edit)
GOOGLE_SHEETS_ID=1AZcol2DMpn-BafC-zOMRs45LRpImUPu7F_iqc3f9cno
```

### ¿Cómo obtener la contraseña de aplicación de Gmail?

1. Ve a [myaccount.google.com](https://myaccount.google.com)
2. Seguridad → Verificación en dos pasos → Contraseñas de aplicaciones
3. Crea una para "Correo" / "Windows"
4. Copia las 16 letras que aparecen (con espacios: `mgye mnqz sjqg wela`)

---

## 📊 PASO 3 — Configurar Google Sheets

### ¿Para qué?
En vez de guardar en un archivo Excel local, los datos van directo a una hoja de Google que puedes ver desde cualquier lugar.

### Cómo se hizo

**3.1 — Crear una cuenta de servicio en Google Cloud**

1. Ir a [console.cloud.google.com](https://console.cloud.google.com)
2. Crear proyecto → Habilitar las APIs: **Google Sheets API** y **Google Drive API**
3. IAM y Administración → Cuentas de servicio → Crear cuenta de servicio
4. Descargar el JSON de credenciales → guardarlo como `service_account.json` en la carpeta del proyecto

El archivo `service_account.json` tiene este aspecto:
```json
{
  "type": "service_account",
  "project_id": "gen-lang-client-XXXXXXXX",
  "private_key_id": "...",
  "private_key": "-----BEGIN RSA PRIVATE KEY-----\n...",
  "client_email": "agent-354@gen-lang-client-XXXXXXXX.iam.gserviceaccount.com",
  "client_id": "...",
  ...
}
```

**3.2 — Compartir la hoja con la cuenta de servicio**

1. Abre tu Google Sheet
2. Clic en "Compartir" (botón verde arriba a la derecha)
3. Agrega el email de la cuenta de servicio: `agent-354@gen-lang-client-XXXXXXXX.iam.gserviceaccount.com`
4. Dale permiso de **Editor**

---

## 🤖 PASO 4 — Obtener la API Key de Claude (Anthropic)

1. Ir a [console.anthropic.com](https://console.anthropic.com)
2. Crear una cuenta y luego ir a API Keys
3. Crear una nueva clave → copiarla al `.env` como `ANTHROPIC_API_KEY`

El modelo que usa el agente es: `claude-haiku-4-5` (rápido y económico)

---

## 💻 PASO 5 — Entender los dos archivos principales

### `app.py` — El agente (cerebro)

Este archivo NO tiene servidor web. Solo hace el trabajo de procesar correos.

**Flujo completo:**

```
process_emails()
    │
    ├── setup_sheets()          → Conecta a Google Sheets, crea headers si no existen
    ├── connect_imap()          → Se conecta por SSL a Gmail
    ├── search_invoice_emails() → Trae los últimos 20 correos
    │
    └── Para cada correo nuevo:
         ├── _get_email_content()  → Extrae texto del cuerpo y adjuntos (PDF/DOCX/XLSX)
         ├── _extract_with_ai()    → Llama a Claude para detectar factura y extraer datos
         └── save_to_sheets()      → Guarda en Google Sheets (evita duplicados)
```

**Función clave — Llamada a Claude AI:**

```python
def _extract_with_ai(subject, body, attachments):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    
    prompt = """Analiza el siguiente correo y extrae los datos de factura.
    Si NO es una factura, responde: NO_ES_FACTURA
    Si SÍ es una factura, responde SOLO en JSON:
    {
      "numero_factura": "...",
      "proveedor": "...",
      "fecha_emision": "DD/MM/YYYY",
      "fecha_vencimiento": "DD/MM/YYYY",
      "total": 0.0,
      "moneda": "COP",
      "estado": "PENDIENTE",
      "descripcion": "..."
    }
    """
    
    resp = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=512,
        messages=[{"role": "user", "content": prompt}]
    )
    # Procesa la respuesta y extrae el JSON
```

**Columnas que se guardan en Google Sheets:**

| N° Factura | Proveedor | Fecha Emisión | Fecha Vencimiento | Total | Moneda | Estado | Descripción | Email Origen | Fecha Procesado |

**Función de conexión a Google Sheets (soporta local Y nube):**

```python
def get_sheet():
    creds_json_env = os.getenv("GOOGLE_CREDENTIALS_JSON")
    
    if creds_json_env:
        # ☁️ En Railway (nube): lee credenciales desde variable de entorno
        creds_dict = json.loads(creds_json_env)
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        # 💻 En local: lee el archivo service_account.json
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    client = gspread.authorize(creds)
    return client.open_by_key(GOOGLE_SHEETS_ID).worksheet("Facturas")
```

---

### `backend.py` — El servidor web (API)

Este archivo crea el servidor FastAPI que sirve el dashboard y expone los endpoints.

**Endpoints disponibles:**

| Método | Ruta | ¿Qué hace? |
|--------|------|-----------|
| GET | `/` | Dashboard principal |
| GET | `/dashboard.html` | Dashboard HTML |
| GET | `/api/health` | Estado del servidor |
| GET | `/api/stats` | Estadísticas (total, pendientes, pagadas, vencidas) |
| GET | `/api/invoices` | Lista las últimas 50 facturas desde Sheets |
| POST | `/api/process` | Dispara el procesamiento de correos manualmente |
| POST | `/api/export-excel` | Descarga un Excel con todos los datos de Sheets |
| GET | `/api/logs` | Logs del proceso |
| DELETE | `/api/logs` | Limpia los logs |
| GET | `/api/status` | Estado del agente y configuración |
| GET/POST | `/api/scheduler` | Ver/configurar el scheduler automático |
| WS | `/ws/logs` | WebSocket para logs en tiempo real |
| GET | `/docs` | Documentación automática de la API (Swagger) |

**Cómo se capturan los logs para el dashboard:**

```python
class LogCapture(io.TextIOBase):
    def write(self, msg):
        # Clasifica el mensaje como error, success, warning o info
        # Lo agrega al log_store que transmite por WebSocket
        log_store.add_log(msg, level)
```

Todo `print()` que hace `app.py` durante el proceso aparece en tiempo real en el dashboard.

---

## 🚀 PASO 6 — Arrancar el servidor localmente

```powershell
# Activar entorno
.\AGENT_AI_ENV\Scripts\Activate.ps1

# Arrancar
python start.py
```

Verás en la consola:
```
  Dashboard : http://localhost:9000/dashboard.html
  API Docs  : http://localhost:9000/docs
```

Abre el navegador en `http://localhost:9000/dashboard.html`

### Verificar que funciona

```powershell
Invoke-RestMethod "http://localhost:9000/api/health"
# Debe responder: {"status": "ok", "timestamp": "2026-03-02T..."}
```

---

## ⚙️ PASO 7 — Scheduler (procesamiento automático)

El agente puede procesar correos automáticamente sin que hagas clic. Desde el dashboard o desde la API:

```powershell
# Activar scheduler cada 30 minutos
Invoke-RestMethod -Method POST -Uri "http://localhost:9000/api/scheduler" `
  -ContentType "application/json" `
  -Body '{"enabled": true, "mode": "interval", "interval_minutes": 30}'

# O activar todos los días a las 8:00 AM
Invoke-RestMethod -Method POST -Uri "http://localhost:9000/api/scheduler" `
  -ContentType "application/json" `
  -Body '{"enabled": true, "mode": "daily", "daily_time": "08:00"}'
```

---

## ☁️ PASO 8 — Desplegar en Railway (servidor 24/7 en la nube)

> Railway es una plataforma que corre tu código en la nube gratis (500 horas/mes en plan Hobby).

### 8.1 — Preparar el repositorio en GitHub

```powershell
git init
git add .
git commit -m "Agente de facturas v1.0"
# Crear repo en github.com y luego:
git remote add origin https://github.com/TU_USUARIO/invoice-agent.git
git push -u origin main
```

> **Importante:** `.gitignore` ya está configurado para NO subir `.env`, `service_account.json` ni el entorno virtual.

### 8.2 — Crear proyecto en Railway

1. Ir a [railway.app](https://railway.app) → Login con GitHub
2. New Project → Deploy from GitHub repo → Seleccionar tu repo
3. Railway detectará automáticamente `railway.toml` y usará `python start.py`

### 8.3 — Configurar variables de entorno en Railway

En Railway, ir a tu proyecto → Variables y agregar **una por una**:

| Variable | Valor |
|----------|-------|
| `EMAIL_USER` | `tu_correo@gmail.com` |
| `EMAIL_PASS` | `mgye mnqz sjqg wela` |
| `IMAP_HOST` | `imap.gmail.com` |
| `IMAP_PORT` | `993` |
| `IMAP_FOLDER` | `INBOX` |
| `ANTHROPIC_API_KEY` | `sk-ant-api03-...` |
| `GOOGLE_SHEETS_ID` | `1AZcol2DMpn-BafC-zOMRs45LRpImUPu7F_iqc3f9cno` |
| `GOOGLE_CREDENTIALS_JSON` | *(ver abajo)* |

**Para `GOOGLE_CREDENTIALS_JSON`:**
Abre `service_account.json`, copia TODO su contenido (todo el JSON en una sola línea) y pégalo como valor de esa variable.

```powershell
# En PowerShell, convierte el archivo a una sola línea para copiar fácil:
(Get-Content service_account.json -Raw) | ConvertFrom-Json | ConvertTo-Json -Compress
```

### 8.4 — El archivo `railway.toml` (ya creado)

```toml
[build]
builder = "NIXPACKS"

[deploy]
startCommand = "python start.py"
healthcheckPath = "/api/health"
healthcheckTimeout = 30
restartPolicyType = "ON_FAILURE"
restartPolicyMaxRetries = 3
```

Después de configurar las variables, Railway hace el deploy automático y en ~2 minutos tendrás una URL como:
`https://invoice-agent-production.up.railway.app`

---

## 🔄 Flujo completo del agente (resumen visual)

```
┌─────────────────────────────────────────────────────────────────┐
│                         start.py                                │
│                   python start.py                               │
│            Arranca FastAPI en puerto 9000 (o PORT de Railway)   │
└──────────────────────────┬──────────────────────────────────────┘
                           │ importa
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│                        backend.py                               │
│  - Servidor FastAPI con todos los endpoints                     │
│  - WebSocket para logs en tiempo real                           │
│  - Scheduler para ejecutar proceso automáticamente              │
│  - Lee estadísticas y facturas directamente desde Google Sheets │
└──────────────────────────┬──────────────────────────────────────┘
                           │ importa como módulo
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│                          app.py                                 │
│                                                                 │
│  process_emails()                                               │
│       │                                                         │
│       ├── Conecta a Gmail por IMAP/SSL                          │
│       ├── Toma los últimos 20 correos                           │
│       ├── Descarta los ya procesados (processed_emails.json)    │
│       │                                                         │
│       └── Por cada correo nuevo:                                │
│            ├── Extrae texto del cuerpo (HTML → texto limpio)    │
│            ├── Extrae texto de adjuntos (PDF, DOCX, XLSX)       │
│            ├── Envía todo a Claude AI para detectar factura     │
│            │        └── Claude responde JSON con datos          │
│            └── Guarda en Google Sheets (evita duplicados)       │
│                                                                 │
│  export_to_excel()  → Genera Excel desde los datos de Sheets    │
└─────────────────────────────────────────────────────────────────┘
                           │ escribe / lee
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│                    Google Sheets                                │
│   Hoja: "Facturas"                                              │
│   Columnas: N°Factura | Proveedor | Fechas | Total | Estado...  │
└─────────────────────────────────────────────────────────────────┘
```

---

## ❌ Problemas que se resolvieron (y cómo)

### Problema 1: `app.py` se importaba a sí mismo (circular import)
**Síntoma:** El servidor no arrancaba, error de importación.  
**Causa:** `backend.py` importaba `app.py`, que a su vez tenía una instancia de FastAPI y se llamaba a sí mismo.  
**Solución:** `app.py` es solo un módulo de procesamiento (sin FastAPI). `backend.py` es el único servidor.

### Problema 2: "0 correos nuevos" aunque había correos
**Síntoma:** Cada vez que se ejecutaba, decía que no había correos nuevos.  
**Causa:** `processed_emails.json` tenía guardados los IDs de todos los correos de ejecuciones anteriores.  
**Solución:** `backend.py` vacía el archivo antes de cada ejecución manual, luego lo restaura.

### Problema 3: Claude respondía JSON envuelto en ```json```
**Síntoma:** Error al parsear la respuesta de la IA.  
**Causa:** Claude a veces envuelve el JSON en bloques de código markdown.  
**Solución:** Parser de 3 niveles: intento directo → busca bloque ```json``` → busca llaves balanceadas.

### Problema 4: La IA de Google (Gemini) daba error 429
**Síntoma:** "RESOURCE_EXHAUSTED" al llamar a la API.  
**Causa:** El plan gratuito de Gemini 2.5 Flash tiene límite de 5 llamadas por minuto.  
**Solución:** Se migró a Claude (Anthropic) que tiene límites más generosos.

### Problema 5: Excel bloqueado cuando se intentaba guardar
**Síntoma:** `PermissionError: [Errno 13]` al escribir el archivo Excel.  
**Causa:** El archivo estaba abierto en Excel mientras Python intentaba sobreescribirlo.  
**Solución:** Se migró a Google Sheets como almacenamiento principal. Excel solo se genera bajo demanda desde el botón "Exportar Excel" del dashboard.

---

## 📋 Comandos útiles del día a día

```powershell
# Activar entorno e iniciar
.\AGENT_AI_ENV\Scripts\Activate.ps1
python start.py

# Probar que el servidor está vivo
Invoke-RestMethod "http://localhost:9000/api/health"

# Ver estadísticas de facturas
Invoke-RestMethod "http://localhost:9000/api/stats"

# Disparar procesamiento manual (sin abrir dashboard)
Invoke-RestMethod -Method POST "http://localhost:9000/api/process"

# Ver los últimos logs
Invoke-RestMethod "http://localhost:9000/api/logs"

# Matar el servidor si quedó colgado
Get-Process python -ErrorAction SilentlyContinue | Stop-Process -Force

# Ver qué está usando el puerto 9000
netstat -ano | findstr :9000
```

---

## 🔒 Seguridad — qué NO subir a GitHub

El archivo `.gitignore` ya protege estos archivos automáticamente:

```
.env                    ← Claves de email, API keys
service_account.json    ← Credenciales de Google
AGENT_AI_ENV/           ← Entorno virtual (pesa mucho y no es necesario)
__pycache__/            ← Archivos compilados de Python
processed_emails.json   ← Cache local
*.xlsx                  ← Archivos Excel generados
```

---

## 📦 Versiones y tecnologías usadas

| Tecnología | Versión | Para qué |
|-----------|---------|---------|
| Python | 3.13 | Lenguaje principal |
| FastAPI | 0.109.0 | Servidor web / API REST |
| Uvicorn | 0.27.0 | Servidor ASGI (corre FastAPI) |
| Anthropic (Claude) | ≥0.25.0 | IA para extraer datos de facturas |
| gspread | ≥6.0.0 | Leer/escribir en Google Sheets |
| google-auth | ≥2.0.0 | Autenticación con Google |
| pandas | ≥2.1.0 | Manejo de tablas de datos |
| openpyxl | ≥3.1.0 | Generar archivos Excel |
| PyPDF2 | ≥3.0.0 | Leer PDFs adjuntos |
| python-docx | ≥1.0.0 | Leer Word adjuntos |
| beautifulsoup4 | ≥4.12.0 | Limpiar HTML de correos |
| schedule | 1.2.2 | Automatización de tareas |
| python-dotenv | 1.2.1 | Leer archivo `.env` |

---

*Documentación generada: 2 de marzo de 2026*
