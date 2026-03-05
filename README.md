# 🤖 Sistema de Extracción de Facturas con IA

Sistema automático que lee correos Gmail, extrae datos de facturas usando IA (Claude de Anthropic), y los guarda en Google Sheets con un dashboard en tiempo real.

## 📋 Características

- ✅ **Lectura automática de Gmail** vía IMAP
- ✅ **Extracción de datos con IA** (Claude Sonnet — Anthropic)
- ✅ **Procesamiento de PDFs, Word, XML DIAN y XLSX** adjuntos
- ✅ **API REST con FastAPI** para integración
- ✅ **Dashboard en tiempo real** con WebSocket
- ✅ **Exportación a Excel** bajo demanda
- ✅ **Logs en vivo** del procesamiento
- ✅ **Google Sheets** como almacenamiento principal (organizado por mes)

---

## 🚀 Instalación Rápida

### 1️⃣ Clonar/Descargar el Proyecto
```bash
cd "C:\Users\Stiven I.A\Desktop\AI_LANGCHAIN STPM"
```

### 2️⃣ Crear y Activar Entorno Virtual
```powershell
python -m venv AGENT_AI_ENV
.\AGENT_AI_ENV\Scripts\Activate.ps1
```

### 3️⃣ Instalar Dependencias
```bash
pip install -r requirements.txt
```

### 4️⃣ Configurar Variables de Entorno
Crea un archivo `.env` en la carpeta raíz:

```env
# Gmail
EMAIL_USER=tu@gmail.com
EMAIL_PASS=contraseña_de_aplicación_16_caract

# IMAP
IMAP_HOST=imap.gmail.com
IMAP_PORT=993
IMAP_FOLDER=INBOX

# Claude AI (Anthropic)
ANTHROPIC_API_KEY=sk-ant-api03-...

# Google Sheets
GOOGLE_SHEETS_ID=tu_id_de_google_sheets
```

### 5️⃣ Obtener Credenciales

#### Claude API Key (Anthropic):
1. Ve a: https://console.anthropic.com
2. API Keys → Crear nueva clave
3. Cópiala en `ANTHROPIC_API_KEY`

#### Gmail App Password:
1. Ve a: https://myaccount.google.com/apppasswords
2. Selecciona "Correo" y "Windows"
3. Copia la contraseña de 16 caracteres
4. Pégala en `EMAIL_PASS`

#### Google Sheets + Service Account:
1. Ve a: https://console.cloud.google.com
2. Crea proyecto → Activa Google Sheets API y Google Drive API
3. IAM → Cuentas de servicio → Crear → Descargar JSON → guardar como `service_account.json`
4. Comparte tu Google Sheet con el email de la cuenta de servicio

---

## 🎮 Ejecutar el Sistema

```powershell
# Opción 1: Script directo
.\AGENT_AI_ENV\Scripts\Activate.ps1
python start.py

# Opción 2: Doble clic en el archivo .bat
start_server.bat
```

Abre en el navegador:
- 🎨 **Dashboard:** `http://localhost:9000/dashboard.html`
- 📚 **API Docs:** `http://localhost:9000/docs`

---

## 📊 Rutas de la API

```bash
GET  /api/stats          # Estadísticas generales
GET  /api/invoices       # Lista de facturas
GET  /api/months         # Meses con datos
POST /api/process        # Procesar correos del mes actual
POST /api/process-month  # Procesar mes específico {"mes":"febrero","year":2026}
POST /api/export-excel   # Descargar Excel
GET  /api/status         # Estado del sistema
GET  /api/logs           # Logs del proceso
WS   /ws/logs            # WebSocket para logs en tiempo real
```

---

## 🗂️ Estructura de Archivos

```
AI_LANGCHAIN STPM/
├── app.py               # 🤖 Agente principal (IMAP + Claude AI + Sheets)
├── backend.py           # 🚀 API FastAPI + WebSocket + Scheduler
├── dashboard.html       # 🎨 Dashboard en tiempo real
├── start.py             # ▶️  Punto de entrada
├── requirements.txt     # 📦 Dependencias
├── railway.toml         # ☁️  Config despliegue Railway
├── .env                 # 🔐 Variables de entorno (NO subir a Git)
├── service_account.json # 🔑 Credenciales Google (NO subir a Git)
└── AGENT_AI_ENV/        # 🐍 Entorno virtual Python
```

---

## 🧠 Modelo de IA

El agente usa **Claude Sonnet** (`claude-sonnet-4-6`) de Anthropic para:
- Detectar si un correo contiene una factura electrónica DIAN o cuenta de cobro
- Extraer todos los campos contables (NIT, subtotal, IVA, retenciones, total)
- Interpretar PDFs con tablas complejas, XMLs DIAN UBL 2.1 e imágenes de facturas

---

## 🔒 Seguridad

El archivo `.gitignore` ya protege automáticamente:
```
.env                    ← API keys, credenciales de correo
service_account.json    ← Credenciales de Google
AGENT_AI_ENV/           ← Entorno virtual
processed_emails.json   ← Cache local
*.xlsx                  ← Archivos Excel generados
```

**Nunca** compartas ni subas estos archivos a GitHub.

---

## ☁️ Deploy en Railway

Ver `GUIA_DEPLOY_RAILWAY.md` para instrucciones detalladas paso a paso.

---

## 🚨 Solución de Problemas

| Error | Causa | Solución |
|-------|-------|---------|
| `ModuleNotFoundError` | Falta dependencia | `pip install -r requirements.txt` |
| `IMAP login failed` | Contraseña incorrecta | Verifica `EMAIL_PASS` en `.env` |
| `Error 429 Google Sheets` | Cuota excedida | El sistema reintenta automáticamente |
| `Dashboard en blanco` | Backend no corre | Ejecuta `python start.py` |
| `WebSocket no conecta` | Puerto bloqueado | Abre F12 y verifica consola del navegador |

---

**¡Listo! Tu sistema de extracción de facturas está configurado.** 🎉
