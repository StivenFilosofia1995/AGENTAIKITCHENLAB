# 🤖 Sistema de Extracción de Facturas con IA

Sistema automático que lee correos Gmail, extrae datos de facturas usando IA (Gemini), y los guarda en Excel con un dashboard en tiempo real.

## 📋 Características

- ✅ **Lectura automática de Gmail** vía IMAP
- ✅ **Extracción de datos con IA** (Google Gemini Pro)
- ✅ **Procesamiento de PDFs y Word** adjuntos
- ✅ **API REST con FastAPI** para integración
- ✅ **Dashboard en tiempo real** con WebSocket
- ✅ **Excel profesional** con formateo y gráficos
- ✅ **Logs en vivo** del procesamiento
- ✅ **Límite configurable** de emails procesados

---

## 🚀 Instalación Rápida

### 1️⃣ Clonar/Descargar el Proyecto
```bash
cd "C:\Users\Stiven I.A\Desktop\AI_LANGCHAIN STPM"
```

### 2️⃣ Crear y Activar Entorno Virtual
```powershell
# Si aún no lo has hecho
python -m venv AGENT_AI_ENV

# Activar
.\AGENT_AI_ENV\Scripts\Activate.ps1
```

### 3️⃣ Instalar Dependencias
```bash
pip install -r requirements.txt
```

### 4️⃣ Configurar Variables de Entorno
Crea un archivo `.env` en la carpeta raíz:

```env
# Google API
GOOGLE_API_KEY=tu_google_api_key_aqui

# Gmail - CREDENCIALES
EMAIL_USER=tu@gmail.com
EMAIL_PASS=contraseña_de_aplicación_16_caract

# IMAP
IMAP_HOST=imap.gmail.com
IMAP_PORT=993
IMAP_FOLDER=INBOX

# Excel
EXCEL_OUTPUT_PATH=facturas_seguimiento.xlsx

# Scheduler
CHECK_INTERVAL_MINUTES=30

# SerpAPI (Opcional)
SERPAPI_API_KEY=tu_serpapi_key_opcional
```

### 5️⃣ Obtener Credenciales

#### Google Gemini API Key:
1. Ve a: https://ai.google.dev/
2. Click "Get API Key"
3. Copia tu clave

#### Gmail App Password:
1. Ve a: https://myaccount.google.com/apppasswords
2. Selecciona "Correo" y "Windows"
3. Copia la contraseña de 16 caracteres
4. Pégala en `EMAIL_PASS`

---

## 🎮 Ejecutar el Sistema

### Opción 1: Solo Agente (Sin Dashboard)
```powershell
# Ejecución única
.\AGENT_AI_ENV\Scripts\python.exe app.py --once

# Ejecución continua (cada 30 min)
.\AGENT_AI_ENV\Scripts\python.exe app.py
```

### Opción 2: FastAPI + Dashboard 🎯 (RECOMENDADO)

#### Terminal 1 - Backend FastAPI:
```powershell
# Instalar FastAPI si no lo has hecho
pip install fastapi uvicorn

# Ejecutar servidor
.\AGENT_AI_ENV\Scripts\python.exe -m uvicorn backend:app --reload --host 0.0.0.0 --port 8000
```

#### Terminal 2 - Ver Dashboard:
```
🌐 Abre en tu navegador: http://localhost:8000/docs
💻 O: http://localhost:8000/dashboard.html
```

---

## 📊 Rutas de la API FastAPI

### Obtener Estadísticas
```bash
GET http://localhost:8000/api/stats
```

Respuesta:
```json
{
  "total": 15,
  "pendientes": 8,
  "pagadas": 5,
  "vencidas": 2,
  "total_cop": 5000000,
  "total_usd": 1200.50
}
```

### Obtener Facturas
```bash
GET http://localhost:8000/api/invoices?limit=10
```

### Procesar Emails
```bash
POST http://localhost:8000/api/process
```

### WebSocket para Logs
```
WS ws://localhost:8000/ws/logs
```

### Estado del Sistema
```bash
GET http://localhost:8000/api/status
```

---

## 📝 Estructura de Archivos

```
AI_LANGCHAIN STPM/
├── app.py               # 🤖 Agente principal
├── backend.py           # 🚀 API FastAPI
├── dashboard.html       # 🎨 Dashboard en tiempo real
├── requirements.txt     # 📦 Dependencias
├── .env                 # 🔐 Variables de entorno
├── AGENT_AI_ENV/        # Entorno virtual
├── facturas_seguimiento.xlsx  # 📊 Excel con datos
└── processed_emails.json      # 📄 IDs procesados
```

---

## 🔧 Configuración Avanzada

### Cambiar Límite de Emails
En `app.py`, línea ~498:
```python
all_ids = all_ids[-10:] if len(all_ids) > 10 else all_ids  # Cambiar 10 por el número deseado
```

### Cambiar Modelo de IA
En `app.py`, línea 242:
```python
model="gemini-pro",  # O "gemini-2.5-flash" para más rápido
```

### Intervalo de Procesamiento
En `.env`:
```env
CHECK_INTERVAL_MINUTES=30  # Procesar cada 30 minutos
```

---

## 🚨 Solución de Problemas

### ❌ "ModuleNotFoundError: No module named 'schedule'"
```bash
pip install schedule
# O instalar todas las dependencias:
pip install -r requirements.txt
```

### ❌ "Error IMAP login"
- Verifica que `EMAIL_USER` y `EMAIL_PASS` sean correctos
- Para Gmail, DEBES usar contraseña de aplicación, no tu contraseña normal
- Ve a: https://myaccount.google.com/apppasswords

### ❌ "getaddrinfo failed"
- Verifica `IMAP_HOST=imap.gmail.com` (sin tu email)
- Verifica conexión a internet

### ❌ Dashboard en blanco
- Abre http://localhost:8000/docs (Swagger UI)
- Verifica que backend.py está ejecutándose
- Abre consola del navegador (F12) para ver errores

---

## 📈 Monitoreo

### Logs del Agente
```powershell
# En la terminal donde corre app.py verás logs como:
🔍 Revisando correos: 26/02/2026 10:23:41
📨 2 correos nuevos para analizar
📧 Procesando UID: 12345
✅ FACTURA: FAC-001 | Proveedor XYZ | 100000 COP
💾 Factura guardada → fila 5
```

### Excel
Abre `facturas_seguimiento.xlsx`:
- Hoja "📋 Facturas" con todos los datos
- Hoja "📊 Dashboard" con gráficos y resúmenes

### WebSocket en tiempo real
El dashboard muestra logs en vivo mientras se procesan emails.

---

## 🤝 API REST Completa (Swagger)

Cuando el backend esté ejecutándose, accede a:
```
http://localhost:8000/docs
```

Aquí puedes probar todas las rutas interactivamente.

---

## 📞 Soporte

Si tienes dudas:
1. Revisa los logs en la terminal
2. Verifica las variables en `.env`
3. Comprueba que tienes credenciales correctas
4. Abre el navegador con F12 para ver errores del frontend

---

## ⭐ Tips

- 💡 Mantén el dashboard abierto para monitorear en tiempo real
- 🔄 El agente se ejecuta automáticamente cada 30 minutos
- 📱 El dashboard funciona en móvil también
- 🎯 Usa la API para integraciones con otros sistemas
- 📊 Descarga los Excel regularmente para auditoría

---

**¡Listo! Tu sistema de extracción de facturas está configurado.** 🎉
