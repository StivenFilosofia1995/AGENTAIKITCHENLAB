# 🚀 INICIO RÁPIDO - Dashboard FastAPI

## ⚡ En 3 pasos:

### 1️⃣ Abre PowerShell en esta carpeta y ejecuta:

```powershell
.\AGENT_AI_ENV\Scripts\Activate.ps1
python -m uvicorn backend:app --reload --reload-dir . --reload-include "*.py" --host 0.0.0.0 --port 9000
```

**O simplemente:**
```powershell
.\run_backend.ps1
```

### 2️⃣ Cuando veas esto:

```
INFO:     Uvicorn running on http://0.0.0.0:9000
INFO:     Application startup complete
```

### 3️⃣ Abre en tu navegador:

🎨 **Dashboard Bonito:**
```
http://localhost:9000/dashboard.html
```

📚 **API Swagger (Probar endpoints):**
```
http://localhost:9000/docs
```

---

## 🎮 En el Dashboard puedes:

✅ **Ver estadísticas en vivo:**
- Total de facturas
- Pendientes de pagar
- Facturas pagadas
- Facturas vencidas
- Totales en COP y USD

✅ **Procesar emails:**
- Click en "▶️ Procesar Emails"
- Ve los logs en tiempo real
- El flujo del agente se anima

✅ **Ver últimas facturas:**
- Fecha, proveedor, monto
- Estado de cada una
- Detalles completos

✅ **Logs en vivo:**
- Terminal del agente en el navegador
- Actualización automática
- Colores por tipo (info, error, éxito)

---

## 📊 Rutas de API para usar:

```bash
# Obtener estadísticas
curl http://localhost:9000/api/stats

# Obtener facturas
curl http://localhost:9000/api/invoices

# Procesar emails
curl -X POST http://localhost:9000/api/process

# Estado del sistema
curl http://localhost:9000/api/status

# Obtener logs
curl http://localhost:9000/api/logs
```

---

## 🔄 Flujo que ves en el Dashboard:

```
📧 Conectar IMAP → 🔍 Buscar Facturas → 🤖 Analizar IA → 📊 Guardar Excel
```

Mientras procesa, verás animaciones en cada paso.

---

## ⚙️ Configuración Automática:

El agente se ejecuta automáticamente cada 30 minutos. Puedes cambiar esto en `.env`:

```env
CHECK_INTERVAL_MINUTES=30
```

---

## 🐛 Si hay problemas:

**WebSocket no conecta:**
- Abre la consola del navegador (F12)
- Verifica que veas en consola: "WebSocket conectado"

**API responde 404:**
- Verifica que backend.py está en la misma carpeta
- Reinicia: Ctrl+C en PowerShell y ejecuta de nuevo

**Excel no se actualiza:**
- El Excel se actualiza cada vez que procesa
- En el dashboard verás "Última actualización: Ahora"

---

## 💾 Archivos generados:

- `facturas_seguimiento.xlsx` - Excel con todas las facturas
- `processed_emails.json` - IDs de emails ya procesados
- Logs en consola del backend

---

**¡Listo! Tu sistema está funcionando completamente.** 🎉
