 (cd "$(git rev-parse --show-toplevel)" && git apply --3way <<'EOF' 
diff --git a/backend.py b/backend.py
index 84dcff747731b7b7232612407ba9819bc3f91f1f..5ee22c15cf66cc812ae63c2e87666dab98ea8d00 100644
--- a/backend.py
+++ b/backend.py
@@ -1,50 +1,50 @@
 """
 FastAPI Backend — Invoice Agent
 Servidor único. app.py es el módulo de procesamiento de correos.
 
 CORRECCIONES v5.1:
 - [FIX 1] Lock en running flag — evita race condition en requests simultáneos
 - [FIX 2] LogCapture unificado como clase de módulo — sin código duplicado
 - [FIX 3] process / process-month comparten _run_month_processing() unificada
 - [FIX 4] Chat con historial por sesión — Claude recuerda el contexto anterior
 - [FIX 5] read_sheets_df() siempre retorna DataFrame, nunca None
 - [FIX 6] Cache invalidation protegida con Lock — thread-safe
 """
 import logging
 logging.getLogger("watchfiles").setLevel(logging.ERROR)
 
 from dotenv import load_dotenv
 load_dotenv()
 
 import app as app_module
 
 from fastapi import FastAPI, WebSocket, WebSocketDisconnect
 from starlette.middleware.cors import CORSMiddleware
 from fastapi.responses import HTMLResponse
 from fastapi.staticfiles import StaticFiles
-import asyncio, threading, json, os, schedule, time, sys, io, re
+import asyncio, threading, json, os, schedule, time, sys, io, re, unicodedata
 from datetime import datetime
 from typing import List, Dict, Optional
 import pandas as pd
 
 fastapi_app = FastAPI(title="Invoice Agent API", version="5.1.0")
 fastapi_app.add_middleware(
     CORSMiddleware, allow_origins=["*"],
     allow_credentials=True, allow_methods=["*"], allow_headers=["*"],
 )
 
 # ── MODELO CHAT ────────────────────────────────────────────────────────────────
 CLAUDE_CHAT_MODEL = "claude-sonnet-4-6"
 
 # ── LOG STORE ──────────────────────────────────────────────────────────────────
 class LogStore:
     def __init__(self):
         self.logs: List[Dict] = []
         self.lock = threading.Lock()
 
     def add_log(self, message: str, level: str = "info", timestamp: str = None):
         if timestamp is None:
             timestamp = datetime.now().strftime("%H:%M:%S")
         with self.lock:
             self.logs.append({"message": message, "level": level, "timestamp": timestamp})
             if len(self.logs) > 500:
@@ -63,50 +63,59 @@ log_store = LogStore()
 connected_clients: set = set()
 
 # ── [FIX 1] Lock global para el flag running ──────────────────────────────────
 _process_lock = threading.Lock()
 
 # ── [FIX 2] LogCapture unificado — una sola clase, sin duplicación ────────────
 class LogCapture(io.TextIOBase):
     """Redirige print() de app.py al log_store en tiempo real."""
     def write(self, msg):
         msg = msg.strip()
         if msg:
             level = (
                 "error"   if any(x in msg for x in ["Error", "error", "ERROR", "❌"]) else
                 "success" if any(x in msg for x in ["completado", "guardada", "✅", "exitosamente", "Factura guardada"]) else
                 "warning" if any(x in msg for x in ["Warning", "warning", "⚠️"]) else
                 "info"
             )
             log_store.add_log(msg, level)
         return len(msg if msg else "")
 
     def flush(self):
         pass
 
 
 # ── ESTADO SCHEDULER ──────────────────────────────────────────────────────────
+
+
+def _normalize_month_input(mes: str) -> str:
+    """Normaliza nombres de mes para aceptar acentos y variantes comunes."""
+    txt = unicodedata.normalize("NFKD", str(mes or ""))
+    txt = "".join(c for c in txt if not unicodedata.combining(c)).lower().strip()
+    return "septiembre" if txt == "setiembre" else txt
+
+
 scheduler_state = {
     "enabled": False,
     "mode": "interval",
     "interval_minutes": 30,
     "daily_time": "08:00",
     "next_run": None,
     "last_run": None,
     "running": False,
     # ── v5.2 error tracking ──────────────────────────────────────────────
     "last_error": None,           # Último mensaje de error (None si todo OK)
     "last_error_time": None,      # Timestamp del último error
     "sheets_ok": None,            # True/False/None (None = no chequeado aún)
     "anthropic_ok": None,         # True/False/None
     "facturas_ultimo_run": 0,     # Facturas guardadas en el último procesamiento
     "completion_token": 0,        # Incrementa cada vez que termina un proceso — el dashboard lo compara para saber cuándo terminó
 }
 _scheduler_thread: Optional[threading.Thread] = None
 _scheduler_stop = threading.Event()
 
 
 # ── [FIX 3] Función de procesamiento unificada ────────────────────────────────
 def _run_month_processing(mes: str, year: int):
     """
     Núcleo único de procesamiento. Lo usan:
       - El scheduler automático
@@ -437,51 +446,51 @@ async def get_months():
 @fastapi_app.get("/api/invoices/by-status/{status}")
 async def get_invoices_by_status(status: str):
     try:
         df = read_sheets_df()
         if df.empty or "Estado" not in df.columns:
             return {"invoices": [], "count": 0}
         df_f = df[df["Estado"] == status.upper()].fillna("N/A")
         return {"status": status.upper(), "count": len(df_f),
                 "invoices": df_f.to_dict(orient="records")}
     except Exception as e:
         return {"error": str(e), "invoices": []}
 
 
 @fastapi_app.post("/api/process")
 async def trigger_process():
     with _process_lock:  # [FIX 1]
         if scheduler_state["running"]:
             return {"status": "ya_ejecutando", "message": "Proceso activo, espera que termine"}
     log_store.add_log("Procesamiento manual iniciado desde dashboard", "info")
     _launch_process_thread()
     return {"status": "iniciado", "timestamp": datetime.now().isoformat()}
 
 
 @fastapi_app.post("/api/process-month")
 async def trigger_process_month(body: dict):
-    mes  = body.get("mes", "").strip()
+    mes  = _normalize_month_input(body.get("mes", ""))
     year = int(body.get("year", datetime.now().year))
     if not mes:
         return {"status": "error", "message": "Debes indicar el mes"}
     with _process_lock:  # [FIX 1]
         if scheduler_state["running"]:
             return {"status": "ya_ejecutando", "message": "Proceso activo, espera que termine"}
     log_store.add_log(f"Procesamiento por mes: {mes.capitalize()} {year}", "info")
     _launch_process_thread(mes=mes, year=year)  # [FIX 3] función unificada
     return {"status": "iniciado", "mes": mes, "year": year, "timestamp": datetime.now().isoformat()}
 
 
 @fastapi_app.get("/api/logs")
 async def get_logs():
     return {"logs": log_store.get_logs()}
 
 
 @fastapi_app.delete("/api/logs")
 async def clear_logs():
     log_store.clear()
     return {"status": "ok"}
 
 
 @fastapi_app.get("/api/status")
 async def get_status():
     return {
@@ -573,52 +582,53 @@ NORMAS:
         response = client.messages.create(
             model=CLAUDE_CHAT_MODEL,
             max_tokens=1024,
             system=system_prompt,
             messages=history  # [FIX 4] historial completo
         )
 
         assistant_response = response.content[0].text if response.content else ""
 
         # Guardar turno completo en el historial
         history.append({"role": "assistant", "content": assistant_response})
         _save_chat_history(session_id, history)
 
         # Detectar intención
         intent      = None
         intent_mes  = None
         intent_year = datetime.now().year
         lower_msg   = user_msg.lower()
 
         meses_map = {
             "enero": "enero", "febrero": "febrero", "marzo": "marzo",
             "abril": "abril", "mayo": "mayo", "junio": "junio",
             "julio": "julio", "agosto": "agosto", "septiembre": "septiembre",
             "octubre": "octubre", "noviembre": "noviembre", "diciembre": "diciembre"
         }
+        lower_msg_norm = _normalize_month_input(lower_msg)
         for mes_es in meses_map:
-            if mes_es in lower_msg:
+            if mes_es in lower_msg_norm:
                 intent_mes = mes_es
                 break
 
         year_match = re.search(r'\b(202[0-9])\b', lower_msg)
         if year_match:
             intent_year = int(year_match.group(1))
 
         if intent_mes:
             intent = "process_month"
         elif any(kw in lower_msg for kw in [
             "procesar", "procesa", "correo", "email", "factura", "gmail",
             "buscar", "busca", "extraer", "extrae", "revisar", "revisa",
             "agarra", "toma", "sube", "agregar", "agrega", "añade",
             "pasa", "manda", "junta", "trae", "sheets"
         ]):
             intent = "process_emails"
 
         return {
             "response":   assistant_response,
             "intent":     intent,
             "mes":        intent_mes,
             "year":       intent_year,
             "session_id": session_id,
             "timestamp":  datetime.now().isoformat(),
         }
 
EOF
)
