"""
FastAPI Backend — Invoice Agent
Servidor único. app.py es el módulo de procesamiento de correos.
"""
import logging
logging.getLogger("watchfiles").setLevel(logging.ERROR)

from dotenv import load_dotenv
load_dotenv()

# Importar el módulo agente (app.py — SIN circular import)
import app as app_module

from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from starlette.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
import asyncio, threading, json, os, schedule, time
from datetime import datetime
from typing import List, Dict, Optional

fastapi_app = FastAPI(title="Invoice Agent API", version="5.0.0")
fastapi_app.add_middleware(
    CORSMiddleware, allow_origins=["*"],
    allow_credentials=True, allow_methods=["*"], allow_headers=["*"],
)

# ── MODELO CHAT ───────────────────────────────────────────────────────────────
# claude-sonnet-4-6: mismo modelo que usa app.py para consistencia
# Garantiza que el chat del dashboard use el mismo nivel de inteligencia
CLAUDE_CHAT_MODEL = "claude-sonnet-4-6"

# ── LOG STORE ─────────────────────────────────────────────────────────────────
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
                self.logs = self.logs[-500:]

    def get_logs(self) -> List[Dict]:
        with self.lock:
            return self.logs.copy()

    def clear(self):
        with self.lock:
            self.logs = []


log_store = LogStore()
connected_clients: set = set()

# ── ESTADO SCHEDULER ──────────────────────────────────────────────────────────
scheduler_state = {
    "enabled": False,
    "mode": "interval",
    "interval_minutes": 30,
    "daily_time": "08:00",
    "next_run": None,
    "last_run": None,
    "running": False,
}
_scheduler_thread: Optional[threading.Thread] = None
_scheduler_stop = threading.Event()


# ── PROCESO PRINCIPAL ──────────────────────────────────────────────────────────
def run_process_with_logs():
    """Ejecuta app.process_emails() capturando prints como logs del dashboard."""
    if scheduler_state["running"]:
        log_store.add_log("Proceso ya en ejecución, espera que termine", "warning")
        return

    scheduler_state["running"] = True
    scheduler_state["last_run"] = datetime.now().strftime("%H:%M %d/%m/%Y")

    import sys, io

    class LogCapture(io.TextIOBase):
        def write(self, msg):
            msg = msg.strip()
            if msg:
                level = "error"   if any(x in msg for x in ["Error", "error", "ERROR", "❌"]) else \
                        "success" if any(x in msg for x in ["completado", "guardada", "✅", "exitosamente", "Factura guardada"]) else \
                        "warning" if any(x in msg for x in ["Warning", "warning", "⚠️"]) else "info"
                log_store.add_log(msg, level)
            return len(msg if msg else "")
        def flush(self):
            pass

    old_stdout = sys.stdout
    sys.stdout = LogCapture()

    # Vaciar processed_emails.json para que todos los correos sean "nuevos"
    processed_path = getattr(app_module, 'PROCESSED_FILE', 'processed_emails.json')
    backup_ids = set()
    try:
        if os.path.exists(processed_path):
            with open(processed_path) as f:
                backup_ids = set(json.load(f))
        with open(processed_path, 'w') as f:
            json.dump([], f)
        log_store.add_log(f"Cache limpiado — {len(backup_ids)} IDs previos en backup", "info")
    except Exception as ex:
        log_store.add_log(f"Aviso limpiando cache: {ex}", "warning")

    try:
        mes_actual = ["enero","febrero","marzo","abril","mayo","junio",
                      "julio","agosto","septiembre","octubre","noviembre","diciembre"][datetime.now().month - 1]
        year_actual = datetime.now().year
        log_store.add_log(f"Iniciando procesamiento — {mes_actual.capitalize()} {year_actual}", "info")
        log_store.add_log(f"🤖 Modelo: {app_module.CLAUDE_MODEL}", "info")
        app_module.process_emails_for_month(mes_actual, year_actual)
        invalidate_sheets_cache()  # Forzar recarga de datos actualizados
    except Exception as e:
        import traceback
        log_store.add_log(f"Error en procesamiento: {str(e)}", "error")
        log_store.add_log(traceback.format_exc(), "error")
    finally:
        sys.stdout = old_stdout
        scheduler_state["running"] = False
        # Restaurar IDs previos + nuevos encontrados en esta ejecución
        try:
            if os.path.exists(processed_path):
                with open(processed_path) as f:
                    new_ids = set(json.load(f))
                with open(processed_path, 'w') as f:
                    json.dump(list(backup_ids | new_ids), f)
        except Exception:
            pass


def _launch_process_thread():
    t = threading.Thread(
        target=run_process_with_logs,
        daemon=True,
        name="invoice-processor"
    )
    t.start()


# ── SCHEDULER ──────────────────────────────────────────────────────────────────
def _rebuild_schedule():
    schedule.clear()
    if not scheduler_state["enabled"]:
        return
    if scheduler_state["mode"] == "interval":
        mins = scheduler_state["interval_minutes"]
        schedule.every(mins).minutes.do(_launch_process_thread)
        log_store.add_log(f"Scheduler: cada {mins} minutos", "info")
    elif scheduler_state["mode"] == "daily":
        t = scheduler_state["daily_time"]
        schedule.every().day.at(t).do(_launch_process_thread)
        log_store.add_log(f"Scheduler: diario a las {t}", "info")


def _scheduler_loop(stop: threading.Event):
    while not stop.is_set():
        schedule.run_pending()
        jobs = schedule.get_jobs()
        if jobs:
            nxt = min(jobs, key=lambda j: j.next_run)
            scheduler_state["next_run"] = nxt.next_run.strftime("%H:%M %d/%m/%Y")
        else:
            scheduler_state["next_run"] = None
        stop.wait(timeout=10)


def _start_scheduler():
    global _scheduler_thread, _scheduler_stop
    if _scheduler_thread and _scheduler_thread.is_alive():
        return
    _scheduler_stop.clear()
    _scheduler_thread = threading.Thread(
        target=_scheduler_loop, args=(_scheduler_stop,),
        daemon=True, name="scheduler"
    )
    _scheduler_thread.start()


_start_scheduler()


# ── LEER EXCEL SEGURO ──────────────────────────────────────────────────────────
# Cache para evitar exceder cuota de Google Sheets API (error 429)
_sheets_cache_df = None
_sheets_cache_time = 0.0
_SHEETS_CACHE_TTL = 90  # segundos


def invalidate_sheets_cache():
    """Invalida el cache para forzar recarga en la proxima llamada."""
    global _sheets_cache_df, _sheets_cache_time
    _sheets_cache_df = None
    _sheets_cache_time = 0.0


def read_sheets_df():
    """Lee todos los registros de todos los meses. Usa cache de 90s para evitar error 429."""
    global _sheets_cache_df, _sheets_cache_time
    import pandas as pd

    # Retornar cache si sigue vigente
    if _sheets_cache_df is not None and (time.time() - _sheets_cache_time) < _SHEETS_CACHE_TTL:
        return _sheets_cache_df

    COLS = app_module.COLUMNS  # 27 columnas oficiales
    try:
        # get_all_monthly_worksheets() usa spreadsheet.worksheets() — 1 sola llamada API
        all_worksheets = app_module.get_all_monthly_worksheets()

        all_dfs = []
        for ws in all_worksheets:
            try:
                rows = app_module._api_call_with_retry(ws.get_all_values)
                if not rows or len(rows) < 2:
                    continue
                headers = rows[0]
                data    = rows[1:]
                df = pd.DataFrame(data, columns=headers)
                # Normalizar: asegurarse de tener exactamente las 27 columnas en orden correcto
                for col in COLS:
                    if col not in df.columns:
                        df[col] = ""
                df = df[COLS]  # reordenar y descartar columnas extra
                df = df.replace("", pd.NA).dropna(how="all")
                if not df.empty:
                    all_dfs.append(df)
            except Exception as e_ws:
                log_store.add_log(f"read_sheets_df error en hoja {ws.title}: {e_ws}", "warning")
                continue

        if not all_dfs:
            _sheets_cache_df = pd.DataFrame(columns=COLS)
            _sheets_cache_time = time.time()
            return _sheets_cache_df

        consolidated_df = pd.concat(all_dfs, ignore_index=True)
        # Guardar en cache con timestamp
        _sheets_cache_df = consolidated_df
        _sheets_cache_time = time.time()
        return consolidated_df
    except Exception as e_main:
        log_store.add_log(f"read_sheets_df error general: {e_main}", "error")
        # Si falla, retornar cache anterior aunque esté vencido
        return _sheets_cache_df


# ── RUTAS API ──────────────────────────────────────────────────────────────────

@fastapi_app.get("/api/health")
async def health_check():
    return {
        "status": "ok",
        "timestamp": datetime.now().isoformat(),
        "model": app_module.CLAUDE_MODEL
    }


@fastapi_app.get("/api/stats")
async def get_statistics(mes: str = None):
    import pandas as pd
    try:
        df = read_sheets_df()
        if df is None or df.empty or "Estado" not in df.columns:
            return {"total": 0, "pendientes": 0, "pagadas": 0, "vencidas": 0,
                    "total_cop": 0.0, "total_usd": 0.0, "error": None}
        if mes and "Mes" in df.columns:
            df = df[df["Mes"].str.strip().str.lower() == mes.strip().lower()]

        df["Valor Total"] = pd.to_numeric(df.get("Valor Total", df.get("Total", 0)), errors="coerce").fillna(0)
        numero_factura_col = "Número Factura"

        return {
            "total":      int(df[numero_factura_col].notna().sum()) if numero_factura_col in df.columns else len(df),
            "pendientes": int((df["Estado"] == "PENDIENTE").sum()),
            "pagadas":    int((df["Estado"] == "PAGADA").sum()),
            "vencidas":   int((df["Estado"] == "VENCIDA").sum()),
            "total_cop":  float(df["Valor Total"].sum()),
            "total_usd":  0.0,
            "error": None,
        }
    except Exception as e:
        log_store.add_log(f"Error estadísticas: {e}", "error")
        return {"total": 0, "pendientes": 0, "pagadas": 0, "vencidas": 0,
                "total_cop": 0.0, "total_usd": 0.0, "error": str(e)}


@fastapi_app.get("/api/invoices")
async def get_invoices(limit: int = 1000, mes: str = None):
    try:
        df = read_sheets_df()
        if df is None or df.empty:
            return {"invoices": [], "total": 0}
        if mes and "Mes" in df.columns:
            df = df[df["Mes"].str.strip().str.lower() == mes.strip().lower()]
        records = df.fillna("N/A").to_dict(orient="records")
        if limit:
            records = records[:limit]
        return {"invoices": records, "total": len(df)}
    except Exception as e:
        log_store.add_log(f"Error obteniendo facturas: {e}", "error")
        return {"invoices": [], "total": 0, "error": str(e)}


@fastapi_app.get("/api/months")
async def get_months():
    """Devuelve los 12 meses del año indicando cuáles tienen datos."""
    ALL_MONTHS = ["enero","febrero","marzo","abril","mayo","junio",
                  "julio","agosto","septiembre","octubre","noviembre","diciembre"]
    try:
        df = read_sheets_df()
        if df is not None and not df.empty and "Mes" in df.columns:
            with_data = set(df["Mes"].str.strip().str.lower().dropna().unique().tolist())
        else:
            with_data = set()
        result = [{"mes": m, "has_data": m in with_data} for m in ALL_MONTHS]
        return {"months": result}
    except Exception as e:
        return {"months": [{"mes": m, "has_data": False} for m in ALL_MONTHS], "error": str(e)}


@fastapi_app.get("/api/invoices/by-status/{status}")
async def get_invoices_by_status(status: str):
    try:
        df = read_sheets_df()
        if df is None or "Estado" not in df.columns:
            return {"invoices": [], "count": 0}
        df_f = df[df["Estado"] == status.upper()].fillna("N/A")
        return {"status": status.upper(), "count": len(df_f),
                "invoices": df_f.to_dict(orient="records")}
    except Exception as e:
        return {"error": str(e), "invoices": []}


@fastapi_app.post("/api/process")
async def trigger_process():
    if scheduler_state["running"]:
        return {"status": "ya_ejecutando", "message": "Proceso activo, espera que termine"}
    log_store.add_log("Procesamiento manual iniciado desde dashboard", "info")
    _launch_process_thread()
    return {"status": "iniciado", "timestamp": datetime.now().isoformat()}


@fastapi_app.post("/api/process-month")
async def trigger_process_month(body: dict):
    """Procesa todos los correos de un mes específico (ej: {'mes': 'febrero', 'year': 2026})."""
    mes = body.get("mes", "").strip()
    year = body.get("year", datetime.now().year)
    if not mes:
        return {"status": "error", "message": "Debes indicar el mes"}
    if scheduler_state["running"]:
        return {"status": "ya_ejecutando", "message": "Proceso activo, espera que termine"}

    log_store.add_log(f"Procesamiento por mes iniciado: {mes.capitalize()} {year}", "info")
    scheduler_state["running"] = True
    scheduler_state["last_run"] = datetime.now().strftime("%H:%M %d/%m/%Y")

    import sys, io

    class LogCapture(io.TextIOBase):
        def write(self, msg):
            msg = msg.strip()
            if msg:
                level = "error"   if any(x in msg for x in ["Error", "error", "ERROR", "❌"]) else \
                        "success" if any(x in msg for x in ["completado", "guardada", "✅", "exitosamente"]) else \
                        "warning" if "⚠️" in msg else "info"
                log_store.add_log(msg, level)
            return len(msg if msg else "")
        def flush(self): pass

    def _run():
        old_stdout = sys.stdout
        sys.stdout = LogCapture()
        try:
            app_module.process_emails_for_month(mes, int(year))
            invalidate_sheets_cache()
        except Exception as e:
            import traceback
            log_store.add_log(f"Error: {e}", "error")
            log_store.add_log(traceback.format_exc(), "error")
        finally:
            sys.stdout = old_stdout
            scheduler_state["running"] = False

    threading.Thread(target=_run, daemon=True).start()
    return {"status": "iniciado", "mes": mes, "year": year, "timestamp": datetime.now().isoformat()}


@fastapi_app.get("/api/logs")
async def get_logs():
    return {"logs": log_store.get_logs()}


@fastapi_app.post("/api/export-excel")
async def export_excel():
    """Genera un Excel descargable a partir de los datos en Google Sheets."""
    try:
        path = app_module.export_to_excel()
        from fastapi.responses import FileResponse
        return FileResponse(path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            filename="facturas_seguimiento.xlsx")
    except Exception as e:
        from fastapi.responses import JSONResponse
        return JSONResponse(status_code=500, content={"error": str(e)})


@fastapi_app.delete("/api/logs")
async def clear_logs():
    log_store.clear()
    return {"status": "ok"}


@fastapi_app.get("/api/status")
async def get_status():
    return {
        "agente":         "activo",
        "almacenamiento": "google_sheets",
        "sheets_id":      app_module.GOOGLE_SHEETS_ID,
        "procesando":     scheduler_state["running"],
        "clientes_ws":    len(connected_clients),
        "timestamp":      datetime.now().isoformat(),
        "logs_totales":   len(log_store.get_logs()),
        "modelo":         app_module.CLAUDE_MODEL,
    }


@fastapi_app.get("/api/scheduler")
async def get_scheduler():
    return {**scheduler_state}


@fastapi_app.post("/api/scheduler")
async def set_scheduler(config: dict):
    if "enabled"          in config: scheduler_state["enabled"]          = bool(config["enabled"])
    if "mode"             in config: scheduler_state["mode"]             = config["mode"]
    if "interval_minutes" in config: scheduler_state["interval_minutes"] = max(5, int(config["interval_minutes"]))
    if "daily_time"       in config: scheduler_state["daily_time"]       = config["daily_time"]
    _rebuild_schedule()
    estado = "activado" if scheduler_state["enabled"] else "desactivado"
    log_store.add_log(f"Scheduler {estado}", "success" if scheduler_state["enabled"] else "warning")
    return {**scheduler_state}


@fastapi_app.post("/api/chat")
async def chat_with_claude(message: dict):
    """Endpoint para chatear con Claude y ejecutar comandos mediante prompts."""
    try:
        user_msg = message.get("message", "").strip()
        if not user_msg:
            return {"error": "Mensaje vacío", "response": ""}

        import anthropic
        api_key = os.getenv("ANTHROPIC_API_KEY", "")
        if not api_key:
            return {"error": "ANTHROPIC_API_KEY no configurada", "response": ""}

        client = anthropic.Anthropic(api_key=api_key)

        system_prompt = """Eres un asistente del Invoice Agent. Respondes SIEMPRE en español.
NORMAS ESTRICTAS:
- NUNCA uses bloques de código ni texto técnico (no uses ```, no uses python, no menciones funciones)
- NUNCA finjas ejecutar código ni muestres resultados hipotéticos
- Responde en máximo 2-3 oraciones cortas y directas
- Si el usuario pide procesar un mes → solo di que lo vas a hacer y confirma el mes detectado
- Si el usuario pide procesar correos → confirma brevemente que iniciarás el procesamiento
- Si pregunta estadísticas o información → responde con los datos disponibles o pide que revise el dashboard
- NO expliques cómo funciona el sistema internamente"""

        response = client.messages.create(
            model=CLAUDE_CHAT_MODEL,   # claude-sonnet-4-6
            max_tokens=1024,
            system=system_prompt,
            messages=[{
                "role": "user",
                "content": user_msg
            }]
        )

        assistant_response = response.content[0].text if response.content else ""

        # Detectar intención y mes específico
        intent = None
        intent_mes = None
        intent_year = datetime.now().year
        lower_msg = user_msg.lower()

        # Detectar mes mencionado
        meses_map = {
            "enero": "enero", "febrero": "febrero", "marzo": "marzo",
            "abril": "abril", "mayo": "mayo", "junio": "junio",
            "julio": "julio", "agosto": "agosto", "septiembre": "septiembre",
            "octubre": "octubre", "noviembre": "noviembre", "diciembre": "diciembre"
        }
        for mes_es in meses_map:
            if mes_es in lower_msg:
                intent_mes = mes_es
                break

        # Detectar año mencionado
        import re as _re
        year_match = _re.search(r'\b(202[0-9])\b', lower_msg)
        if year_match:
            intent_year = int(year_match.group(1))

        # Si se menciona un mes específico → siempre es process_month
        if intent_mes:
            intent = "process_month"
        elif any(kw in lower_msg for kw in [
            "procesar", "procesa", "correo", "email", "factura", "gmail",
            "buscar", "busca", "extraer", "extrae", "revisar", "revisa",
            "agarra", "agarrar", "toma", "tomar", "sube", "subir",
            "agregar", "agrega", "añade", "añadir", "pasa", "pasar",
            "manda", "mandar", "junta", "juntar", "trae", "traer", "sheets"
        ]):
            intent = "process_emails"

        return {
            "response":  assistant_response,
            "intent":    intent,
            "mes":       intent_mes,
            "year":      intent_year,
            "timestamp": datetime.now().isoformat()
        }

    except Exception as e:
        return {"error": str(e), "response": f"Error al comunicar con Claude: {str(e)}"}


# ── WEBSOCKET ──────────────────────────────────────────────────────────────────
@fastapi_app.websocket("/ws/logs")
async def websocket_logs(websocket: WebSocket):
    await websocket.accept()
    connected_clients.add(websocket)
    try:
        # Enviar logs existentes al conectar
        for log in log_store.get_logs():
            await websocket.send_json(log)
        last_count = len(log_store.get_logs())

        while True:
            await asyncio.sleep(1.5)
            # Ping para mantener la conexión viva
            try:
                await websocket.send_json({
                    "type": "ping", "message": "", "level": "ping",
                    "timestamp": datetime.now().strftime("%H:%M:%S")
                })
            except Exception:
                break
            # Enviar nuevos logs
            logs = log_store.get_logs()
            for log in logs[last_count:]:
                try:
                    await websocket.send_json(log)
                except Exception:
                    break
            last_count = len(logs)

    except WebSocketDisconnect:
        pass
    except Exception:
        pass
    finally:
        connected_clients.discard(websocket)


# ── DASHBOARD ──────────────────────────────────────────────────────────────────
if os.path.exists("static"):
    fastapi_app.mount("/static", StaticFiles(directory="static"), name="static")


@fastapi_app.get("/dashboard.html", response_class=HTMLResponse)
async def get_dashboard():
    path = os.path.join(os.getcwd(), "dashboard.html")
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        return f"<h1>Error cargando dashboard: {e}</h1>"


@fastapi_app.get("/", response_class=HTMLResponse)
async def root():
    path = os.path.join(os.getcwd(), "dashboard.html")
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return f.read()
    except Exception:
        pass
    return "<h1>Invoice Agent</h1><a href='/dashboard.html'>Dashboard</a> | <a href='/docs'>API Docs</a>"


# ── ENTRY POINT ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    print("  Dashboard : http://localhost:9000/dashboard.html")
    print("  API Docs  : http://localhost:9000/docs")
    uvicorn.run(fastapi_app, host="0.0.0.0", port=9000)
