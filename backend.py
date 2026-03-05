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
import asyncio, threading, json, os, schedule, time, sys, io, re, unicodedata
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
                self.logs = self.logs[-500:]

    def get_logs(self) -> List[Dict]:
        with self.lock:
            return self.logs.copy()

    def clear(self):
        with self.lock:
            self.logs = []


log_store = LogStore()
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


def _normalize_month_input(mes: str) -> str:
    """Normaliza nombres de mes para aceptar acentos y variantes comunes."""
    txt = unicodedata.normalize("NFKD", str(mes or ""))
    txt = "".join(c for c in txt if not unicodedata.combining(c)).lower().strip()
    return "septiembre" if txt == "setiembre" else txt


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
      - El botón principal del dashboard (/api/process)
      - El procesamiento por mes (/api/process-month)
    Evita duplicar lógica y garantiza comportamiento idéntico en todos los casos.
    """
    # [FIX 1] Verificar y setear flag con lock atómico
    with _process_lock:
        if scheduler_state["running"]:
            log_store.add_log("Proceso ya en ejecución, espera que termine", "warning")
            return False
        scheduler_state["running"] = True
        scheduler_state["last_run"] = datetime.now().strftime("%H:%M %d/%m/%Y")
        scheduler_state["last_error"] = None  # limpiar error anterior

    processed_path = getattr(app_module, "PROCESSED_FILE", "processed_emails.json")
    backup_ids = set()
    old_stdout = sys.stdout
    sys.stdout = LogCapture()  # [FIX 2] Usa la clase unificada

    try:
        # Limpiar cache para reprocesar correos
        if os.path.exists(processed_path):
            with open(processed_path) as f:
                backup_ids = set(json.load(f))
        with open(processed_path, "w") as f:
            json.dump([], f)

        log_store.add_log(f"Cache limpiado — {len(backup_ids)} IDs previos en backup", "info")
        log_store.add_log(f"🤖 Modelo: {app_module.CLAUDE_MODEL}", "info")
        log_store.add_log(f"📅 Procesando: {mes.capitalize()} {year}", "info")

        # ── [v5.2] Verificar Google Sheets antes de arrancar ──────────────
        try:
            app_module.setup_sheets()
            scheduler_state["sheets_ok"] = True
        except Exception as e_sheets:
            msg = f"❌ No se pudo conectar a Google Sheets: {e_sheets}"
            log_store.add_log(msg, "error")
            log_store.add_log("Verifica el archivo service_account.json y los permisos del Sheet", "error")
            scheduler_state["sheets_ok"] = False
            scheduler_state["last_error"] = msg
            scheduler_state["last_error_time"] = datetime.now().isoformat()
            return False

        # ── [v5.2] Verificar API key de Anthropic ─────────────────────────
        if not getattr(app_module, "ANTHROPIC_API_KEY", ""):
            msg = "❌ ANTHROPIC_API_KEY no está configurada en las variables de entorno"
            log_store.add_log(msg, "error")
            scheduler_state["anthropic_ok"] = False
            scheduler_state["last_error"] = msg
            scheduler_state["last_error_time"] = datetime.now().isoformat()
            return False
        scheduler_state["anthropic_ok"] = True

        app_module.process_emails_for_month(mes, year)
        invalidate_sheets_cache()

        # Contar facturas guardadas en este run (leyendo logs recientes)
        recent_logs = log_store.get_logs()
        facturas = sum(1 for l in recent_logs if "Factura guardada" in l.get("message", ""))
        scheduler_state["facturas_ultimo_run"] = facturas
        log_store.add_log(f"✅ Procesamiento completado — {facturas} facturas nuevas guardadas", "success")

    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        # [v5.2] Detectar tipos de error conocidos para dar mensajes útiles
        err_str = str(e)
        if "529" in err_str or "overloaded" in err_str.lower():
            msg = "❌ Anthropic API sobrecargada (error 529). Intenta en unos minutos."
        elif "quota" in err_str.lower() or "429" in err_str:
            msg = "❌ Google Sheets cuota excedida. El retry automático falló — intenta más tarde."
        elif "auth" in err_str.lower() or "credential" in err_str.lower():
            msg = "❌ Error de autenticación con Google. Verifica service_account.json."
            scheduler_state["sheets_ok"] = False
        elif "imap" in err_str.lower() or "login" in err_str.lower():
            msg = "❌ Error de conexión a Gmail. Verifica EMAIL_USER y EMAIL_PASS en el .env."
        else:
            msg = f"❌ Error en procesamiento: {err_str[:120]}"
        log_store.add_log(msg, "error")
        scheduler_state["last_error"] = msg
        scheduler_state["last_error_time"] = datetime.now().isoformat()
        log_store.add_log(tb[:800], "error")
    finally:
        sys.stdout = old_stdout
        # Restaurar IDs previos + nuevos detectados en esta ejecución
        try:
            new_ids = set()
            if os.path.exists(processed_path):
                with open(processed_path) as f:
                    new_ids = set(json.load(f))
            with open(processed_path, "w") as f:
                json.dump(list(backup_ids | new_ids), f)
        except Exception:
            pass
        # [FIX 1] Liberar flag con lock + incrementar completion_token
        with _process_lock:
            scheduler_state["running"] = False
            scheduler_state["completion_token"] = scheduler_state.get("completion_token", 0) + 1

    return True


def _launch_process_thread(mes: str = None, year: int = None):
    """Lanza el procesamiento en background. Sin mes → usa el mes actual."""
    if mes is None:
        mes = ["enero","febrero","marzo","abril","mayo","junio",
               "julio","agosto","septiembre","octubre","noviembre","diciembre"][datetime.now().month - 1]
    if year is None:
        year = datetime.now().year

    threading.Thread(
        target=_run_month_processing,
        args=(mes, year),
        daemon=True,
        name=f"invoice-{mes}-{year}"
    ).start()


def run_process_with_logs():
    """Entry point del scheduler — procesa el mes actual."""
    mes  = ["enero","febrero","marzo","abril","mayo","junio",
            "julio","agosto","septiembre","octubre","noviembre","diciembre"][datetime.now().month - 1]
    _run_month_processing(mes, datetime.now().year)


# ── SCHEDULER ──────────────────────────────────────────────────────────────────
def _rebuild_schedule():
    schedule.clear()
    if not scheduler_state["enabled"]:
        return
    if scheduler_state["mode"] == "interval":
        mins = scheduler_state["interval_minutes"]
        schedule.every(mins).minutes.do(run_process_with_logs)
        log_store.add_log(f"Scheduler activado: cada {mins} minutos", "info")
    elif scheduler_state["mode"] == "daily":
        t = scheduler_state["daily_time"]
        schedule.every().day.at(t).do(run_process_with_logs)
        log_store.add_log(f"Scheduler activado: diario a las {t}", "info")


def _scheduler_loop(stop: threading.Event):
    while not stop.is_set():
        schedule.run_pending()
        jobs = schedule.get_jobs()
        scheduler_state["next_run"] = (
            min(jobs, key=lambda j: j.next_run).next_run.strftime("%H:%M %d/%m/%Y")
            if jobs else None
        )
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


# ── [FIX 5 + FIX 6] Cache de Sheets thread-safe, nunca retorna None ──────────
_sheets_cache_df: Optional[pd.DataFrame] = None
_sheets_cache_time: float = 0.0
_sheets_cache_lock = threading.Lock()   # [FIX 6]
_SHEETS_CACHE_TTL = 90  # segundos


def invalidate_sheets_cache():
    """Invalida el cache de forma thread-safe. [FIX 6]"""
    global _sheets_cache_df, _sheets_cache_time
    with _sheets_cache_lock:
        _sheets_cache_df = None
        _sheets_cache_time = 0.0


def read_sheets_df() -> pd.DataFrame:
    """
    Lee todos los registros mensuales de Google Sheets.
    [FIX 5] Siempre retorna un DataFrame (nunca None).
    [FIX 6] Escritura en cache protegida con Lock.
    """
    global _sheets_cache_df, _sheets_cache_time

    # Retornar desde cache si sigue vigente
    with _sheets_cache_lock:
        if _sheets_cache_df is not None and (time.time() - _sheets_cache_time) < _SHEETS_CACHE_TTL:
            return _sheets_cache_df.copy()

    COLS = app_module.COLUMNS
    empty_df = pd.DataFrame(columns=COLS)  # [FIX 5] fallback siempre válido

    try:
        all_worksheets = app_module.get_all_monthly_worksheets()
        all_dfs = []

        for ws in all_worksheets:
            try:
                rows = app_module._api_call_with_retry(ws.get_all_values)
                if not rows or len(rows) < 2:
                    continue
                df = pd.DataFrame(rows[1:], columns=rows[0])
                for col in COLS:
                    if col not in df.columns:
                        df[col] = ""
                df = df[COLS].replace("", pd.NA).dropna(how="all")
                if not df.empty:
                    all_dfs.append(df)
            except Exception as e_ws:
                log_store.add_log(f"Error leyendo hoja {ws.title}: {e_ws}", "warning")

        result = pd.concat(all_dfs, ignore_index=True) if all_dfs else empty_df

        with _sheets_cache_lock:  # [FIX 6]
            _sheets_cache_df = result
            _sheets_cache_time = time.time()

        return result.copy()

    except Exception as e_main:
        log_store.add_log(f"read_sheets_df error general: {e_main}", "error")
        # [FIX 5] Retorna cache anterior o DataFrame vacío, NUNCA None
        with _sheets_cache_lock:
            return _sheets_cache_df.copy() if _sheets_cache_df is not None else empty_df


# ── [FIX 4] Historial de chat por sesión ──────────────────────────────────────
_chat_histories: Dict[str, List[Dict]] = {}
_chat_lock = threading.Lock()
_MAX_CHAT_TURNS = 20  # 20 pares = 40 mensajes máx por sesión


def _get_chat_history(session_id: str) -> List[Dict]:
    with _chat_lock:
        return _chat_histories.get(session_id, []).copy()


def _save_chat_history(session_id: str, messages: List[Dict]):
    with _chat_lock:
        # Mantener solo los últimos N turnos para controlar tokens
        _chat_histories[session_id] = messages[-(_MAX_CHAT_TURNS * 2):]


def _clear_chat_history(session_id: str):
    with _chat_lock:
        _chat_histories.pop(session_id, None)


# ── RUTAS API ──────────────────────────────────────────────────────────────────

@fastapi_app.get("/api/health")
async def health_check():
    return {
        "status":    "ok",
        "version":   "5.1.0",
        "modelo":    app_module.CLAUDE_MODEL,
        "timestamp": datetime.now().isoformat(),
    }


@fastapi_app.get("/api/stats")
async def get_statistics(mes: str = None):
    try:
        df = read_sheets_df()  # [FIX 5] nunca None
        if df.empty or "Estado" not in df.columns:
            return {"total": 0, "pendientes": 0, "pagadas": 0, "vencidas": 0,
                    "total_cop": 0.0, "total_usd": 0.0, "error": None}
        if mes and "Mes" in df.columns:
            df = df[df["Mes"].str.strip().str.lower() == mes.strip().lower()]

        df["Valor Total"] = pd.to_numeric(
            df.get("Valor Total", df.get("Total", 0)), errors="coerce"
        ).fillna(0)

        return {
            "total":      int(df["Número Factura"].notna().sum()) if "Número Factura" in df.columns else len(df),
            "pendientes": int((df["Estado"] == "PENDIENTE").sum()),
            "pagadas":    int((df["Estado"] == "PAGADA").sum()),
            "vencidas":   int((df["Estado"] == "VENCIDA").sum()),
            "total_cop":  float(df["Valor Total"].sum()),
            "total_usd":  0.0,
            "error":      None,
        }
    except Exception as e:
        log_store.add_log(f"Error estadísticas: {e}", "error")
        return {"total": 0, "pendientes": 0, "pagadas": 0, "vencidas": 0,
                "total_cop": 0.0, "total_usd": 0.0, "error": str(e)}


@fastapi_app.get("/api/invoices")
async def get_invoices(limit: int = 1000, mes: str = None):
    try:
        df = read_sheets_df()
        if df.empty:
            return {"invoices": [], "total": 0}
        if mes and "Mes" in df.columns:
            df = df[df["Mes"].str.strip().str.lower() == mes.strip().lower()]
        records = df.fillna("N/A").to_dict(orient="records")
        return {"invoices": records[:limit] if limit else records, "total": len(df)}
    except Exception as e:
        log_store.add_log(f"Error obteniendo facturas: {e}", "error")
        return {"invoices": [], "total": 0, "error": str(e)}


@fastapi_app.get("/api/months")
async def get_months():
    ALL_MONTHS = ["enero","febrero","marzo","abril","mayo","junio",
                  "julio","agosto","septiembre","octubre","noviembre","diciembre"]
    try:
        df = read_sheets_df()
        with_data = (
            set(df["Mes"].str.strip().str.lower().dropna().unique())
            if not df.empty and "Mes" in df.columns else set()
        )
        return {"months": [{"mes": m, "has_data": m in with_data} for m in ALL_MONTHS]}
    except Exception as e:
        return {"months": [{"mes": m, "has_data": False} for m in ALL_MONTHS], "error": str(e)}


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
    mes  = _normalize_month_input(body.get("mes", ""))
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
        "agente":               "activo",
        "version":              "5.2.0",
        "almacenamiento":       "google_sheets",
        "sheets_id":            app_module.GOOGLE_SHEETS_ID,
        "procesando":           scheduler_state["running"],
        "clientes_ws":          len(connected_clients),
        "modelo":               app_module.CLAUDE_MODEL,
        "timestamp":            datetime.now().isoformat(),
        "logs_totales":         len(log_store.get_logs()),
        # ── v5.2 health / error ──────────────────────────────────────────
        "sheets_ok":            scheduler_state["sheets_ok"],
        "anthropic_ok":         scheduler_state["anthropic_ok"],
        "last_error":           scheduler_state["last_error"],
        "last_error_time":      scheduler_state["last_error_time"],
        "facturas_ultimo_run":  scheduler_state["facturas_ultimo_run"],
        "completion_token":     scheduler_state["completion_token"],
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


@fastapi_app.post("/api/export-excel")
async def export_excel():
    try:
        path = app_module.export_to_excel()
        from fastapi.responses import FileResponse
        return FileResponse(
            path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="facturas_seguimiento.xlsx"
        )
    except Exception as e:
        from fastapi.responses import JSONResponse
        return JSONResponse(status_code=500, content={"error": str(e)})


# ── [FIX 4] Chat con historial ─────────────────────────────────────────────────
@fastapi_app.post("/api/chat")
async def chat_with_claude(message: dict):
    """
    Chat con Claude manteniendo historial por sesión.
    [FIX 4] Claude recuerda el contexto de los últimos 20 turnos.
    """
    try:
        user_msg   = message.get("message", "").strip()
        session_id = message.get("session_id", "default")

        if not user_msg:
            return {"error": "Mensaje vacío", "response": ""}

        api_key = os.getenv("ANTHROPIC_API_KEY", "")
        if not api_key:
            return {"error": "ANTHROPIC_API_KEY no configurada", "response": ""}

        import anthropic
        client = anthropic.Anthropic(api_key=api_key)

        system_prompt = """Eres un asistente del Invoice Agent. Respondes SIEMPRE en español.
Recuerdas el contexto completo de la conversación para dar respuestas coherentes y continuas.
NORMAS:
- NUNCA uses bloques de código ni texto técnico
- Responde en máximo 2-3 oraciones cortas y directas
- Si el usuario pide procesar un mes → confirma el mes detectado
- Si el usuario pide procesar correos → confirma que iniciarás el procesamiento
- Puedes referirte a mensajes anteriores de la conversación cuando sea útil"""

        # [FIX 4] Cargar historial, agregar mensaje, enviar todo a Claude
        history = _get_chat_history(session_id)
        history.append({"role": "user", "content": user_msg})

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
        lower_msg_norm = _normalize_month_input(lower_msg)
        for mes_es in meses_map:
            if mes_es in lower_msg_norm:
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

    except Exception as e:
        return {"error": str(e), "response": f"Error al comunicar con Claude: {str(e)}"}


@fastapi_app.delete("/api/chat/{session_id}")
async def clear_chat_history_endpoint(session_id: str):
    """Limpia el historial de una sesión de chat específica."""
    _clear_chat_history(session_id)
    return {"status": "ok", "session_id": session_id}


# ── WEBSOCKET ──────────────────────────────────────────────────────────────────
@fastapi_app.websocket("/ws/logs")
async def websocket_logs(websocket: WebSocket):
    await websocket.accept()
    connected_clients.add(websocket)
    try:
        for log in log_store.get_logs():
            await websocket.send_json(log)
        last_count = len(log_store.get_logs())
        while True:
            await asyncio.sleep(1.5)
            try:
                await websocket.send_json({
                    "type": "ping", "message": "", "level": "ping",
                    "timestamp": datetime.now().strftime("%H:%M:%S")
                })
            except Exception:
                break
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
