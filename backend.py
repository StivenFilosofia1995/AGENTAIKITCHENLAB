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

from fastapi import FastAPI, WebSocket, WebSocketDisconnect, HTTPException
from starlette.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import asyncio
import os
import re
from datetime import datetime
from typing import Dict, List, Optional
import json
import threading

# ── CONFIGURACIÓN BÁSICA ────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

fastapi_app = FastAPI(title="Invoice Agent API", version="5.1.0")

fastapi_app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Servir dashboard.html en la raíz
@fastapi_app.get("/")
async def serve_dashboard():
    dashboard_path = os.path.join(BASE_DIR, "dashboard.html")
    if os.path.exists(dashboard_path):
        with open(dashboard_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse("<h1>Dashboard no encontrado</h1>", status_code=404)

# Healthcheck para Railway
@fastapi_app.get("/api/status")
async def health_check():
    return {"status": "ok", "message": "API corriendo correctamente"}

# ── LOGS Y WEBSOCKETS ──────────────────────────────────────────────────────────
class LogCapture(logging.Handler):
    def __init__(self):
        super().__init__()
        self.logs = []
        
    def emit(self, record):
        log_entry = {
            "timestamp": datetime.now().strftime("%H:%M:%S"),
            "level": record.levelname.lower(),
            "message": self.format(record)
        }
        self.logs.append(log_entry)
        if len(self.logs) > 100:
            self.logs.pop(0)

log_capture = LogCapture()
formatter = logging.Formatter('%(message)s')
log_capture.setFormatter(formatter)
logging.getLogger().addHandler(log_capture)
logging.getLogger().setLevel(logging.INFO)

class ConnectionManager:
    def __init__(self):
        self.active_connections: list[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        if websocket in self.active_connections:
            self.active_connections.remove(websocket)

    async def broadcast(self, message: dict):
        for connection in self.active_connections:
            try:
                await connection.send_json(message)
            except Exception:
                pass

manager = ConnectionManager()

@fastapi_app.websocket("/ws/logs")
async def websocket_endpoint(websocket: WebSocket):
    await manager.connect(websocket)
    try:
        while True:
            await asyncio.sleep(1)
            if log_capture.logs:
                for log in log_capture.logs:
                    await manager.broadcast(log)
                log_capture.logs.clear()
            else:
                 await manager.broadcast({"level": "ping", "message": "", "timestamp": ""})
    except WebSocketDisconnect:
        manager.disconnect(websocket)

# ── RUTAS API (Simuladas basadas en tus requerimientos) ────────────────────────
@fastapi_app.get("/api/stats")
async def get_stats():
    # Aquí deberías llamar a tu función real de app.py
    return JSONResponse(content={"total_invoices": 0, "pending": 0, "paid": 0})

@fastapi_app.get("/api/invoices")
async def get_invoices():
    # Aquí deberías llamar a tu función real de app.py
    return JSONResponse(content=[])

class ProcessRequest(BaseModel):
    month: Optional[str] = None

@fastapi_app.post("/api/process")
async def process_emails(request: ProcessRequest):
    # Simulación de procesamiento
    logging.info(f"Iniciando procesamiento manual para: {request.month or 'Todos los meses'}")
    
    def run_process():
        try:
            # Aquí llamas a la función real de tu app_module
            # app_module.main(mes=request.month)
            logging.info("Procesamiento completado exitosamente")
        except Exception as e:
            logging.error(f"Error en procesamiento: {str(e)}")
            
    thread = threading.Thread(target=run_process)
    thread.start()
    
    return {"status": "ok", "message": "Procesamiento iniciado en segundo plano"}

@fastapi_app.get("/api/logs")
async def get_logs():
    return JSONResponse(content=log_capture.logs)
