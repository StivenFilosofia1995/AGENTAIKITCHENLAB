@echo off
REM Script para ejecutar el Sistema de Facturas con FastAPI + Dashboard

echo.
echo ╔════════════════════════════════════════════════════════╗
echo ║  🤖 SISTEMA DE EXTRACCIÓN DE FACTURAS                 ║
echo ║   FastAPI Backend + Dashboard en Tiempo Real          ║
echo ╚════════════════════════════════════════════════════════╝
echo.

REM Activar entorno virtual
call .\AGENT_AI_ENV\Scripts\activate.bat

echo.
echo ✅ Entorno virtual activado
echo.
echo 🚀 Iniciando FastAPI Backend en http://localhost:8000
echo.
echo    Dashboard: http://localhost:8000/dashboard.html
echo    Swagger UI: http://localhost:8000/docs
echo.
echo Presiona Ctrl+C para detener...
echo.

REM Ejecutar FastAPI
python -m uvicorn backend:app --reload --host 0.0.0.0 --port 8000
