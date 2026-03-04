@echo off
cd /d "C:\Users\Stiven I.A\Desktop\AI_LANGCHAIN STPM"
chcp 65001 >nul
cls

echo.
echo  Invoice Agent — Sistema de Facturas
echo  =====================================
echo.
echo  Activando entorno virtual...
call AGENT_AI_ENV\Scripts\activate.bat

echo.
echo  Dashboard : http://localhost:9000/dashboard.html
echo  API Docs  : http://localhost:9000/docs
echo.
echo  Presiona Ctrl+C para detener
echo.

REM CRITICO: backend:fastapi_app — no backend:app
python -m uvicorn backend:fastapi_app --reload --host 0.0.0.0 --port 9000

pause