# Activar entorno virtual
.\AGENT_AI_ENV\Scripts\Activate.ps1

Write-Host ""
Write-Host "  Dashboard : http://localhost:9000/dashboard.html"
Write-Host "  API Docs  : http://localhost:9000/docs"
Write-Host ""

# CRITICO: backend:fastapi_app (no backend:app)
python -m uvicorn backend:fastapi_app --reload --host 0.0.0.0 --port 9000