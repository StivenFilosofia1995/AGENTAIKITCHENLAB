# start.py 
import os
import uvicorn

if __name__ == "__main__":
    port = int(os.getenv("PORT", 9000))  # Railway inyecta PORT automáticamente
    print("")
    print(f" Dashboard : http://localhost:{port}/dashboard.html")
    print(f" API Docs : http://localhost:{port}/docs")
    print("")
    # backend:fastapi_app -> coincide con la variable FastAPI en backend.py
    uvicorn.run("backend:fastapi_app", host="0.0.0.0", port=port, reload=False, log_level="info")
