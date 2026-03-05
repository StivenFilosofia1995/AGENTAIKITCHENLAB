# start.py
import os
from pathlib import Path
import uvicorn

if __name__ == "__main__":
    port = int(os.getenv("PORT", 9000))
    print("")
    print(f" Dashboard : http://localhost:{port}/dashboard.html")
    print(f" API Docs  : http://localhost:{port}/docs")
    print("")
    uvicorn.run("backend:fastapi_app", host="0.0.0.0", port=port, reload=False, log_level="info")
```

5. Clic en **Commit changes** (botón verde)

---

### Paso 3 — Edita `backend.py`

Mismo proceso: abre el archivo, clic en ✏️, **borra todo**, y pega el contenido real del backend.

El problema es que el archivo empieza con esta línea corrupta que no es Python:
```
(cd "$(git rev-parse --show-toplevel)" && git apply --3way <<'EOF'
