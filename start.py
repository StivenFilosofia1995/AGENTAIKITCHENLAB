# start.py
import os
from pathlib import Path
import uvicorn


def _sanitize_backend_file() -> bool:
    """Elimina prefijos accidentales de comandos shell pegados en backend.py."""
    backend_path = Path(__file__).with_name("backend.py")
    if not backend_path.exists():
        return False

    lines = backend_path.read_text(encoding="utf-8", errors="replace").splitlines()
    if not lines:
        return False

    first = lines[0].strip()
    # Caso observado en Railway: línea 1 empieza con '(cd "$(git rev-parse ...'
    if not first.startswith("(cd "):
        return False

    cutoff = 0
    for idx, line in enumerate(lines, start=1):
        txt = line.strip()
        if txt.startswith('"""') or txt.startswith("#") or txt.startswith("import ") or txt.startswith("from "):
            cutoff = idx - 1
            break

    if cutoff <= 0:
        return False

    cleaned = "\n".join(lines[cutoff:]).strip() + "\n"
    backend_path.write_text(cleaned, encoding="utf-8")
    print(f"⚠️ backend.py saneado automáticamente (se removieron {cutoff} líneas corruptas al inicio).")
    return True


if __name__ == "__main__":
    port = int(os.getenv("PORT", 9000))  # Railway inyecta PORT automáticamente
    print("")
    print(f" Dashboard : http://localhost:{port}/dashboard.html")
    print(f" API Docs : http://localhost:{port}/docs")
    print("")

    _sanitize_backend_file()

    # backend:fastapi_app -> coincide con la variable FastAPI en backend.py
    uvicorn.run("backend:fastapi_app", host="0.0.0.0", port=port, reload=False, log_level="info")
