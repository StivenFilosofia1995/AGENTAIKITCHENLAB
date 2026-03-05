# start.py
import os
from pathlib import Path

import uvicorn


def _looks_like_python_start(line: str) -> bool:
    txt = (line or "").strip()
    if not txt:
        return True
    return (
        txt.startswith('"""')
        or txt.startswith("'''")
        or txt.startswith("#")
        or txt.startswith("import ")
        or txt.startswith("from ")
        or txt.startswith("def ")
        or txt.startswith("class ")
    )


def _sanitize_backend_file() -> bool:
    """Limpia encabezados corruptos en backend.py antes de importar FastAPI."""
    backend_path = Path(__file__).with_name("backend.py")
    if not backend_path.exists():
        return False

    raw = backend_path.read_text(encoding="utf-8", errors="replace")
    lines = raw.splitlines()
    if not lines:
        return False

    # Si el inicio ya parece Python válido, no tocar archivo
    if _looks_like_python_start(lines[0]):
        return False

    # Buscar primera línea que parezca código Python válido
    start_idx = None
    for idx, line in enumerate(lines):
        if _looks_like_python_start(line):
            start_idx = idx
            break

    if start_idx is None:
        return False

    cleaned = "\n".join(lines[start_idx:]).strip() + "\n"
    if cleaned == raw:
        return False

    removed = start_idx
    backend_path.write_text(cleaned, encoding="utf-8")
    print(f"⚠️ backend.py saneado automáticamente (se removieron {removed} líneas corruptas al inicio).")
    return True


if __name__ == "__main__":
    port = int(os.getenv("PORT", 9000))  # Railway inyecta PORT automáticamente
    print("")
    print(f" Dashboard : http://localhost:{port}/dashboard.html")
    print(f" API Docs : http://localhost:{port}/docs")
    print("")

    _sanitize_backend_file()

    # Import explícito tras saneamiento para fallar temprano con contexto claro
    import backend

    uvicorn.run(backend.fastapi_app, host="0.0.0.0", port=port, reload=False, log_level="info")
