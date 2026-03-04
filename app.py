"""
Módulo agente de extracción de facturas.
- Conexión IMAP a Gmail
- Extracción de datos con Gemini AI
- Procesamiento de PDF, DOCX, XLSX adjuntos
- Guardado en Excel
NO contiene rutas FastAPI — eso es responsabilidad de backend.py
"""

from dotenv import load_dotenv
load_dotenv()

import os
import json
import time
import imaplib
import email
from email.header import decode_header
import base64
import io
import re
from datetime import datetime
from typing import Optional

# ── CONSTANTES ────────────────────────────────────────────────────────────────
_BASE_DIR            = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH           = os.getenv("EXCEL_OUTPUT_PATH", os.path.join(_BASE_DIR, "facturas_seguimiento.xlsx"))
PROCESSED_FILE       = os.path.join(_BASE_DIR, "processed_emails.json")
IMAP_HOST            = os.getenv("IMAP_HOST", "imap.gmail.com")
IMAP_PORT            = int(os.getenv("IMAP_PORT", "993"))
IMAP_FOLDER          = os.getenv("IMAP_FOLDER", "INBOX")
EMAIL_USER           = os.getenv("EMAIL_USER", "")
EMAIL_PASS           = os.getenv("EMAIL_PASS", "")
ANTHROPIC_API_KEY    = os.getenv("ANTHROPIC_API_KEY", "")
GOOGLE_SHEETS_ID     = os.getenv("GOOGLE_SHEETS_ID", "")
SERVICE_ACCOUNT_FILE = os.path.join(_BASE_DIR, os.getenv("SERVICE_ACCOUNT_FILE", "service_account.json"))
SHEET_NAME           = "Facturas"

COLUMNS = [
    "Mes", "Fecha Factura", "Número Factura", "Proveedor", "ID", "Número ID",
    "Subtotal", "Descuento", "IVA", "Rete IVA", "Rete ICA", "Impto Consumo",
    "Propina", "Otros Impuestos", "Retención en la fuente", "Administración",
    "Utilidad", "Imprevistos", "Valor Total", "Clasificación", "Estado",
    "Valor Pagado", "Valor Por Pagar", "Fecha Pago", "Cliente",
    "Cotización Inventto", "Observaciones"
]

# ── GOOGLE SHEETS CLIENT ──────────────────────────────────────────────────────
_gspread_client = None
_gspread_spreadsheet = None  # Cache del spreadsheet para evitar open_by_key repetidos
_gspread_worksheets = {}     # Cache de worksheets por nombre {nombre: ws}

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def _api_call_with_retry(fn, *args, max_retries=5, **kwargs):
    """Ejecuta una llamada a la API de Google con retry exponencial en caso de 429."""
    import gspread
    for attempt in range(max_retries):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            if "429" in str(e) or "Quota" in str(e):
                wait = (2 ** attempt) * 2  # 2s, 4s, 8s, 16s, 32s
                print(f"⚠️  Google Sheets cuota excedida. Esperando {wait}s antes de reintentar ({attempt+1}/{max_retries})...")
                time.sleep(wait)
            else:
                raise
    raise Exception("❌ Límite de reintentos alcanzado para Google Sheets API")


def _get_spreadsheet():
    """Retorna el spreadsheet cacheado. Llama open_by_key solo una vez por sesión."""
    global _gspread_client, _gspread_spreadsheet
    import gspread
    from google.oauth2.service_account import Credentials

    SCOPES = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]

    if _gspread_client is None:
        creds_json_env = os.getenv("GOOGLE_CREDENTIALS_JSON")
        if creds_json_env:
            import json as _json
            creds_dict = _json.loads(creds_json_env)
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        else:
            creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        _gspread_client = gspread.authorize(creds)

    if _gspread_spreadsheet is None:
        _gspread_spreadsheet = _api_call_with_retry(_gspread_client.open_by_key, GOOGLE_SHEETS_ID)

    return _gspread_spreadsheet


def get_sheet(mes_nombre: str = None):
    """Retorna la hoja de Google Sheets del mes especificado, creándola si no existe."""
    global _gspread_worksheets
    import gspread

    if mes_nombre is None:
        mes_nombre = MESES[datetime.now().month - 1]

    # Retornar desde cache si ya se obtuvo antes
    if mes_nombre in _gspread_worksheets:
        return _gspread_worksheets[mes_nombre]

    spreadsheet = _get_spreadsheet()

    try:
        ws = _api_call_with_retry(spreadsheet.worksheet, mes_nombre)
    except gspread.WorksheetNotFound:
        ws = _api_call_with_retry(spreadsheet.add_worksheet,
                                  title=mes_nombre, rows=1000, cols=len(COLUMNS))
        _api_call_with_retry(ws.append_row, COLUMNS)
        print(f"✅ Hoja '{mes_nombre}' creada con encabezados")

    _gspread_worksheets[mes_nombre] = ws
    return ws


def get_all_monthly_worksheets():
    """Obtiene todas las hojas mensuales existentes en UNA sola llamada a la API."""
    global _gspread_worksheets
    spreadsheet = _get_spreadsheet()
    # Una sola llamada a la API para listar todas las hojas
    all_ws = _api_call_with_retry(spreadsheet.worksheets)
    result = []
    for ws in all_ws:
        if ws.title in MESES:
            _gspread_worksheets[ws.title] = ws  # Actualizar cache
            result.append(ws)
    return result


def setup_sheets():
    """Verifica conexión a Google Sheets y crea hoja del mes actual si no existe."""
    try:
        mes_actual = MESES[datetime.now().month - 1]
        ws = get_sheet(mes_actual)  # Usa cache interno — NO repite open_by_key
        print(f"✅ Google Sheets conectado: Hoja '{mes_actual}'")
    except Exception as e:
        print(f"❌ Error conectando a Google Sheets: {e}")
        raise


def _load_processed_ids() -> set:
    try:
        if os.path.exists(PROCESSED_FILE):
            with open(PROCESSED_FILE) as f:
                return set(json.load(f))
    except Exception:
        pass
    return set()


def _save_processed_ids(ids: set):
    try:
        with open(PROCESSED_FILE, "w") as f:
            json.dump(list(ids), f)
    except Exception as e:
        print(f"⚠️ No se pudo guardar IDs procesados: {e}")


# ── IMAP ──────────────────────────────────────────────────────────────────────
def connect_imap() -> imaplib.IMAP4_SSL:
    mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select(IMAP_FOLDER)
    print(f"✅ Conectado a {IMAP_HOST} como {EMAIL_USER}")
    return mail


def search_invoice_emails(mail: imaplib.IMAP4_SSL) -> list:
    """Devuelve los IDs de los 50 correos más recientes."""
    try:
        _, data = mail.search(None, "ALL")
        all_ids = [uid.decode() for uid in data[0].split()]
        recientes = all_ids[-50:]
        print(f"📬 {len(recientes)} correos para analizar")
        return recientes
    except Exception as e:
        print(f"❌ Error buscando correos: {e}")
        return []


# Mapa de meses español → número
_MES_NUM = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
}
# Nombres en inglés para IMAP (formato requerido: "01-Jan-2026")
_MES_EN = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}


def search_emails_by_month(mail: imaplib.IMAP4_SSL, mes_nombre: str, year: int = None) -> list:
    """Busca en Gmail todos los correos del mes especificado usando filtro IMAP por fecha."""
    import calendar
    mes_nombre = mes_nombre.lower().strip()
    mes_num = _MES_NUM.get(mes_nombre)
    if not mes_num:
        print(f"❌ Mes no reconocido: {mes_nombre}")
        return []

    if year is None:
        year = datetime.now().year

    # Calcular rango de fechas del mes
    ultimo_dia = calendar.monthrange(year, mes_num)[1]
    fecha_inicio = f"01-{_MES_EN[mes_num]}-{year}"
    # Primer día del mes siguiente
    mes_siguiente = mes_num + 1 if mes_num < 12 else 1
    year_siguiente = year if mes_num < 12 else year + 1
    fecha_fin = f"01-{_MES_EN[mes_siguiente]}-{year_siguiente}"

    try:
        # IMAP SINCE/BEFORE para filtrar por rango de fechas
        criterio = f'(SINCE "{fecha_inicio}" BEFORE "{fecha_fin}")'
        _, data = mail.search(None, criterio)
        ids = [uid.decode() for uid in data[0].split() if uid]
        print(f"📬 {len(ids)} correos encontrados en {mes_nombre.capitalize()} {year}")
        return ids
    except Exception as e:
        print(f"❌ Error buscando correos por mes: {e}")
        return []


def process_emails_for_month(mes_nombre: str, year: int = None):
    """Procesa TODOS los correos de un mes específico, ignorando el cache de procesados."""
    if not EMAIL_USER or not EMAIL_PASS:
        print("❌ Credenciales de correo no configuradas en .env")
        return
    if not ANTHROPIC_API_KEY:
        print("❌ ANTHROPIC_API_KEY no configurada en .env")
        return

    mes_nombre_cap = mes_nombre.strip().capitalize()
    if year is None:
        year = datetime.now().year

    print(f"🗓️  Iniciando procesamiento del mes: {mes_nombre_cap} {year}")
    setup_sheets()

    try:
        mail = connect_imap()
    except Exception as e:
        print(f"❌ Error de conexión IMAP: {e}")
        return

    email_ids = search_emails_by_month(mail, mes_nombre, year)
    if not email_ids:
        print(f"📭 No se encontraron correos en {mes_nombre_cap} {year}")
        return

    # Para procesamiento por mes NO se usa el cache de IDs procesados
    # — se reprocesa todo el mes para asegurar completitud
    facturas_encontradas = 0
    for eid in email_ids:
        try:
            _, raw_data = mail.fetch(eid, "(RFC822)")
            msg = email.message_from_bytes(raw_data[0][1])
            asunto = _decode_str(msg.get("Subject", "(sin asunto)"))
            from_  = _decode_str(msg.get("From", ""))
            body, attachments = _get_email_content(msg)
            if not body.strip() and not attachments:
                print(f"⏭️  Sin contenido: {asunto[:50]}")
                continue
            print(f"🔍 Analizando: {asunto[:60]}")
            datos = _extract_with_ai(asunto, body, attachments)
            if datos:
                # Forzar el mes al procesado si Gemini no lo detectó
                if not datos.get("mes") or str(datos.get("mes")).upper() in ("", "N/A"):
                    datos["mes"] = mes_nombre_cap
                save_to_sheets(datos, from_)
                facturas_encontradas += 1
                print(f"✅ Factura guardada: {datos.get('numero_factura','?')} – {datos.get('proveedor','?')}")
        except Exception as e:
            import traceback
            print(f"⚠️  Error procesando correo {eid}: {e}")
            print(traceback.format_exc())
            continue

    try:
        mail.logout()
    except Exception:
        pass

    print(f"🏁 Procesamiento {mes_nombre_cap} completado: {facturas_encontradas} facturas guardadas de {len(email_ids)} correos")


# ── EXTRACCIÓN DE TEXTO ───────────────────────────────────────────────────────
def _decode_str(s) -> str:
    if not s:
        return ""
    parts = decode_header(s)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            result.append(part.decode(enc or "utf-8", errors="replace"))
        else:
            result.append(str(part))
    return " ".join(result)


def _extract_pdf_text(data: bytes) -> str:
    """Extrae texto de PDF. Usa pdfplumber (tablas) + PyPDF2 fallback."""
    text = ""
    # 1) pdfplumber — mejor para tablas y layouts complejos
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            parts = []
            for page in pdf.pages:
                pg_text = page.extract_text() or ""
                # Extraer tablas como texto estructurado
                tables = page.extract_tables() or []
                for table in tables:
                    for row in table:
                        if row:
                            line = " | ".join(str(c).strip() for c in row if c is not None and str(c).strip())
                            if line:
                                parts.append(line)
                if pg_text.strip():
                    parts.append(pg_text)
            text = "\n".join(parts).strip()
    except Exception:
        pass
    # 2) PyPDF2 como fallback
    if not text:
        try:
            import PyPDF2
            reader = PyPDF2.PdfReader(io.BytesIO(data))
            text = "\n".join(p.extract_text() or "" for p in reader.pages)
        except Exception as e:
            text = f"[Error PDF: {e}]"
    return text


def _extract_docx_text(data: bytes) -> str:
    try:
        from docx import Document
        doc = Document(io.BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as e:
        return f"[Error DOCX: {e}]"


def _extract_xlsx_text(data: bytes) -> str:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        lines = []
        for ws in wb.worksheets:
            lines.append(f"=== Hoja: {ws.title} ===")
            for row in ws.iter_rows(values_only=True):
                row_vals = [str(c).strip() for c in row if c is not None and str(c).strip()]
                if row_vals:
                    lines.append(" | ".join(row_vals))
        return "\n".join(lines[:300])
    except Exception as e:
        return f"[Error XLSX: {e}]"


def _extract_xml_text(data: bytes) -> str:
    """Extrae campos clave de XML DIAN UBL 2.1 + fallback genérico."""
    try:
        import xml.etree.ElementTree as ET
        raw_str = data.decode("utf-8", errors="replace")
        root = ET.fromstring(raw_str)

        def _t(elem, *tags):
            """Busca un tag en múltiples namespaces y retorna su texto."""
            for tag in tags:
                for ns_prefix in ("", "{urn:oasis:names:specification:ubl:schema:xsd:Invoice-2}",
                                  "{urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2}",
                                  "{urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2}"):
                    e = root.find(f".//{ns_prefix}{tag}")
                    if e is not None and e.text and e.text.strip():
                        return e.text.strip()
            return None

        # Campos DIAN UBL 2.1 —                                 también cubren CFDI MX
        DIAN_FIELDS = [
            ("Número Factura",       ["ID", "InvoiceID", "Folio"]),
            ("Fecha Factura",        ["IssueDate", "FechaEmision"]),
            ("Proveedor NIT",        ["CompanyID", "TaxSchemeID"]),
            ("Proveedor Nombre",     ["RegistrationName", "Name", "PartyName"]),
            ("Subtotal",             ["LineExtensionAmount", "SubTotal"]),
            ("IVA",                  ["TaxAmount"]),
            ("Valor Total",          ["TaxInclusiveAmount", "PayableAmount", "Total"]),
            ("Moneda",               ["DocumentCurrencyCode", "Moneda"]),
            ("Tipo Factura",         ["InvoiceTypeCode"]),
        ]

        targeted = []
        for label, tags in DIAN_FIELDS:
            val = _t(root, *tags)
            if val:
                targeted.append(f"{label}: {val}")

        # Fallback genérico: todos los nodos con texto + atributos relevantes
        generic = []
        for elem in root.iter():
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if elem.text and elem.text.strip():
                generic.append(f"{tag}: {elem.text.strip()}")
            for ak, av in elem.attrib.items():
                ak = ak.split("}")[-1] if "}" in ak else ak
                generic.append(f"{tag}[{ak}]: {av}")

        combined = "\n".join(targeted) + "\n\n--- XML completo ---\n" + "\n".join(generic[:300])
        return combined
    except Exception as e:
        return f"[Error XML: {e}]"


def _extract_zip_text(data: bytes) -> str:
    """Extrae texto de archivos dentro de un ZIP (PDF, XLSX, XML, DOCX)."""
    try:
        import zipfile
        results = []
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            for name in zf.namelist():
                lower = name.lower()
                try:
                    file_data = zf.read(name)
                except Exception:
                    continue
                if lower.endswith(".pdf"):
                    results.append(f"[ZIP/{name}]\n" + _extract_pdf_text(file_data))
                elif lower.endswith(".xlsx") or lower.endswith(".xls"):
                    results.append(f"[ZIP/{name}]\n" + _extract_xlsx_text(file_data))
                elif lower.endswith(".docx"):
                    results.append(f"[ZIP/{name}]\n" + _extract_docx_text(file_data))
                elif lower.endswith(".xml"):
                    results.append(f"[ZIP/{name}]\n" + _extract_xml_text(file_data))
                elif lower.endswith(".txt"):
                    results.append(f"[ZIP/{name}]\n" + file_data.decode("utf-8", errors="replace")[:2000])
        return "\n\n".join(results) if results else "[ZIP vacío o sin archivos reconocibles]"
    except Exception as e:
        return f"[Error ZIP: {e}]"


def _extract_links_from_body(body: str) -> list:
    """Extrae URLs de un cuerpo de correo e intenta descargar documentos enlazados."""
    import urllib.request
    url_pattern = re.compile(r'https?://[^\s<>"]+', re.IGNORECASE)
    urls = url_pattern.findall(body)
    downloaded = []
    # Solo intentar URLs que parezcan documentos (PDF, XML, XLSX, ZIP)
    doc_exts = (".pdf", ".xml", ".xlsx", ".xls", ".zip", ".docx")
    seen = set()
    for url in urls:
        url_clean = url.rstrip(".,);'\"")
        if any(url_clean.lower().endswith(ext) for ext in doc_exts) and url_clean not in seen:
            seen.add(url_clean)
            try:
                req = urllib.request.Request(url_clean, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=8) as resp:
                    raw = resp.read()
                ext = url_clean.split("?")[0].lower()
                if ext.endswith(".pdf"):
                    downloaded.append(f"[URL: {url_clean}]\n" + _extract_pdf_text(raw))
                elif ext.endswith(".xml"):
                    downloaded.append(f"[URL: {url_clean}]\n" + _extract_xml_text(raw))
                elif ext.endswith(".xlsx") or ext.endswith(".xls"):
                    downloaded.append(f"[URL: {url_clean}]\n" + _extract_xlsx_text(raw))
                elif ext.endswith(".zip"):
                    downloaded.append(f"[URL: {url_clean}]\n" + _extract_zip_text(raw))
                elif ext.endswith(".docx"):
                    downloaded.append(f"[URL: {url_clean}]\n" + _extract_docx_text(raw))
                print(f"🔗 Descargado adjunto desde URL: {url_clean}")
            except Exception as e:
                print(f"⚠️  No se pudo descargar {url_clean}: {e}")
    return downloaded


def _get_email_content(msg) -> tuple[str, list]:
    """Retorna (cuerpo_texto, lista_de_textos_adjuntos)."""
    body = ""
    attachments = []

    for part in msg.walk():
        ct = part.get_content_type()
        cd = str(part.get("Content-Disposition", ""))

        if "attachment" in cd or ct in (
            "application/pdf",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "application/zip", "application/x-zip-compressed",
            "application/xml", "text/xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/pkcs7-mime",  # .p7m — XML DIAN firmado
            "image/png", "image/jpeg", "image/jpg", "image/gif", "image/webp",
        ):
            filename = part.get_filename() or ""
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            fname_lower = filename.lower()
            if fname_lower.endswith(".pdf") or ct == "application/pdf":
                attachments.append(f"[PDF: {filename}]\n" + _extract_pdf_text(payload))
                print(f"📎 PDF adjunto procesado: {filename}")
            elif fname_lower.endswith(".docx"):
                attachments.append(f"[DOCX: {filename}]\n" + _extract_docx_text(payload))
                print(f"📎 DOCX adjunto procesado: {filename}")
            elif fname_lower.endswith(".xlsx") or fname_lower.endswith(".xls"):
                attachments.append(f"[XLSX: {filename}]\n" + _extract_xlsx_text(payload))
                print(f"📎 XLSX adjunto procesado: {filename}")
            elif fname_lower.endswith(".xml") or fname_lower.endswith(".p7m") or ct in ("application/xml", "text/xml", "application/pkcs7-mime"):
                # .p7m puede ser XML firmado — intentar desenvuelto
                xml_bytes = payload
                if fname_lower.endswith(".p7m"):
                    # Intentar extraer XML del wrapper PKCS#7 buscando la secuencia XML
                    idx = payload.find(b"<?xml")
                    if idx != -1:
                        xml_bytes = payload[idx:]
                attachments.append(f"[XML: {filename}]\n" + _extract_xml_text(xml_bytes))
                print(f"📎 XML adjunto procesado: {filename}")
            elif fname_lower.endswith(".zip") or ct in ("application/zip", "application/x-zip-compressed"):
                attachments.append(f"[ZIP: {filename}]\n" + _extract_zip_text(payload))
                print(f"📎 ZIP adjunto procesado: {filename}")
            elif ct.startswith("image/") or fname_lower.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp")):
                # Guardar imagen en base64 para enviar a Claude Vision
                img_b64 = base64.b64encode(payload).decode()
                media_type = ct if ct.startswith("image/") else "image/png"
                attachments.append(("__IMAGE__", media_type, img_b64, filename))
                print(f"📎 Imagen adjunta registrada: {filename}")

        elif ct == "text/plain" and "attachment" not in cd:
            try:
                body += part.get_payload(decode=True).decode(
                    part.get_content_charset() or "utf-8", errors="replace"
                )
            except Exception:
                pass
        elif ct == "text/html" and "attachment" not in cd:
            try:
                from bs4 import BeautifulSoup
                raw = part.get_payload(decode=True).decode(
                    part.get_content_charset() or "utf-8", errors="replace"
                )
                html_text = BeautifulSoup(raw, "html.parser").get_text(separator="\n")
                # Siempre agregar el HTML aunque ya haya texto plano:
                # el HTML suele tener tablas con totales, más datos
                if html_text.strip() and html_text.strip() not in body:
                    body += "\n" + html_text
            except Exception:
                pass

    # Intentar descargar documentos enlazados en el cuerpo del correo
    if body:
        linked_docs = _extract_links_from_body(body)
        attachments.extend(linked_docs)

    return body, attachments


# ── INTELIGENCIA ARTIFICIAL (Claude) ─────────────────────────────────────────
JSON_SCHEMA = """
{
  "numero_factura": "(string — número/código de la factura)",
  "proveedor": "(string — razón social del emisor/vendedor)",
  "fecha_factura": "DD/MM/YYYY",
  "id_tipo": "NIT | CC | CE | RUT | PASAPORTE",
  "numero_id": "(NIT o cédula del proveedor, sin dígito verificador separado)",
  "subtotal": 0.0,
  "descuento": 0.0,
  "iva": 0.0,
  "rete_iva": 0.0,
  "rete_ica": 0.0,
  "impto_consumo": 0.0,
  "propina": 0.0,
  "otros_impuestos": 0.0,
  "retencion_fuente": 0.0,
  "administracion": 0.0,
  "utilidad": 0.0,
  "imprevistos": 0.0,
  "valor_total": 0.0,
  "clasificacion": "Servicios | Productos | Mixto",
  "estado": "PENDIENTE | PAGADA | VENCIDA",
  "valor_pagado": 0.0,
  "valor_por_pagar": 0.0,
  "fecha_pago": "DD/MM/YYYY o N/A",
  "cliente": "(string — razón social del receptor/comprador)",
  "cotizacion": "(número de cotización si aparece, si no N/A)",
  "observaciones": "(notas relevantes, descripción del servicio/producto)"
}"""


def _smart_truncate(text: str, max_chars: int) -> str:
    """Toma el inicio y el final del texto para no perder datos clave al final del doc."""
    if len(text) <= max_chars:
        return text
    half = max_chars // 2
    return text[:half] + "\n...[CONTENIDO OMITIDO]...\n" + text[-half:]


def _extract_with_ai(subject: str, body: str, attachments: list) -> Optional[dict]:
    """Usa Claude (Anthropic) para extraer datos de la factura. Retorna dict o None."""
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

        # ── Construir mensaje multimodal (texto + imágenes si las hay) ─────
        # Separar imágenes de textos
        text_parts  = [f"ASUNTO DEL CORREO: {subject}"]
        image_parts = []  # lista de (media_type, b64, filename)

        body_clean = _smart_truncate(body.strip(), 8000)
        if body_clean:
            text_parts.append(f"CUERPO DEL CORREO:\n{body_clean}")

        for i, att in enumerate(attachments, 1):
            if isinstance(att, tuple) and att[0] == "__IMAGE__":
                _, media_type, b64, fname = att
                image_parts.append((media_type, b64, fname))
            else:
                att_text = _smart_truncate(str(att), 7000)
                text_parts.append(f"ADJUNTO {i}:\n{att_text}")

        full_text = "\n\n".join(text_parts)

        prompt = f"""Eres un experto en contabilidad colombiana y facturas electrónicas DIAN.
Analiza TODO el contenido que se te da (cuerpo del correo, adjuntos PDF, XML DIAN, XLSX, imágenes) y extrae los datos de la factura.

REGLAS ESTRICTAS:
1. Si NO existe ninguna factura en el contenido, responde EXACTAMENTE: NO_ES_FACTURA
2. Si HAY factura, extrae TODOS los campos posibles aunque sean parciales.
3. Números: usa formato decimal con punto (ej: 1234567.50). Sin puntos de miles. Si el campo no aparece usa 0.
4. Texto: valor exacto tal como aparece. Si no aparece usa "N/A".
5. numero_factura: busca 'No. Factura', 'Número', 'FACT-', 'FE-', 'FV', 'FES', 'Invoice No', campo ID en XML DIAN.
6. proveedor: quien EMITE (vende). cliente: quien RECIBE (compra/paga).
7. numero_id: NIT sin dígito verificador (ej: '900123456' no '900123456-1').
8. iva: suma de todos los TaxAmount con TaxCode=01 o nombre 'IVA'.
9. rete_iva: retención sobre el IVA (normalmente 15% del IVA).
10. rete_ica: Impuesto de Industria y Comercio retenido.
11. retencion_fuente: retención en la fuente (busca 'RteFte', 'RetFte', 'Retención fuente').
12. valor_total: valor final a pagar después de impuestos y retenciones.
13. estado: PAGADA si aparece 'pagado/cancelado', VENCIDA si venció sin pagar, si no PENDIENTE.
14. clasificacion: Servicios/Productos/Mixto según descripción de ítems de la factura.
15. En XML DIAN UBL 2.1: LineExtensionAmount=subtotal, TaxInclusiveAmount o PayableAmount=valor_total.

Respóndeme ÚNICAMENTE con el JSON, sin texto antes ni después, sin bloques ```:
{JSON_SCHEMA}

CONTENIDO COMPLETO:
{full_text}
"""

        # ── Construir content (texto + imágenes opcionales) ───────────────
        content: list = []
        if image_parts:
            # Claude admite hasta 20 imágenes por mensaje
            for media_type, b64, fname in image_parts[:5]:
                content.append({"type": "text", "text": f"[Imagen adjunta: {fname}]"})
                content.append({"type": "image", "source": {
                    "type": "base64",
                    "media_type": media_type,
                    "data": b64
                }})
        content.append({"type": "text", "text": prompt})

        resp = client.messages.create(
            model="claude-haiku-4-5",
            max_tokens=2048,
            messages=[{"role": "user", "content": content}]
        )
        text = resp.content[0].text.strip()

        if "NO_ES_FACTURA" in text:
            return None

        # Extraer JSON: intentar directo, luego buscar bloque JSON con llaves balanceadas
        def _extract_json(raw: str) -> Optional[dict]:
            # 1) Intentar parsear todo el texto directamente
            try:
                return json.loads(raw)
            except Exception:
                pass
            # 2) Buscar bloque ```json ... ``` o ``` ... ```
            block = re.search(r'```(?:json)?\s*([\s\S]*?)```', raw)
            if block:
                try:
                    return json.loads(block.group(1).strip())
                except Exception:
                    pass
            # 3) Encontrar el bloque JSON con llaves balanceadas
            start = raw.find('{')
            if start == -1:
                return None
            depth = 0
            for i, ch in enumerate(raw[start:], start):
                if ch == '{':
                    depth += 1
                elif ch == '}':
                    depth -= 1
                    if depth == 0:
                        try:
                            return json.loads(raw[start:i+1])
                        except Exception:
                            return None
            return None

        data = _extract_json(text)
        if data:
            return data
        print(f"⚠️ Claude no devolvió JSON válido. Respuesta: {text[:200]}")
        return None

    except Exception as e:
        print(f"⚠️ Error IA: {e}")
        return None


def _is_duplicate_invoice(numero_factura: str, mes_nombre: str = None) -> tuple:
    """Busca la factura en TODAS las hojas mensuales.
    Retorna (encontrado: bool, fila_num: int, worksheet_object).
    Si numero_factura es N/A/vacío devuelve (False, -1, None)."""
    if not numero_factura or str(numero_factura).strip() in ("", "N/A"):
        return False, -1, None
    num_clean = str(numero_factura).strip()
    hojas_revisadas = set()
    try:
        # 1) Revisar todas las hojas existentes
        all_ws = get_all_monthly_worksheets()
        for ws in all_ws:
            hojas_revisadas.add(ws.title)
            try:
                col_values = _api_call_with_retry(ws.col_values, 3)  # col 3 = "Número Factura"
                for i, val in enumerate(col_values[1:], start=2):    # fila 1=cabecera, fila 2=primer dato
                    if str(val).strip() == num_clean:
                        return True, i, ws
            except Exception:
                continue
        # 2) Si se especificó un mes y su hoja no estaba en get_all_monthly_worksheets, revisarla igual
        if mes_nombre and mes_nombre not in hojas_revisadas:
            try:
                ws = get_sheet(mes_nombre)
                col_values = _api_call_with_retry(ws.col_values, 3)
                for i, val in enumerate(col_values[1:], start=2):
                    if str(val).strip() == num_clean:
                        return True, i, ws
            except Exception:
                pass
        return False, -1, None
    except Exception:
        return False, -1, None


def _build_invoice_row(invoice_data: dict, mes_nombre: str) -> list:
    """Construye la lista de valores en el mismo orden que COLUMNS."""
    valor_total  = float(invoice_data.get("valor_total")  or 0)
    valor_pagado = float(invoice_data.get("valor_pagado") or 0)
    return [
        mes_nombre,
        str(invoice_data.get("fecha_factura") or "N/A"),
        str(invoice_data.get("numero_factura") or "N/A"),
        str(invoice_data.get("proveedor")      or "N/A"),
        str(invoice_data.get("id_tipo")        or "N/A"),
        str(invoice_data.get("numero_id")      or "N/A"),
        float(invoice_data.get("subtotal")          or 0),
        float(invoice_data.get("descuento")         or 0),
        float(invoice_data.get("iva")               or 0),
        float(invoice_data.get("rete_iva")          or 0),
        float(invoice_data.get("rete_ica")          or 0),
        float(invoice_data.get("impto_consumo")     or 0),
        float(invoice_data.get("propina")           or 0),
        float(invoice_data.get("otros_impuestos")   or 0),
        float(invoice_data.get("retencion_fuente")  or 0),
        float(invoice_data.get("administracion")    or 0),
        float(invoice_data.get("utilidad")          or 0),
        float(invoice_data.get("imprevistos")       or 0),
        valor_total,
        str(invoice_data.get("clasificacion") or "N/A"),
        str(invoice_data.get("estado")        or "PENDIENTE"),
        valor_pagado,
        valor_total - valor_pagado,
        str(invoice_data.get("fecha_pago")    or "N/A"),
        str(invoice_data.get("cliente")       or "N/A"),
        str(invoice_data.get("cotizacion")    or "N/A"),
        str(invoice_data.get("observaciones") or ""),
    ]


# ── GUARDAR EN GOOGLE SHEETS ──────────────────────────────────────────────────
def save_to_sheets(invoice_data: dict, email_from: str):
    """INSERT si la factura no existe, UPDATE si ya existe con datos incompletos."""
    try:
        numero    = str(invoice_data.get("numero_factura") or "N/A")
        proveedor = str(invoice_data.get("proveedor")      or "N/A")

        # ── 1. Determinar el mes ANTES de cualquier otra cosa ──────────────
        fecha_factura_str = str(invoice_data.get("fecha_factura") or "N/A")
        meses_list = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                      "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        try:
            if "/" in fecha_factura_str:
                mes_num  = int(fecha_factura_str.split("/")[1])
                mes_nombre = meses_list[mes_num - 1] if 1 <= mes_num <= 12 else meses_list[datetime.now().month - 1]
            else:
                mes_nombre = meses_list[datetime.now().month - 1]
        except Exception:
            mes_nombre = meses_list[datetime.now().month - 1]

        # ── 2. Buscar duplicado en TODAS las hojas ─────────────────────────
        found, dup_row, dup_ws = _is_duplicate_invoice(numero, mes_nombre)

        if found and dup_ws is not None:
            # UPDATE: si el dato nuevo tiene valor_total > 0 y el existente es 0/vacío, actualizar
            try:
                existing_row_vals = _api_call_with_retry(dup_ws.row_values, dup_row)
                # col 19 = Valor Total (índice 18 en la lista), col 4 = Proveedor (índice 3)
                existing_total = float(existing_row_vals[18]) if len(existing_row_vals) > 18 and existing_row_vals[18] not in ("", "0", "N/A") else 0
                new_total      = float(invoice_data.get("valor_total") or 0)

                if new_total > 0 and existing_total == 0:
                    new_row = _build_invoice_row(invoice_data, mes_nombre)
                    # Actualizar cada celda de la fila (A=1 … AA=27)
                    for col_idx, val in enumerate(new_row, start=1):
                        _api_call_with_retry(dup_ws.update_cell, dup_row, col_idx, val)
                    print(f"🔄 Factura ACTUALIZADA (datos completos): {numero} — {proveedor}")
                else:
                    print(f"⏭️  Factura ya existe y está completa, omitida: {numero} — {proveedor}")
            except Exception as e_upd:
                print(f"⚠️  No se pudo actualizar fila duplicada {numero}: {e_upd}")
            return

        # ── 3. INSERT ──────────────────────────────────────────────────────
        row = _build_invoice_row(invoice_data, mes_nombre)
        ws  = get_sheet(mes_nombre)
        _api_call_with_retry(ws.append_row, row, value_input_option="USER_ENTERED")
        print(f"✅ Factura guardada en hoja '{mes_nombre}': {numero} — {proveedor}")

    except Exception as e:
        print(f"❌ Error guardando en Sheets: {e}")
        import traceback
        print(traceback.format_exc())


def export_to_excel() -> str:
    """Descarga todos los registros de todas las hojas mensuales y genera un Excel con múltiples pestañas."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    try:
        # Usar spreadsheet cacheado — NO repite open_by_key
        spreadsheet = _get_spreadsheet()
        worksheets = _api_call_with_retry(spreadsheet.worksheets)
        
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Eliminar la hoja por defecto

        # Procesar solo las hojas que son meses (usa la constante global MESES)
        for ws_gsheet in worksheets:
            if ws_gsheet.title in MESES:
                rows = ws_gsheet.get_all_values()
                if not rows:  # Saltar hojas vacías
                    continue
                
                sheet = wb.create_sheet(title=ws_gsheet.title)
                for i, row in enumerate(rows, 1):
                    sheet.append(row)
                    if i == 1:  # Formatear encabezados
                        for cell in sheet[1]:
                            cell.font = Font(bold=True, color="FFFFFF")
                            cell.fill = PatternFill("solid", fgColor="1F4E79")
                            cell.alignment = Alignment(horizontal="center")
        
        # Si no hay hojas mensuales, crear una por defecto
        if not wb.worksheets:
            sheet = wb.create_sheet(title="Sin Datos")
            sheet.append(["No hay datos disponibles"])
        
        wb.save(EXCEL_PATH)
        print(f"✅ Excel exportado con {len(wb.worksheets)} hojas: {EXCEL_PATH}")
        return EXCEL_PATH
    except Exception as e:
        print(f"❌ Error exportando Excel: {e}")
        raise


# ── PROCESO PRINCIPAL ──────────────────────────────────────────────────────────
def process_emails():
    """Función principal del agente. Conecta, lee, extrae y guarda."""
    if not EMAIL_USER or not EMAIL_PASS:
        print("❌ Credenciales de correo no configuradas en .env")
        return
    if not ANTHROPIC_API_KEY:
        print("❌ ANTHROPIC_API_KEY no configurada en .env")
        return

    setup_sheets()
    processed_ids = _load_processed_ids()

    print(f"📧 Conectando a Gmail ({EMAIL_USER})...")
    try:
        mail = connect_imap()
    except Exception as e:
        print(f"❌ Error de conexión IMAP: {e}")
        return

    email_ids = search_invoice_emails(mail)
    nuevos = [eid for eid in email_ids if eid not in processed_ids]
    print(f"🔍 {len(nuevos)} correos nuevos de {len(email_ids)} encontrados")

    facturas_encontradas = 0

    for eid in nuevos:
        try:
            _, data = mail.fetch(eid, "(RFC822)")
            msg = email.message_from_bytes(data[0][1])
            subject = _decode_str(msg.get("Subject", "(sin asunto)"))
            from_   = _decode_str(msg.get("From", ""))

            print(f"📨 Analizando: {subject[:60]}")

            body, attachments = _get_email_content(msg)
            invoice_data = _extract_with_ai(subject, body, attachments)

            if invoice_data:
                save_to_sheets(invoice_data, from_)
                facturas_encontradas += 1

            processed_ids.add(eid)
            _save_processed_ids(processed_ids)


        except Exception as e:
            print(f"⚠️ Error procesando correo {eid}: {e}")
            continue

    try:
        mail.logout()
    except Exception:
        pass

    print(f"✅ Proceso completado — {facturas_encontradas} factura(s) extraída(s) de {len(nuevos)} correos")


# ── ENTRY POINT STANDALONE ─────────────────────────────────────────────────────
if __name__ == "__main__":
    process_emails()
