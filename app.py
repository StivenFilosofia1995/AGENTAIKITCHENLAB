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
            msg_data = fetch_email(mail, eid)
            if not msg_data:
                continue
            msg = email.message_from_bytes(msg_data)
            asunto = _decode_str(msg.get("Subject", ""))
            texto = extract_email_text(msg)
            if not texto.strip():
                print(f"⏭️  Sin texto: {asunto[:50]}")
                continue
            print(f"🔍 Analizando: {asunto[:60]}")
            datos = extract_invoice_data(texto, asunto)
            if datos:
                # Forzar el mes del parámetro si el campo Mes está vacío
                if not datos.get("mes") or datos.get("mes") == "N/A":
                    datos["mes"] = mes_nombre_cap
                save_to_sheets(datos)
                facturas_encontradas += 1
                print(f"✅ Factura guardada: {datos.get('numero_factura','?')} – {datos.get('proveedor','?')}")
        except Exception as e:
            print(f"⚠️  Error procesando correo {eid}: {e}")
            continue

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
    try:
        import PyPDF2
        reader = PyPDF2.PdfReader(io.BytesIO(data))
        return "\n".join(p.extract_text() or "" for p in reader.pages)
    except Exception as e:
        return f"[Error PDF: {e}]"


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
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True)
        lines = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                line = " | ".join(str(c) for c in row if c is not None)
                if line.strip():
                    lines.append(line)
        return "\n".join(lines[:100])
    except Exception as e:
        return f"[Error XLSX: {e}]"


def _get_email_content(msg) -> tuple[str, list]:
    """Retorna (cuerpo_texto, lista_de_textos_adjuntos)."""
    body = ""
    attachments = []

    for part in msg.walk():
        ct = part.get_content_type()
        cd = str(part.get("Content-Disposition", ""))

        if "attachment" in cd or ct in ("application/pdf", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"):
            filename = part.get_filename() or ""
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            if filename.lower().endswith(".pdf") or ct == "application/pdf":
                attachments.append(_extract_pdf_text(payload))
            elif filename.lower().endswith(".docx"):
                attachments.append(_extract_docx_text(payload))
            elif filename.lower().endswith(".xlsx"):
                attachments.append(_extract_xlsx_text(payload))

        elif ct == "text/plain" and "attachment" not in cd:
            try:
                body += part.get_payload(decode=True).decode(
                    part.get_content_charset() or "utf-8", errors="replace"
                )
            except Exception:
                pass
        elif ct == "text/html" and not body and "attachment" not in cd:
            try:
                from bs4 import BeautifulSoup
                raw = part.get_payload(decode=True).decode(
                    part.get_content_charset() or "utf-8", errors="replace"
                )
                body += BeautifulSoup(raw, "html.parser").get_text(separator="\n")
            except Exception:
                pass

    return body, attachments


# ── INTELIGENCIA ARTIFICIAL (Claude) ─────────────────────────────────────────
def _extract_with_ai(subject: str, body: str, attachments: list) -> Optional[dict]:
    """Usa Claude (Anthropic) para extraer datos de la factura. Retorna dict o None."""
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

        content_parts = [f"Asunto: {subject}", f"Cuerpo:\n{body[:4000]}"]
        for i, att in enumerate(attachments, 1):
            content_parts.append(f"Adjunto {i}:\n{att[:4000]}")
        full_content = "\n\n".join(content_parts)

        prompt = f"""Analiza el siguiente correo y extrae los datos de factura si los hay.
Si NO es una factura, responde exactamente: NO_ES_FACTURA

Si SÍ es una factura, responde SOLO en formato JSON con TODOS estos campos (usa "N/A" o 0 si no encuentras el dato):
{{
  "numero_factura": "...",
  "proveedor": "...",
  "fecha_factura": "DD/MM/YYYY",
  "id_tipo": "NIT/CC/CE",
  "numero_id": "...",
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
  "clasificacion": "Servicios/Productos/Otro",
  "estado": "PENDIENTE/PAGADA",
  "valor_pagado": 0.0,
  "valor_por_pagar": 0.0,
  "fecha_pago": "DD/MM/YYYY o N/A",
  "cliente": "...",
  "observaciones": "..."
}}

CORREO:
{full_content}
"""
        resp = client.messages.create(
            model="claude-haiku-4-5",
            max_tokens=512,
            messages=[{"role": "user", "content": prompt}]
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


def _is_duplicate_invoice(numero_factura: str) -> bool:
    """Comprueba si la factura ya existe en Google Sheets."""
    if not numero_factura or numero_factura == "N/A":
        return False
    try:
        ws = get_sheet()
        col_values = ws.col_values(1)  # columna "N° Factura"
        for val in col_values[1:]:     # saltar encabezado
            if str(val).strip() == str(numero_factura).strip():
                return True
        return False
    except Exception:
        return False


# ── GUARDAR EN GOOGLE SHEETS ──────────────────────────────────────────────────
def save_to_sheets(invoice_data: dict, email_from: str):
    """Agrega una fila a Google Sheets con los datos de la factura en la hoja del mes correspondiente."""
    try:
        numero    = str(invoice_data.get("numero_factura") or "N/A")
        proveedor = str(invoice_data.get("proveedor") or "N/A")

        if _is_duplicate_invoice(numero):
            print(f"⏭️ Factura duplicada omitida: {numero} — {proveedor}")
            return

        # Determinar el mes de la factura
        fecha_factura_str = str(invoice_data.get("fecha_factura") or "N/A")
        meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                 "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        
        # Intentar extraer el mes de la fecha de factura (formato DD/MM/YYYY)
        try:
            if "/" in fecha_factura_str:
                mes_num = int(fecha_factura_str.split("/")[1])
                mes_nombre = meses[mes_num - 1] if 1 <= mes_num <= 12 else meses[datetime.now().month - 1]
            else:
                mes_nombre = meses[datetime.now().month - 1]
        except:
            mes_nombre = meses[datetime.now().month - 1]

        # Calcular valor por pagar
        valor_total = float(invoice_data.get("valor_total") or 0)
        valor_pagado = float(invoice_data.get("valor_pagado") or 0)
        valor_por_pagar = valor_total - valor_pagado

        row = [
            mes_nombre,
            fecha_factura_str,
            numero,
            proveedor,
            str(invoice_data.get("id_tipo") or "N/A"),
            str(invoice_data.get("numero_id") or "N/A"),
            float(invoice_data.get("subtotal") or 0),
            float(invoice_data.get("descuento") or 0),
            float(invoice_data.get("iva") or 0),
            float(invoice_data.get("rete_iva") or 0),
            float(invoice_data.get("rete_ica") or 0),
            float(invoice_data.get("impto_consumo") or 0),
            float(invoice_data.get("propina") or 0),
            float(invoice_data.get("otros_impuestos") or 0),
            float(invoice_data.get("retencion_fuente") or 0),
            float(invoice_data.get("administracion") or 0),
            float(invoice_data.get("utilidad") or 0),
            float(invoice_data.get("imprevistos") or 0),
            valor_total,
            str(invoice_data.get("clasificacion") or "N/A"),
            str(invoice_data.get("estado") or "PENDIENTE"),
            valor_pagado,
            valor_por_pagar,
            str(invoice_data.get("fecha_pago") or "N/A"),
            str(invoice_data.get("cliente") or "N/A"),
            str(invoice_data.get("cotizacion") or "N/A"),
            str(invoice_data.get("observaciones") or ""),
        ]
        
        ws = get_sheet(mes_nombre)
        ws.append_row(row, value_input_option="USER_ENTERED")
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
