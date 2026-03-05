"""
 Módulo agente de extracción de facturas.
 - Conexión IMAP a Gmail
 - Extracción de datos con Claude AI (Anthropic)
 - Procesamiento de PDF, DOCX, XLSX adjuntos
 - Guardado en Google Sheets
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
+import unicodedata
 from datetime import datetime
 from typing import Optional
 
 # ── MODELO ────────────────────────────────────────────────────────────────────
 # claude-sonnet-4-6: modelo ideal para extracción de facturas DIAN
 # → Mayor precisión en campos complejos (rete_iva, rete_ica, XML UBL 2.1)
 # → Mejor manejo de PDFs con tablas, imágenes y adjuntos múltiples
 # → Mismo costo que Sonnet 4.5, con mejoras en razonamiento estructurado
 CLAUDE_MODEL = "claude-sonnet-4-6"
 
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
@@ -192,124 +193,132 @@ def search_invoice_emails(mail: imaplib.IMAP4_SSL) -> list:
             if status != 'OK':
                 continue
             _, data = mail.search(None, "ALL")
             ids = [uid.decode() for uid in data[0].split() if uid]
             all_ids_set.update(ids)
             if ids:
                 break  # encontrado el folder correcto, detener búsqueda
         except Exception:
             continue
     # Restaurar INBOX
     try:
         mail.select('INBOX')
     except Exception:
         pass
     recientes = sorted(all_ids_set, key=lambda x: int(x) if x.isdigit() else 0)[-200:]
     print(f"📬 {len(recientes)} correos para analizar")
     return recientes
 
 
 # Mapa de meses español → número
 _MES_NUM = {
     "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
     "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
     "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
 }
+
+
+def _normalize_mes_nombre(mes_nombre: str) -> str:
+    """Normaliza variantes de mes (acentos, espacios, setiembre→septiembre)."""
+    txt = unicodedata.normalize("NFKD", str(mes_nombre or ""))
+    txt = "".join(c for c in txt if not unicodedata.combining(c)).lower().strip()
+    return "septiembre" if txt == "setiembre" else txt
 # Nombres en inglés para IMAP (formato requerido: "01-Jan-2026")
 _MES_EN = {
     1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
     7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
 }
 
 
 def search_emails_by_month(mail: imaplib.IMAP4_SSL, mes_nombre: str, year: int = None) -> list:
     """Busca en Gmail TODOS los correos del mes usando IMAP SINCE/BEFORE, revisando All Mail."""
     import calendar
-    mes_nombre = mes_nombre.lower().strip()
+    mes_nombre = _normalize_mes_nombre(mes_nombre)
     mes_num = _MES_NUM.get(mes_nombre)
     if not mes_num:
         print(f"❌ Mes no reconocido: {mes_nombre}")
         return []
 
     if year is None:
         year = datetime.now().year
 
     ultimo_dia = calendar.monthrange(year, mes_num)[1]
     fecha_inicio = f"01-{_MES_EN[mes_num]}-{year}"
     mes_siguiente = mes_num + 1 if mes_num < 12 else 1
     year_siguiente = year if mes_num < 12 else year + 1
     fecha_fin = f"01-{_MES_EN[mes_siguiente]}-{year_siguiente}"
     criterio = f'(SINCE "{fecha_inicio}" BEFORE "{fecha_fin}")'
 
     all_ids_set = set()
     # Buscar en todos los folders relevantes de Gmail
     folders_to_try = ['INBOX', '"[Gmail]/All Mail"', '[Gmail]/All Mail',
                       '"[Gmail]/Starred"', 'All Mail']
     for folder in folders_to_try:
         try:
             status, _ = mail.select(folder, readonly=True)
             if status != 'OK':
                 continue
             _, data = mail.search(None, criterio)
             ids = [uid.decode() for uid in data[0].split() if uid]
             all_ids_set.update(ids)
         except Exception:
             continue
     # Restaurar INBOX
     try:
         mail.select('INBOX')
     except Exception:
         pass
     ids = sorted(all_ids_set, key=lambda x: int(x) if x.isdigit() else 0)
     print(f"📬 {len(ids)} correos encontrados en {mes_nombre.capitalize()} {year}")
     return ids
 
 
 def process_emails_for_month(mes_nombre: str, year: int = None):
     """Procesa TODOS los correos de un mes específico, ignorando el cache de procesados."""
     if not EMAIL_USER or not EMAIL_PASS:
         print("❌ Credenciales de correo no configuradas en .env")
         return
     if not ANTHROPIC_API_KEY:
         print("❌ ANTHROPIC_API_KEY no configurada en .env")
         return
 
-    mes_nombre_cap = mes_nombre.strip().capitalize()
+    mes_normalizado = _normalize_mes_nombre(mes_nombre)
+    mes_nombre_cap = mes_normalizado.capitalize()
     if year is None:
         year = datetime.now().year
 
     print(f"🗓️  Iniciando procesamiento del mes: {mes_nombre_cap} {year}")
     print(f"🤖 Modelo de extracción: {CLAUDE_MODEL}")
     setup_sheets()
 
     try:
         mail = connect_imap()
     except Exception as e:
         print(f"❌ Error de conexión IMAP: {e}")
         return
 
-    email_ids = search_emails_by_month(mail, mes_nombre, year)
+    email_ids = search_emails_by_month(mail, mes_normalizado, year)
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
                 # ── [v5.2] Validación post-extracción ─────────────────────
                 valida, motivo = _validate_invoice_data(datos, asunto)
                 if not valida:
                     print(f"🚫 Rechazado por validación: {motivo} | Asunto: {asunto[:60]}")
                     continue
@@ -815,51 +824,54 @@ def _validate_invoice_data(datos: dict, subject: str = "") -> tuple[bool, str]:
         r"^PSE",           # Links PSE
         r"^WOMPI",         # Wompi
         r"^PAGO",          # Links de pago genéricos
         r"^SOL",           # Solicitudes
         r"^RESET",         # Reset de clave
     ]
     num_upper = numero.upper()
     for patron in patrones_invalidos:
         if re.match(patron, num_upper):
             return False, f"Número de factura con patrón no DIAN: '{numero}'"
 
     # 3. El número debe tener al menos 3 caracteres útiles
     if len(numero) < 3:
         return False, f"Número de factura demasiado corto: '{numero}'"
 
     # 4. Proveedor no puede ser genérico o vacío
     if not proveedor or proveedor.upper() in ("N/A", "NA", "NONE", "", "NULL", "PROVEEDOR"):
         return False, f"Proveedor inválido: '{proveedor}'"
 
     # 5. Valor total debe ser positivo (facturas de $0 no tienen sentido)
     if valor_total <= 0:
         # Advertir pero NO rechazar — algunas facturas DIAN tienen total en 0 por descuento total
         print(f"⚠️  Factura {numero} tiene valor_total=0. Verificar manualmente.")
 
     return True, ""
-                           proveedor: str = None, fecha: str = None) -> tuple:
+
+
+def _is_duplicate_invoice(numero_factura: str, mes_nombre: str,
+                          proveedor: str = None, fecha: str = None) -> tuple:
     """Busca la factura en TODAS las hojas mensuales.
     Retorna (encontrado: bool, fila_num: int, worksheet_object).
     Cuando numero_factura es N/A usa clave compuesta proveedor+fecha como fallback."""
     num_clean = str(numero_factura).strip() if numero_factura else ""
     use_num   = bool(num_clean and num_clean not in ("", "N/A"))
     # Clave secundaria para facturas sin número
     prov_clean  = str(proveedor or "").strip().lower()
     fecha_clean = str(fecha or "").strip()
     use_composite = (not use_num) and bool(prov_clean and prov_clean != "n/a")
     if not use_num and not use_composite:
         return False, -1, None
 
     hojas_revisadas = set()
     try:
         all_ws = get_all_monthly_worksheets()
         for ws in all_ws:
             hojas_revisadas.add(ws.title)
             try:
                 if use_num:
                     col_values = _api_call_with_retry(ws.col_values, 3)  # col 3 = Número Factura
                     for i, val in enumerate(col_values[1:], start=2):
                         if str(val).strip() == num_clean:
                             return True, i, ws
                 elif use_composite:
                     # Leer col 2 (Fecha) y col 4 (Proveedor) para clave compuesta
 
EOF
)
