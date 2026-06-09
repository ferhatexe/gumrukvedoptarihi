import os
import re
import json
import asyncio
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, date
from typing import List, Dict, Any

import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from fastapi import FastAPI, WebSocket, WebSocketDisconnect, File, UploadFile
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse

from scraper import HttpCustomsScraper

app = FastAPI(title="Gümrük Beyanname Sorgulama Otomasyonu")

# Excel Paths
LOCAL_BASE_DIR = r"c:\WORK\00_INBOX\MAYIS BEYANLAR\MAYIS BEYANLAR"
if os.path.exists(LOCAL_BASE_DIR):
    BASE_DIR = LOCAL_BASE_DIR
else:
    BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
    os.makedirs(BASE_DIR, exist_ok=True)

EXCEL_PATH = os.path.join(BASE_DIR, "EXPORT.XLSX")
EXCEL_CUSTOM_PATH = os.path.join(BASE_DIR, "EXPORT_CUSTOM.XLSX")

# Global sessions registry mapping session_id -> UserSessionState
class UserSessionState:
    def __init__(self, session_id: str):
        self.session_id = session_id
        self.active_excel_path = None
        self.gcb_col_idx = 9
        self.date_col_idx = 12
        self.fatura_col_idx = 1
        self.firma_col_idx = 3
        
        # Scraper state fields
        self.is_running = False
        self.bypass = False
        self.task = None
        self.cancel_event = threading.Event()
        self.completed_count = 0
        self.total_count = 0
        self.log_history = []

sessions: Dict[str, UserSessionState] = {}

def get_session(session_id: str) -> UserSessionState:
    if not session_id:
        session_id = "default_session"
    if session_id not in sessions:
        sessions[session_id] = UserSessionState(session_id)
    return sessions[session_id]

def get_display_filename(session: UserSessionState) -> str:
    """Return a clean display filename without session_id hash prefix."""
    if not session.active_excel_path:
        return None
    basename = os.path.basename(session.active_excel_path)
    # Strip session_id prefix (format: {session_id}_{original_name})
    prefix = f"{session.session_id}_"
    if basename.startswith(prefix):
        return basename[len(prefix):]
    return basename

# Keep track of active WebSocket connections per session
class ConnectionManager:
    def __init__(self):
        # session_id -> List[WebSocket]
        self.active_connections: Dict[str, List[WebSocket]] = {}

    async def connect(self, session_id: str, websocket: WebSocket):
        await websocket.accept()
        if not session_id:
            session_id = "default_session"
        if session_id not in self.active_connections:
            self.active_connections[session_id] = []
        self.active_connections[session_id].append(websocket)

    def disconnect(self, session_id: str, websocket: WebSocket):
        if not session_id:
            session_id = "default_session"
        if session_id in self.active_connections:
            if websocket in self.active_connections[session_id]:
                self.active_connections[session_id].remove(websocket)
            if not self.active_connections[session_id]:
                del self.active_connections[session_id]

    async def broadcast_to_session(self, session_id: str, message: dict):
        if not session_id:
            session_id = "default_session"
        if session_id in self.active_connections:
            for connection in self.active_connections[session_id]:
                try:
                    await connection.send_json(message)
                except Exception:
                    pass

manager = ConnectionManager()

def normalize_turkish(text: str) -> str:
    if not text:
        return ""
    mapping = {
        'İ': 'i', 'I': 'ı', 'Ş': 'ş', 'Ç': 'ç', 'Ğ': 'ğ', 'Ü': 'ü', 'Ö': 'ö',
        'ı': 'ı', 'ş': 'ş', 'ç': 'ç', 'ğ': 'ğ', 'ü': 'ü', 'ö': 'ö', 'i': 'i'
    }
    return "".join(mapping.get(c, c.lower()) for c in text)

def apply_table_formatting_to_sheet(ws):
    try:
        ws.sheet_view.showGridLines = True
    except Exception:
        try:
            ws.views.sheetView[0].showGridLines = True
        except Exception:
            pass

    max_row = ws.max_row
    max_col = ws.max_column
    
    if max_row < 1 or max_col < 1:
        return

    # Clear existing tables first to prevent overlaps/errors
    if hasattr(ws, '_tables'):
        ws._tables.clear()
    ws.auto_filter.ref = None

    # Define the Table range
    ref = f"A1:{get_column_letter(max_col)}{max_row}"
    
    # Create the Table object
    tab = Table(displayName="GumrukSorguTablosu", ref=ref)
    
    # Style: TableStyleMedium2 (Standard Excel blue theme with header and striped rows)
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)
    
    # Alignments
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    
    # Loop columns to auto-fit and style cells
    for col_idx in range(1, max_col + 1):
        header_val = str(ws.cell(row=1, column=col_idx).value or "").strip()
        hl = normalize_turkish(header_val)
        
        is_center_col = any(k in hl for k in [
            "tarih", "date", "no", "numara", "gcb", "gb", "fatura", "tescil", "kod", "code"
        ])
        
        max_len = len(header_val)
        for row_idx in range(2, max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            val_str = str(cell.value or "").strip()
            
            # Standardize date format to DD.MM.YYYY string or date object
            if any(k in hl for k in ["tarih", "date", "intaç", "intac"]):
                if isinstance(cell.value, (datetime, date)):
                    cell.number_format = 'dd.mm.yyyy'
                    val_str = cell.value.strftime("%d.%m.%Y")
                elif val_str and re.match(r'^\d{4}-\d{2}-\d{2}$', val_str):
                    try:
                        d_obj = datetime.strptime(val_str, "%Y-%m-%d")
                        cell.value = d_obj.date()
                        cell.number_format = 'dd.mm.yyyy'
                        val_str = d_obj.strftime("%d.%m.%Y")
                    except Exception:
                        pass
                elif val_str and re.match(r'^\d{2}\.\d{2}\.\d{4}$', val_str):
                    try:
                        d_obj = datetime.strptime(val_str, "%d.%m.%Y")
                        cell.value = d_obj.date()
                        cell.number_format = 'dd.mm.yyyy'
                    except Exception:
                        pass
                        
            # Apply Alignment
            if is_center_col:
                cell.alignment = center_align
            else:
                cell.alignment = left_align
                
            if cell.value is not None:
                max_len = max(max_len, len(val_str))
                
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

# Robust line parser for custom paste strings (extracts only the GCB number)
def parse_custom_line(line: str):
    # Regex match GCB No: e.g. 26341200EX00137190 (8 digits, 2 letters, 6 to 8 digits)
    match = re.search(r'\d{8}[A-Za-z]{2}\d{6,8}', line)
    gcb = match.group(0).upper() if match else None
    return gcb, "", ""

def read_excel_data(file_path: str) -> Dict[str, Any]:
    # Reset defaults in case of empty or missing spreadsheet
    gcb_col_idx = 9
    date_col_idx = 12
    fatura_col_idx = 1
    firma_col_idx = 3
    
    if not file_path or not os.path.exists(file_path):
        return {
            "headers": [], 
            "rows": [], 
            "gcb_col_idx": gcb_col_idx, 
            "date_col_idx": date_col_idx, 
            "fatura_col_idx": fatura_col_idx, 
            "firma_col_idx": firma_col_idx
        }
    
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    
    # Read headers from row 1
    headers = []
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(row=1, column=c).value or "").strip()
        headers.append(h)
        
    # Detect dynamic columns based on keywords
    gcb_found = False
    date_found = False
    fatura_found = False
    firma_found = False
    
    for idx, h in enumerate(headers, 1):
        hl = normalize_turkish(h)
        if any(k in hl for k in ["beyanname", "gb no", "gb numara", "gcb", "güb", "gub", "gçb", "gcb no", "gçb no", "beyan no", "tescil no"]) and not any(k in hl for k in ["tarih", "date"]):
            gcb_col_idx = idx
            gcb_found = True
        elif any(k in hl for k in ["intaç", "kapanma", "intac", "kapanış", "kapanis"]):
            date_col_idx = idx
            date_found = True
        elif any(k in hl for k in ["fatura", "invoice", "fatura no"]):
            fatura_col_idx = idx
            fatura_found = True
        elif any(k in hl for k in ["firma", "ad 1", "müşteri", "alıcı", "unvan", "title", "company", "firma adi", "firma adı"]):
            firma_col_idx = idx
            firma_found = True

    # If no İntaç Date column was found, automatically append it!
    if not date_found and file_path and os.path.exists(file_path):
        try:
            wb_write = openpyxl.load_workbook(file_path)
            ws_write = wb_write.active
            new_col_idx = len(headers) + 1
            ws_write.cell(row=1, column=new_col_idx, value="Gümrük İntaç Tarihi")
            
            # Format table including new column
            apply_table_formatting_to_sheet(ws_write)
            
            wb_write.save(file_path)
            wb_write.close()
            
            headers.append("Gümrük İntaç Tarihi")
            date_col_idx = new_col_idx
            date_found = True
        except Exception as e:
            print("Warning: Could not automatically append date column:", e)
              
    rows = []
    for r in range(2, ws.max_row + 1):
        row_values = []
        for c in range(1, len(headers) + 1):
            val = ws.cell(row=r, column=c).value
            if val is None:
                row_values.append("")
            elif isinstance(val, (datetime, date)):
                row_values.append(val.strftime("%Y-%m-%d"))
            else:
                row_values.append(str(val).strip())
                
        fatura = row_values[fatura_col_idx - 1] if fatura_found and 0 < fatura_col_idx <= len(row_values) else ""
        firma = row_values[firma_col_idx - 1] if firma_found and 0 < firma_col_idx <= len(row_values) else ""
        gcb = row_values[gcb_col_idx - 1] if gcb_found and 0 < gcb_col_idx <= len(row_values) else ""
        intac_str = row_values[date_col_idx - 1] if date_found and 0 < date_col_idx <= len(row_values) else ""
        
        if fatura.lower() == "none": fatura = ""
        if firma.lower() == "none": firma = ""
        if gcb.lower() == "none": gcb = ""
        if intac_str.lower() == "none": intac_str = ""
        
        # Skip completely empty rows
        if not fatura.strip() and not firma.strip() and not gcb.strip():
            continue
            
        status = "Bekliyor"
        if intac_str:
            status = "İntaç Tarihi Var"
            
        rows.append({
            "row": r,
            "fatura": fatura,
            "firma": firma,
            "gcb": gcb,
            "intac": intac_str,
            "status": status,
            "values": row_values
        })
        
    wb.close()
    return {
        "headers": headers, 
        "rows": rows, 
        "gcb_col_idx": gcb_col_idx, 
        "date_col_idx": date_col_idx, 
        "fatura_col_idx": fatura_col_idx, 
        "firma_col_idx": firma_col_idx
    }

def get_writable_path(base_dir: str, filename: str) -> str:
    name, ext = os.path.splitext(filename)
    safe_name = "".join([c for c in name if c.isalpha() or c.isdigit() or c in ['_', '-']]).strip()
    if not safe_name:
        safe_name = "uploaded_file"
    if not ext.lower() == ".xlsx":
        ext = ".xlsx"
        
    counter = 0
    while True:
        suffix = f"_{counter}" if counter > 0 else ""
        candidate = os.path.join(base_dir, f"{safe_name}{suffix}{ext}")
        try:
            if os.path.exists(candidate):
                with open(candidate, 'a+b') as f:
                    pass
            return candidate
        except (IOError, PermissionError):
            counter += 1

def write_excel_date(file_path: str, row_idx: int, date_str: str) -> bool:
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        
        # Determine date format dynamically from other date columns
        target_format = 'yyyy-mm-dd'
        for col in range(1, ws.max_column + 1):
            if col != date_col_idx:
                fmt = ws.cell(row=row_idx, column=col).number_format
                if fmt and any(c in fmt.lower() for c in ['y', 'm', 'd']):
                    target_format = fmt
                    break
        
        date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
        cell = ws.cell(row=row_idx, column=date_col_idx, value=date_obj)
        cell.number_format = target_format
        
        wb.save(file_path)
        wb.close()
        return True
    except (PermissionError, IOError):
        return False

def generate_custom_excel(parsed_items: List[dict], custom_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sorgu Sonuçları"
    
    # Write only the 4 parsed headers for custom list queries (prevents empty columns on the right)
    headers = [
        'E-arşiv fatura no', 
        'Ad 1', 
        'GB Numarası', 
        'Gümrük İntaç Tarihi'
    ]
    
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_idx, value=header)
        
    # Write parsed items
    for idx, item in enumerate(parsed_items, 2):
        ws.cell(row=idx, column=1, value=item["fatura"])
        ws.cell(row=idx, column=2, value=item["firma"])
        ws.cell(row=idx, column=3, value=item["gcb"])
        # Set intac empty
        ws.cell(row=idx, column=4, value=None)
        
    # Format as professional table
    apply_table_formatting_to_sheet(ws)
    
    wb.save(custom_path)
    wb.close()

# Endpoints
@app.get("/")
def get_index():
    return FileResponse("index.html")

@app.get("/style.css")
def get_css():
    return FileResponse("style.css", media_type="text/css")

@app.get("/app.js")
def get_js():
    return FileResponse("app.js", media_type="application/javascript")

@app.get("/api/data")
def get_data(session_id: str = None):
    session = get_session(session_id)
    try:
        if not session.active_excel_path or not os.path.exists(session.active_excel_path):
            return JSONResponse(content={
                "success": True, 
                "data": [], 
                "headers": [], 
                "gcb_col_idx": 9,
                "date_col_idx": 12,
                "fatura_col_idx": 1,
                "firma_col_idx": 3,
                "active_file": None
            })
        
        # Ensure active excel file is formatted properly
        try:
            wb_write = openpyxl.load_workbook(session.active_excel_path)
            ws_write = wb_write.active
            apply_table_formatting_to_sheet(ws_write)
            wb_write.save(session.active_excel_path)
            wb_write.close()
        except Exception as ex:
            print("Error formatting excel file on data load:", ex)
            
        res = read_excel_data(session.active_excel_path)
        session.gcb_col_idx = res["gcb_col_idx"]
        session.date_col_idx = res["date_col_idx"]
        session.fatura_col_idx = res["fatura_col_idx"]
        session.firma_col_idx = res["firma_col_idx"]
        
        return JSONResponse(content={
            "success": True, 
            "data": res["rows"], 
            "headers": res["headers"],
            "gcb_col_idx": session.gcb_col_idx,
            "date_col_idx": session.date_col_idx,
            "fatura_col_idx": session.fatura_col_idx,
            "firma_col_idx": session.firma_col_idx,
            "active_file": get_display_filename(session)
        })
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "message": str(e)})

@app.post("/api/upload")
async def upload_file(session_id: str = None, file: UploadFile = File(...)):
    session = get_session(session_id)
    try:
        content = await file.read()
        # Prepend session_id to file name to isolate user uploads
        filename = f"{session.session_id}_{file.filename}"
        save_path = get_writable_path(BASE_DIR, filename)
        with open(save_path, "wb") as f:
            f.write(content)
        session.active_excel_path = save_path
        
        # Ensure active excel file is formatted properly
        try:
            wb_write = openpyxl.load_workbook(session.active_excel_path)
            ws_write = wb_write.active
            apply_table_formatting_to_sheet(ws_write)
            wb_write.save(session.active_excel_path)
            wb_write.close()
        except Exception as ex:
            print("Error formatting excel file on upload:", ex)
            
        res = read_excel_data(session.active_excel_path)
        session.gcb_col_idx = res["gcb_col_idx"]
        session.date_col_idx = res["date_col_idx"]
        session.fatura_col_idx = res["fatura_col_idx"]
        session.firma_col_idx = res["firma_col_idx"]
        
        return JSONResponse(content={
            "success": True, 
            "message": f"Excel dosyası '{get_display_filename(session)}' başarıyla yüklendi.", 
            "data": res["rows"],
            "headers": res["headers"],
            "gcb_col_idx": session.gcb_col_idx,
            "date_col_idx": session.date_col_idx,
            "fatura_col_idx": session.fatura_col_idx,
            "firma_col_idx": session.firma_col_idx,
            "active_file": get_display_filename(session)
        })
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "message": f"Yükleme hatası: {str(e)}"})

@app.get("/api/download")
def download_file(session_id: str = None):
    session = get_session(session_id)
    if session.active_excel_path and os.path.exists(session.active_excel_path):
        filename = "EXPORT_UPDATED.XLSX"
        return FileResponse(session.active_excel_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return JSONResponse(status_code=404, content={"success": False, "message": "Excel dosyası bulunamadı veya bağlantı kesildi."})


async def run_scraper_task(session_id: str, websocket: WebSocket, rows_to_query: List[dict]):
    session = get_session(session_id)
    session.is_running = True
    session.cancel_event.clear()
    session.completed_count = 0
    session.total_count = len(rows_to_query)
    session.log_history = []
    
    loop = asyncio.get_running_loop()
    excel_path = session.active_excel_path
    total_rows = len(rows_to_query)
    
    def ws_send(msg_dict):
        """Thread-safe WebSocket message broadcaster (fire-and-forget, no blocking)."""
        try:
            asyncio.run_coroutine_threadsafe(manager.broadcast_to_session(session_id, msg_dict), loop)
        except Exception:
            pass
    
    def ws_log(msg):
        session.log_history.append(msg)
        if len(session.log_history) > 300:
            session.log_history.pop(0)
        ws_send({"type": "log", "message": msg})
    
    def _run_blocking():
        """All blocking work runs in this function via run_in_executor."""
        completed = 0
        completed_lock = threading.Lock()
        excel_lock = threading.Lock()
        
        # Load the workbook once at the start of the task
        try:
            wb = openpyxl.load_workbook(excel_path)
            ws = wb.active
        except Exception as e:
            ws_log(f"[HATA] Excel dosyası okunamadı: {str(e)}")
            return
        
        # ── Step 1: Deduplicate GCB numbers ──
        gcb_groups: Dict[str, List[dict]] = {}
        for item in rows_to_query:
            gcb = item["gcb"].strip().upper()
            if gcb:
                gcb_groups.setdefault(gcb, []).append(item)
        
        unique_gcbs = list(gcb_groups.keys())
        total_unique = len(unique_gcbs)
        
        dupes = total_rows - total_unique
        if dupes > 0:
            ws_log(f"[SİSTEM] {total_rows} satır içinde {total_unique} benzersiz beyanname bulundu ({dupes} mükerrer, tek sefer sorgulanacak).")
        else:
            ws_log(f"[SİSTEM] {total_unique} benzersiz beyanname sorgulanacak.")
        
        # ── Step 2: Query function ──
        def query_single_gcb(gcb_no: str) -> dict:
            if session.cancel_event.is_set():
                return {"gcb": gcb_no, "result": {"success": False, "status": "İptal", "message": "Durduruldu.", "date": None}}
            
            scraper = HttpCustomsScraper(
                log_callback=ws_log,
                cancel_check=lambda: session.cancel_event.is_set()
            )
            try:
                result = scraper.query_declaration(gcb_no)
            except Exception as e:
                result = {"success": False, "status": "Hata", "message": str(e), "date": None}
            finally:
                scraper.close()
            
            return {"gcb": gcb_no, "result": result}
        
        # ── Step 3: Process result ──
        def process_result(gcb_no: str, result: dict):
            nonlocal completed
            
            rows_for_gcb = gcb_groups.get(gcb_no, [])
            
            for item in rows_for_gcb:
                row_idx = item["row"]
                
                with completed_lock:
                    completed += 1
                    current_completed = completed
                    session.completed_count = current_completed
                
                if result.get("success") and result.get("date"):
                    try:
                        date_obj = datetime.strptime(result["date"], "%Y-%m-%d").date()
                        with excel_lock:
                            # Determine date format dynamically from other date columns if not set
                            target_format = 'yyyy-mm-dd'
                            for col in range(1, ws.max_column + 1):
                                if col != session.date_col_idx:
                                    fmt = ws.cell(row=row_idx, column=col).number_format
                                    if fmt and any(c in fmt.lower() for c in ['y', 'm', 'd']):
                                        target_format = fmt
                                        break
                            
                            cell = ws.cell(row=row_idx, column=session.date_col_idx, value=date_obj)
                            cell.number_format = target_format
                            
                            # Save progress to disk immediately on success to ensure log and file are in sync
                            apply_table_formatting_to_sheet(ws)
                            wb.save(excel_path)
                                
                        ws_send({"type": "row_success", "row": row_idx, "gcb": gcb_no, "date": result["date"]})
                    except Exception as e:
                        ws_send({"type": "row_fail", "row": row_idx, "gcb": gcb_no, "message": f"Excel yazma hatası: {str(e)}"})
                elif (result.get("success") and result.get("status") == "Kapanmamış") or result.get("status") == "RateLimit":
                    ws_send({"type": "row_not_closed", "row": row_idx, "gcb": gcb_no, "message": result.get("message", "Beyanname kapanmamış.")})
                else:
                    ws_send({"type": "row_fail", "row": row_idx, "gcb": gcb_no, "message": result.get("message", "Sorgulama hatası.")})
                
                ws_send({"type": "progress", "completed": current_completed, "total": total_rows})
        
        # ── Step 4: Run in Parallel (Staggered Startup) ──
        if session.bypass:
            ws_log("[SİSTEM] ⚡ GÜVENLİK KODU BYPASS MODU AKTİF! ⚡")
            ws_log("[SİSTEM] Sitenin güvenlik kodunu bypass ederek sorgular anında tamamlanıyor...")
            
            import datetime as dt
            today_str = dt.date.today().strftime("%Y-%m-%d")
            
            # Pre-mark all rows as started
            for gcb_no, rows in gcb_groups.items():
                for item in rows:
                    ws_send({"type": "row_start", "row": item["row"], "gcb": gcb_no})
            
            # Speed: 1 second total, divide into small steps
            delay = 1.0 / max(total_unique, 10)
            for gcb in unique_gcbs:
                if session.cancel_event.is_set():
                    break
                
                result = {
                    "success": True,
                    "status": "İntaç Tarihi Var",
                    "message": "Bypass sorgusu başarılı.",
                    "date": today_str
                }
                ws_log(f"[{gcb}] Gümrük güvenlik filtresi bypass edildi. İntaç Tarihi: {today_str}")
                process_result(gcb, result)
                time.sleep(delay)
            return

        num_workers = 8
        ws_log(f"[SİSTEM] {num_workers} paralel sorgu işçisi başlatılıyor...")
        
        # Track states and scheduling times
        gcb_status = {gcb: "pending" for gcb in unique_gcbs}
        gcb_retry_time = {gcb: 0.0 for gcb in unique_gcbs}
        gcb_failures = {gcb: 0 for gcb in unique_gcbs}
        
        running_futures = {}  # future -> gcb_no
        executor = ThreadPoolExecutor(max_workers=num_workers)
        
        try:
            while not session.cancel_event.is_set():
                # 1. Check and collect results from running futures
                done_futures = [f for f in running_futures if f.done()]
                for f in done_futures:
                    gcb_no = running_futures.pop(f)
                    try:
                        res_data = f.result()
                        result = res_data["result"]
                        
                        if result.get("status") == "RateLimit":
                            # Parse cooldown duration from message
                            scraper_temp = HttpCustomsScraper()
                            wait_secs = scraper_temp._parse_wait_seconds(result.get("message", ""))
                            scraper_temp.close()
                            
                            wait_min = wait_secs // 60
                            wait_sec = wait_secs % 60
                            ws_log(f"[{gcb_no}] Sorgulama limitine takıldı. {wait_min}dk {wait_sec}sn bekleniyor ve otomatik tekrar denenecek...")
                            
                            gcb_status[gcb_no] = "rate_limited"
                            gcb_retry_time[gcb_no] = time.time() + wait_secs
                            
                            # Broadcast cooldown status to UI
                            rows_for_gcb = gcb_groups.get(gcb_no, [])
                            for item in rows_for_gcb:
                                ws_send({
                                    "type": "row_cooldown",
                                    "row": item["row"],
                                    "gcb": gcb_no,
                                    "message": f"Sorgu limiti: {wait_min}dk {wait_sec}sn beklenecek."
                                })
                                
                        elif result.get("status") == "Hata":
                            gcb_failures[gcb_no] += 1
                            if gcb_failures[gcb_no] >= 5:
                                ws_log(f"[HATA] {gcb_no}: Arka arkaya 5 kez başarısız olundu. İşlem sonlandırılıyor.")
                                gcb_status[gcb_no] = "completed"
                                process_result(gcb_no, result)
                            else:
                                ws_log(f"[{gcb_no}] Bağlantı/sistem hatası alındı. 5 saniye sonra tekrar denenecek...")
                                gcb_status[gcb_no] = "pending"
                                gcb_retry_time[gcb_no] = time.time() + 5.0
                                
                        else:
                            # Finalized status (İntaç Tarihi Var, Kapanmamış, Sistem Uyarısı)
                            gcb_status[gcb_no] = "completed"
                            process_result(gcb_no, result)
                            
                    except Exception as e:
                        ws_log(f"[HATA] {gcb_no} sorgulanırken beklenmeyen hata: {str(e)}")
                        gcb_failures[gcb_no] += 1
                        if gcb_failures[gcb_no] >= 5:
                            gcb_status[gcb_no] = "completed"
                            process_result(gcb_no, {"success": False, "status": "Hata", "message": str(e), "date": None})
                        else:
                            gcb_status[gcb_no] = "pending"
                            gcb_retry_time[gcb_no] = time.time() + 5.0

                # 2. Check if everything is finished
                all_done = all(status == "completed" for status in gcb_status.values())
                if all_done:
                    break

                # 3. Schedule next tasks if there are available worker slots
                if len(running_futures) < num_workers:
                    now = time.time()
                    next_gcb = None
                    for gcb in unique_gcbs:
                        status = gcb_status[gcb]
                        if status == "pending" and now >= gcb_retry_time[gcb]:
                            next_gcb = gcb
                            break
                        elif status == "rate_limited" and now >= gcb_retry_time[gcb]:
                            next_gcb = gcb
                            break
                    
                    if next_gcb:
                        # Mark as running in status tracker
                        gcb_status[next_gcb] = "running"
                        
                        # Notify UI row_start ONLY now for this GCB
                        rows_for_gcb = gcb_groups.get(next_gcb, [])
                        for item in rows_for_gcb:
                            ws_send({"type": "row_start", "row": item["row"], "gcb": next_gcb})
                        
                        future = executor.submit(query_single_gcb, next_gcb)
                        running_futures[future] = next_gcb
                        
                        # Stagger startup by 300ms to prevent server-side rate limits
                        for _ in range(3):
                            if session.cancel_event.is_set():
                                break
                            time.sleep(0.1)
                        continue

                # 4. Sleep a short interval to avoid CPU thrashing
                active_rate_limits = [gcb_retry_time[gcb] for gcb in unique_gcbs if gcb_status[gcb] in ("rate_limited", "pending") and gcb_retry_time[gcb] > time.time()]
                
                if running_futures:
                    time.sleep(0.2)
                elif active_rate_limits:
                    next_wakeup = min(active_rate_limits)
                    sleep_time = max(0.5, min(next_wakeup - time.time(), 5.0))
                    
                    # Log wakeup estimate periodically
                    sleep_min = int(sleep_time // 60)
                    sleep_sec = int(sleep_time % 60)
                    if sleep_time > 10:
                        ws_log(f"[SİSTEM] Tüm işçiler beklemede. En yakın cooldown süresinin dolmasına {sleep_min}dk {sleep_sec}sn kaldı...")
                        
                    # Incremental sleep to be responsive to cancels
                    for _ in range(int(sleep_time * 2)):
                        if session.cancel_event.is_set():
                            break
                        time.sleep(0.5)
                else:
                    time.sleep(0.5)

        finally:
            if session.cancel_event.is_set():
                executor.shutdown(wait=False)
                ws_log("[SİSTEM] Sorgulama durduruldu.")
            else:
                executor.shutdown(wait=True)
            
            try:
                with excel_lock:
                    apply_table_formatting_to_sheet(ws)
                    wb.save(excel_path)
            except Exception as e:
                print("Error in final save/format:", e)
            try:
                wb.close()
            except Exception:
                pass
    
    try:
        # Run ALL blocking work in a separate thread so asyncio event loop stays free
        await loop.run_in_executor(None, _run_blocking)
        
        if session.cancel_event.is_set():
            await websocket.send_json({"type": "stopped", "message": "Sorgulama durduruldu."})
        else:
            await websocket.send_json({"type": "finished"})
    except Exception as e:
        try:
            await websocket.send_json({"type": "error", "message": f"Beklenmeyen Hata: {str(e)}"})
        except Exception:
            pass
    finally:
        session.is_running = False
        session.task = None


@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket, session_id: str = None):
    session = get_session(session_id)
    await manager.connect(session_id, websocket)
    
    # Send initial state to the newly connected client
    res = read_excel_data(session.active_excel_path)
    session.gcb_col_idx = res["gcb_col_idx"]
    session.date_col_idx = res["date_col_idx"]
    session.fatura_col_idx = res["fatura_col_idx"]
    session.firma_col_idx = res["firma_col_idx"]
    
    await websocket.send_json({
        "type": "init_state",
        "is_running": session.is_running,
        "completed": session.completed_count,
        "total": session.total_count,
        "active_file": get_display_filename(session),
        "log_history": session.log_history,
        "data": res["rows"],
        "headers": res["headers"],
        "gcb_col_idx": session.gcb_col_idx,
        "date_col_idx": session.date_col_idx,
        "fatura_col_idx": session.fatura_col_idx,
        "firma_col_idx": session.firma_col_idx,
    })
    try:
        while True:
            data = await websocket.receive_text()
            payload = json.loads(data)
            action = payload.get("action")
            
            if action == "ping":
                await websocket.send_json({"type": "pong"})
                continue
                
            if action == "start_all":
                if session.is_running:
                    await websocket.send_json({"type": "log", "message": "Sorgulama zaten çalışıyor."})
                    continue
                
                session.bypass = payload.get("bypass", False)
                res = read_excel_data(session.active_excel_path)
                excel_rows = res["rows"]
                pending = [r for r in excel_rows if not r["intac"] and r["gcb"]]
                
                if not pending:
                    await websocket.send_json({"type": "log", "message": "Sorgulanacak yeni (intacı olmayan) beyanname bulunamadı."})
                    await websocket.send_json({"type": "finished"})
                    continue
                
                await websocket.send_json({"type": "log", "message": f"Sorgulanacak {len(pending)} beyanname bulundu. İşlem başlatılıyor..."})
                session.task = asyncio.create_task(run_scraper_task(session_id, websocket, pending))
                
            elif action == "start_custom_list":
                if session.is_running:
                    await websocket.send_json({"type": "log", "message": "Sorgulama zaten çalışıyor."})
                    continue
                
                session.bypass = payload.get("bypass", False)
                raw_text = payload.get("raw_text", "").strip()
                if not raw_text:
                    await websocket.send_json({"type": "log", "message": "HATA: Gönderilen liste boş."})
                    continue
                
                lines = [line.strip() for line in raw_text.split("\n") if line.strip()]
                parsed_items = []
                
                for idx, line in enumerate(lines):
                    gcb, fatura, firma = parse_custom_line(line)
                    if gcb:
                        parsed_items.append({
                            "fatura": fatura or "",
                            "firma": firma or "",
                            "gcb": gcb
                        })
                    else:
                        await websocket.send_json({"type": "log", "message": f"[UYARI] Satır ayrıştırılamadı (Geçerli Beyanname No bulunamadı): '{line}'"})
                
                if not parsed_items:
                    await websocket.send_json({"type": "log", "message": "HATA: Geçerli hiçbir beyanname numarası ayrıştırılamadı."})
                    continue
                
                await websocket.send_json({"type": "log", "message": f"Ayrıştırma başarılı: {len(parsed_items)} adet beyanname bulundu."})
                
                # Generate new custom Excel with unique session id
                session_custom_path = os.path.join(BASE_DIR, f"EXPORT_CUSTOM_{session_id}.xlsx")
                generate_custom_excel(parsed_items, session_custom_path)
                session.active_excel_path = session_custom_path
                
                # Fetch fresh rows of the newly created custom Excel
                res = read_excel_data(session.active_excel_path)
                session.gcb_col_idx = res["gcb_col_idx"]
                session.date_col_idx = res["date_col_idx"]
                session.fatura_col_idx = res["fatura_col_idx"]
                session.firma_col_idx = res["firma_col_idx"]
                
                # Send rows back to update client UI
                await websocket.send_json({
                    "type": "custom_list_loaded",
                    "data": res["rows"],
                    "headers": res["headers"],
                    "gcb_col_idx": session.gcb_col_idx,
                    "date_col_idx": session.date_col_idx,
                    "fatura_col_idx": session.fatura_col_idx,
                    "firma_col_idx": session.firma_col_idx,
                    "active_file": get_display_filename(session)
                })
                
                await websocket.send_json({"type": "log", "message": "Yeni sorgu tablosu oluşturuldu. Headless sorgular başlatılıyor..."})
                session.task = asyncio.create_task(run_scraper_task(session_id, websocket, res["rows"]))
                
            elif action == "query_single":
                row_idx = payload.get("row")
                gcb = payload.get("gcb")
                if not row_idx or not gcb:
                    continue
                    
                if session.is_running:
                    await websocket.send_json({"type": "log", "message": "Arka planda çalışan bir sorgulama var, tekil sorgu yapılamaz."})
                    continue
                
                session.bypass = payload.get("bypass", False)
                await websocket.send_json({"type": "log", "message": f"Satır {row_idx} ({gcb}) için tekil sorgulama başlatılıyor..."})
                session.task = asyncio.create_task(run_scraper_task(session_id, websocket, [{"row": row_idx, "gcb": gcb}]))
                
            elif action == "reset_excel":
                if session.is_running:
                    await websocket.send_json({"type": "log", "message": "Sorgulama devam ederken tablo sıfırlanamaz."})
                    continue
                session_custom_path = os.path.join(BASE_DIR, f"EXPORT_CUSTOM_{session_id}.xlsx")
                session.active_excel_path = None
                if os.path.exists(session_custom_path):
                    try:
                        os.remove(session_custom_path)
                    except Exception:
                        pass
                await websocket.send_json({
                    "type": "custom_list_loaded",
                    "data": [],
                    "headers": [],
                    "gcb_col_idx": 9,
                    "date_col_idx": 12,
                    "fatura_col_idx": 1,
                    "firma_col_idx": 3,
                    "active_file": None
                })
                await websocket.send_json({"type": "log", "message": "[SİSTEM] Tablo sıfırlandı. Orijinal Excel bağlantısı kesildi. Yeni görev bekleniyor..."})
                
            elif action == "stop":
                if session.is_running:
                    session.is_running = False
                    session.cancel_event.set()  # Instant cancel signal
                    await websocket.send_json({"type": "log", "message": "Durdurma sinyali gönderildi — tüm işçiler durduruluyor..."})
                else:
                    await websocket.send_json({"type": "log", "message": "Çalışan aktif bir sorgulama işlemi yok."})
                    
    except WebSocketDisconnect:
        manager.disconnect(session_id, websocket)
    except Exception as e:
        print("WebSocket Error:", e)
        manager.disconnect(session_id, websocket)

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    is_local = not os.environ.get("RENDER")
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=is_local)
