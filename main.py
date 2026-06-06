import os
import re
import json
import asyncio
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, date
from typing import List, Dict, Any

import openpyxl
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

# Global configuration to track active Excel file
active_excel_path = EXCEL_PATH

gcb_col_idx = 9
date_col_idx = 12
fatura_col_idx = 1
firma_col_idx = 3
headers_list = []

# Keep track of active WebSocket connections
class ConnectionManager:
    def __init__(self):
        self.active_connections: List[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        if websocket in self.active_connections:
            self.active_connections.remove(websocket)

    async def broadcast(self, message: dict):
        for connection in self.active_connections:
            try:
                await connection.send_json(message)
            except Exception:
                pass

manager = ConnectionManager()

# Background task state
class ScraperState:
    def __init__(self):
        self.is_running = False
        self.task = None
        self.cancel_event = threading.Event()

state = ScraperState()

def normalize_turkish(text: str) -> str:
    if not text:
        return ""
    mapping = {
        'İ': 'i', 'I': 'ı', 'Ş': 'ş', 'Ç': 'ç', 'Ğ': 'ğ', 'Ü': 'ü', 'Ö': 'ö',
        'ı': 'ı', 'ş': 'ş', 'ç': 'ç', 'ğ': 'ğ', 'ü': 'ü', 'ö': 'ö', 'i': 'i'
    }
    return "".join(mapping.get(c, c.lower()) for c in text)

# Robust line parser for custom paste strings
def parse_custom_line(line: str):
    # Split by spaces, tabs, commas, or semicolons
    clean_line = line.replace('\t', ' ').replace(';', ' ').replace(',', ' ')
    parts = [p.strip() for p in clean_line.split() if p.strip()]
    
    gcb = None
    fatura = None
    
    # 1. Regex match GCB No: e.g. 26341200EX00137190 (8 digits, 2 letters, 6 to 8 digits)
    gcb_pattern = re.compile(r'^\d{8}[A-Za-z]{2}\d{6,8}$')
    for p in parts:
        if gcb_pattern.match(p):
            gcb = p
            break
    if gcb:
        parts.remove(gcb)
        
    # 2. Match Fatura No: alphanumeric string containing digits, starts with BT/BTC or similar
    fatura_pattern = re.compile(r'^(BT[C]?\d+|[A-Za-z0-9\-]+)$')
    for p in parts:
        if fatura_pattern.match(p) and any(c.isdigit() for c in p) and len(p) >= 4:
            fatura = p
            break
    if fatura:
        parts.remove(fatura)
        
    # 3. Remaining parts are Company Name
    firma = " ".join(parts) if parts else ""
    
    return gcb, fatura, firma

def read_excel_data(file_path: str) -> Dict[str, Any]:
    global gcb_col_idx, date_col_idx, fatura_col_idx, firma_col_idx, headers_list
    
    # Reset defaults in case of empty or missing spreadsheet
    gcb_col_idx = 9
    date_col_idx = 12
    fatura_col_idx = 1
    firma_col_idx = 3
    headers_list = []
    
    if not file_path or not os.path.exists(file_path):
        return {"headers": [], "rows": []}
    
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
            wb_write.save(file_path)
            wb_write.close()
            
            headers.append("Gümrük İntaç Tarihi")
            date_col_idx = new_col_idx
            date_found = True
        except Exception as e:
            print("Warning: Could not automatically append date column:", e)
              
    headers_list = headers
    
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
                
        fatura = row_values[fatura_col_idx - 1] if 0 < fatura_col_idx <= len(row_values) else ""
        firma = row_values[firma_col_idx - 1] if 0 < firma_col_idx <= len(row_values) else ""
        gcb = row_values[gcb_col_idx - 1] if 0 < gcb_col_idx <= len(row_values) else ""
        intac_str = row_values[date_col_idx - 1] if 0 < date_col_idx <= len(row_values) else ""
        
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
    return {"headers": headers, "rows": rows}

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

def generate_custom_excel(parsed_items: List[dict]):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sorgu Sonuçları"
    
    # Write exactly identical headers for fully compatible structure
    headers = [
        'E-arşiv fatura no', 'Faturalama tarihi', 'Ad 1', 'Toplam CIF tutarı', 'Toplam FOB', 
        'Para birimi', 'FOB USD', 'Para birimi', 'GB Numarası', 'GB Tarihi', 'Gümrük', 
        'Gümrük İntaç Tarihi', 'Incoterms', 'Satış temsilcisi', 'Satış temsilcisi adı', 
        'Teslimat', 'Ram Fatura No', 'Toplam Palet', 'Toplam Koli', 'Toplam Sandık Paleti', 
        'Toplam Ağırlık', 'Toplam Net Ağırlık', 'Toplam miktar', 'Poliçe No', 'Evrak Sorumlusu 2'
    ]
    
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_idx, value=header)
        
    # Write parsed items
    for idx, item in enumerate(parsed_items, 2):
        ws.cell(row=idx, column=1, value=item["fatura"])
        ws.cell(row=idx, column=3, value=item["firma"])
        ws.cell(row=idx, column=9, value=item["gcb"])
        # Set intac empty
        ws.cell(row=idx, column=12, value=None)
        
    wb.save(EXCEL_CUSTOM_PATH)
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
def get_data():
    try:
        if not active_excel_path or not os.path.exists(active_excel_path):
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
        res = read_excel_data(active_excel_path)
        return JSONResponse(content={
            "success": True, 
            "data": res["rows"], 
            "headers": res["headers"],
            "gcb_col_idx": gcb_col_idx,
            "date_col_idx": date_col_idx,
            "fatura_col_idx": fatura_col_idx,
            "firma_col_idx": firma_col_idx,
            "active_file": os.path.basename(active_excel_path)
        })
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "message": str(e)})

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    global active_excel_path
    try:
        content = await file.read()
        save_path = get_writable_path(BASE_DIR, file.filename)
        with open(save_path, "wb") as f:
            f.write(content)
        active_excel_path = save_path
        res = read_excel_data(active_excel_path)
        return JSONResponse(content={
            "success": True, 
            "message": f"Excel dosyası '{os.path.basename(active_excel_path)}' başarıyla yüklendi.", 
            "data": res["rows"],
            "headers": res["headers"],
            "gcb_col_idx": gcb_col_idx,
            "date_col_idx": date_col_idx,
            "fatura_col_idx": fatura_col_idx,
            "firma_col_idx": firma_col_idx,
            "active_file": os.path.basename(active_excel_path)
        })
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "message": f"Yükleme hatası: {str(e)}"})

@app.get("/api/download")
def download_file():
    if active_excel_path and os.path.exists(active_excel_path):
        filename = "EXPORT_UPDATED.XLSX" if active_excel_path == EXCEL_PATH else "EXPORT_CUSTOM_UPDATED.XLSX"
        return FileResponse(active_excel_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return JSONResponse(status_code=404, content={"success": False, "message": "Excel dosyası bulunamadı veya bağlantı kesildi."})


async def run_scraper_task(websocket: WebSocket, rows_to_query: List[dict]):
    state.is_running = True
    state.cancel_event.clear()
    loop = asyncio.get_running_loop()
    
    excel_path = active_excel_path
    total_rows = len(rows_to_query)
    
    def ws_send(msg_dict):
        """Thread-safe WebSocket message sender (fire-and-forget, no blocking)."""
        try:
            asyncio.run_coroutine_threadsafe(websocket.send_json(msg_dict), loop)
        except Exception:
            pass
    
    def ws_log(msg):
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
            if state.cancel_event.is_set():
                return {"gcb": gcb_no, "result": {"success": False, "status": "İptal", "message": "Durduruldu.", "date": None}}
            
            scraper = HttpCustomsScraper(
                log_callback=ws_log,
                cancel_check=lambda: state.cancel_event.is_set()
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
                
                if result.get("success") and result.get("date"):
                    try:
                        date_obj = datetime.strptime(result["date"], "%Y-%m-%d").date()
                        with excel_lock:
                            # Determine date format dynamically from other date columns if not set
                            target_format = 'yyyy-mm-dd'
                            for col in range(1, ws.max_column + 1):
                                if col != date_col_idx:
                                    fmt = ws.cell(row=row_idx, column=col).number_format
                                    if fmt and any(c in fmt.lower() for c in ['y', 'm', 'd']):
                                        target_format = fmt
                                        break
                            
                            cell = ws.cell(row=row_idx, column=date_col_idx, value=date_obj)
                            cell.number_format = target_format
                            wb.save(excel_path)
                        ws_send({"type": "row_success", "row": row_idx, "gcb": gcb_no, "date": result["date"]})
                    except Exception as e:
                        ws_send({"type": "row_fail", "row": row_idx, "gcb": gcb_no, "message": f"Excel yazma hatası: {str(e)}"})
                elif result.get("success") and result.get("status") == "Kapanmamış":
                    ws_send({"type": "row_not_closed", "row": row_idx, "gcb": gcb_no, "message": "Beyanname kapanmamış."})
                else:
                    ws_send({"type": "row_fail", "row": row_idx, "gcb": gcb_no, "message": result.get("message", "Sorgulama hatası.")})
                
                with completed_lock:
                    completed += 1
                    ws_send({"type": "progress", "completed": completed, "total": total_rows})
        
        # ── Step 4: Run with ThreadPoolExecutor ──
        num_workers = min(5, total_unique)
        ws_log(f"[SİSTEM] {num_workers} paralel HTTP işçisi başlatılıyor (Bellek dostu limit: 5)...")
        
        # Mark all rows as started
        for gcb_no, rows in gcb_groups.items():
            for item in rows:
                ws_send({"type": "row_start", "row": item["row"], "gcb": gcb_no})
        
        try:
            with ThreadPoolExecutor(max_workers=num_workers) as executor:
                future_to_gcb = {
                    executor.submit(query_single_gcb, gcb): gcb
                    for gcb in unique_gcbs
                }
                
                for future in as_completed(future_to_gcb):
                    if state.cancel_event.is_set():
                        for f in future_to_gcb:
                            f.cancel()
                        ws_log("[SİSTEM] Sorgulama durduruldu.")
                        break
                    
                    try:
                        data = future.result()
                        process_result(data["gcb"], data["result"])
                    except Exception as e:
                        gcb = future_to_gcb[future]
                        ws_log(f"[HATA] {gcb}: {str(e)}")
                        process_result(gcb, {"success": False, "status": "Hata", "message": str(e), "date": None})
        finally:
            try:
                wb.close()
            except Exception:
                pass
    
    try:
        # Run ALL blocking work in a separate thread so asyncio event loop stays free
        await loop.run_in_executor(None, _run_blocking)
        
        if state.cancel_event.is_set():
            await websocket.send_json({"type": "stopped", "message": "Sorgulama durduruldu."})
        else:
            await websocket.send_json({"type": "finished"})
    except Exception as e:
        try:
            await websocket.send_json({"type": "error", "message": f"Beklenmeyen Hata: {str(e)}"})
        except Exception:
            pass
    finally:
        state.is_running = False
        state.task = None


@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    global active_excel_path
    await manager.connect(websocket)
    try:
        while True:
            data = await websocket.receive_text()
            payload = json.loads(data)
            action = payload.get("action")
            
            if action == "ping":
                await websocket.send_json({"type": "pong"})
                continue
                
            if action == "start_all":
                if state.is_running:
                    await websocket.send_json({"type": "log", "message": "Sorgulama zaten çalışıyor."})
                    continue
                
                res = read_excel_data(active_excel_path)
                excel_rows = res["rows"]
                pending = [r for r in excel_rows if not r["intac"] and r["gcb"]]
                
                if not pending:
                    await websocket.send_json({"type": "log", "message": "Sorgulanacak yeni (intacı olmayan) beyanname bulunamadı."})
                    await websocket.send_json({"type": "finished"})
                    continue
                
                await websocket.send_json({"type": "log", "message": f"Sorgulanacak {len(pending)} beyanname bulundu. İşlem başlatılıyor..."})
                state.task = asyncio.create_task(run_scraper_task(websocket, pending))
                
            elif action == "start_custom_list":
                if state.is_running:
                    await websocket.send_json({"type": "log", "message": "Sorgulama zaten çalışıyor."})
                    continue
                
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
                
                # Generate new custom Excel
                generate_custom_excel(parsed_items)
                active_excel_path = EXCEL_CUSTOM_PATH
                
                # Fetch fresh rows of the newly created custom Excel
                res = read_excel_data(active_excel_path)
                custom_rows = res["rows"]
                
                # Send rows back to update client UI
                await websocket.send_json({
                    "type": "custom_list_loaded",
                    "data": custom_rows,
                    "headers": res["headers"],
                    "gcb_col_idx": gcb_col_idx,
                    "date_col_idx": date_col_idx,
                    "fatura_col_idx": fatura_col_idx,
                    "firma_col_idx": firma_col_idx,
                    "active_file": os.path.basename(active_excel_path)
                })
                
                await websocket.send_json({"type": "log", "message": "Yeni sorgu tablosu oluşturuldu. Headless sorgular başlatılıyor..."})
                state.task = asyncio.create_task(run_scraper_task(websocket, custom_rows))
                
            elif action == "query_single":
                row_idx = payload.get("row")
                gcb = payload.get("gcb")
                if not row_idx or not gcb:
                    continue
                    
                if state.is_running:
                    await websocket.send_json({"type": "log", "message": "Arka planda çalışan bir sorgulama var, tekil sorgu yapılamaz."})
                    continue
                
                await websocket.send_json({"type": "log", "message": f"Satır {row_idx} ({gcb}) için tekil sorgulama başlatılıyor..."})
                state.task = asyncio.create_task(run_scraper_task(websocket, [{"row": row_idx, "gcb": gcb}]))
                
            elif action == "reset_excel":
                if state.is_running:
                    await websocket.send_json({"type": "log", "message": "Sorgulama devam ederken tablo sıfırlanamaz."})
                    continue
                active_excel_path = None
                if os.path.exists(EXCEL_CUSTOM_PATH):
                    try:
                        os.remove(EXCEL_CUSTOM_PATH)
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
                if state.is_running:
                    state.is_running = False
                    state.cancel_event.set()  # Instant cancel signal
                    await websocket.send_json({"type": "log", "message": "Durdurma sinyali gönderildi — tüm işçiler durduruluyor..."})
                else:
                    await websocket.send_json({"type": "log", "message": "Çalışan aktif bir sorgulama işlemi yok."})
                    
    except WebSocketDisconnect:
        manager.disconnect(websocket)
        if state.is_running:
            state.is_running = False
    except Exception as e:
        print("WebSocket Error:", e)
        manager.disconnect(websocket)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
