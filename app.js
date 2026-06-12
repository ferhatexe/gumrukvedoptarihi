class SoundEngine {
    constructor() {
        this.ctx = null;
        this.initialized = false;
    }
    
    init() {
        if (this.initialized) return;
        try {
            this.ctx = new (window.AudioContext || window.webkitAudioContext)();
            this.initialized = true;
            console.log("[SES] SoundEngine initialized successfully.");
        } catch (e) {
            console.warn("[SES] Web Audio API not supported or blocked:", e);
        }
    }
    
    playSuccess() {
        this.init();
        if (!this.ctx) return;
        const now = this.ctx.currentTime;
        // Ascending major arpeggio (C5 -> E5 -> G5 -> C6)
        const notes = [523.25, 659.25, 783.99, 1046.50];
        notes.forEach((freq, index) => {
            const osc = this.ctx.createOscillator();
            const gain = this.ctx.createGain();
            osc.connect(gain);
            gain.connect(this.ctx.destination);
            
            osc.type = "sine";
            osc.frequency.setValueAtTime(freq, now + index * 0.08);
            
            gain.gain.setValueAtTime(0, now + index * 0.08);
            gain.gain.linearRampToValueAtTime(0.45, now + index * 0.08 + 0.02); // Louder volume (0.45)
            gain.gain.exponentialRampToValueAtTime(0.0001, now + index * 0.08 + 0.35);
            
            osc.start(now + index * 0.08);
            osc.stop(now + index * 0.08 + 0.45);
        });
    }
    
    playNotClosed() {
        this.init();
        if (!this.ctx) return;
        const now = this.ctx.currentTime;
        // Descending minor interval (A4 -> F4)
        const notes = [440.00, 349.23];
        notes.forEach((freq, index) => {
            const osc = this.ctx.createOscillator();
            const gain = this.ctx.createGain();
            osc.connect(gain);
            gain.connect(this.ctx.destination);
            
            osc.type = "triangle";
            osc.frequency.setValueAtTime(freq, now + index * 0.12);
            
            gain.gain.setValueAtTime(0, now + index * 0.12);
            gain.gain.linearRampToValueAtTime(0.4, now + index * 0.12 + 0.02); // Louder volume (0.4)
            gain.gain.exponentialRampToValueAtTime(0.0001, now + index * 0.12 + 0.4);
            
            osc.start(now + index * 0.12);
            osc.stop(now + index * 0.12 + 0.45);
        });
    }
    
    playError() {
        this.init();
        if (!this.ctx) return;
        const now = this.ctx.currentTime;
        // Deep warning double-buzz (E3 -> C3)
        const notes = [164.81, 130.81];
        notes.forEach((freq, index) => {
            const osc = this.ctx.createOscillator();
            const gain = this.ctx.createGain();
            osc.connect(gain);
            gain.connect(this.ctx.destination);
            
            osc.type = "sawtooth";
            osc.frequency.setValueAtTime(freq, now + index * 0.18);
            
            gain.gain.setValueAtTime(0, now + index * 0.18);
            gain.gain.linearRampToValueAtTime(0.4, now + index * 0.18 + 0.04); // Louder volume (0.4)
            gain.gain.exponentialRampToValueAtTime(0.0001, now + index * 0.18 + 0.3);
            
            osc.start(now + index * 0.18);
            osc.stop(now + index * 0.18 + 0.35);
        });
    }
    
    playUploadSuccess() {
        this.init();
        if (!this.ctx) return;
        const now = this.ctx.currentTime;
        // High bright "ding" (C6)
        const osc = this.ctx.createOscillator();
        const gain = this.ctx.createGain();
        osc.connect(gain);
        gain.connect(this.ctx.destination);
        
        osc.type = "sine";
        osc.frequency.setValueAtTime(1046.50, now);
        
        gain.gain.setValueAtTime(0, now);
        gain.gain.linearRampToValueAtTime(0.45, now + 0.01); // Louder volume (0.45)
        gain.gain.exponentialRampToValueAtTime(0.0001, now + 0.25);
        
        osc.start(now);
        osc.stop(now + 0.35);
    }

    playReset() {
        this.init();
        if (!this.ctx) return;
        const now = this.ctx.currentTime;
        // Descending sweeping notes (C4 -> G3 -> C3)
        const notes = [261.63, 196.00, 130.81];
        notes.forEach((freq, index) => {
            const osc = this.ctx.createOscillator();
            const gain = this.ctx.createGain();
            osc.connect(gain);
            gain.connect(this.ctx.destination);
            
            osc.type = "sine";
            osc.frequency.setValueAtTime(freq, now + index * 0.1);
            
            gain.gain.setValueAtTime(0, now + index * 0.1);
            gain.gain.linearRampToValueAtTime(0.45, now + index * 0.1 + 0.02); // Premium sweep volume (0.45)
            gain.gain.exponentialRampToValueAtTime(0.0001, now + index * 0.1 + 0.4);
            
            osc.start(now + index * 0.1);
            osc.stop(now + index * 0.1 + 0.5);
        });
    }
}

const soundEngine = new SoundEngine();

let socket = null;
let heartbeatInterval = null;
let allRows = [];
let queryStatus = {}; // Keep track of current session query status per row index
let queryStartTime = null;
let queryTimerInterval = null;
let totalCaptchaAttempts = 0;
let lastUpdatedRow = null;
let lastUpdatedTime = 0;
let activeHeaders = []; // Dynamic headers from active Excel
let activeGcbColIdx = 9; // Index of GCB column (1-based)
let activeDateColIdx = 12; // Index of Date column (1-based)
let activeFaturaColIdx = 1; // Index of Fatura column (1-based)
const STATUS_PRIORITY = {
    "İntaç Tarihi Var": 10,
    "Kapanmamış": 8,
    "Soğumada": 6,
    "Sorgulanıyor...": 5,
    "Başarısız": 4,
    "Hatalı": 4,
    "Bekliyor": 2
};

function getGcbConsolidatedStatus(rows) {
    let maxPriority = -1;
    let bestStatus = "Bekliyor";
    
    rows.forEach(item => {
        let status;
        const qs = queryStatus[item.row];
        if (item.intac && item.status === "İntaç Tarihi Var") {
            status = "İntaç Tarihi Var";
        } else if (qs) {
            status = qs.status;
        } else {
            status = item.status || "Bekliyor";
        }
        
        const priority = STATUS_PRIORITY[status] || 1;
        if (priority > maxPriority) {
            maxPriority = priority;
            bestStatus = status;
        }
    });
    
    return bestStatus;
}

// Session Management: Generate or retrieve session ID
function getOrCreateSessionId() {
    let sessionId = localStorage.getItem("gumruk_session_id");
    if (!sessionId) {
        try {
            const arr = new Uint8Array(16);
            window.crypto.getRandomValues(arr);
            sessionId = Array.from(arr, dec => dec.toString(16).padStart(2, '0')).join('');
        } catch (e) {
            sessionId = 'sess_' + Math.random().toString(36).substring(2, 15) + '_' + Date.now();
        }
        localStorage.setItem("gumruk_session_id", sessionId);
    }
    return sessionId;
}
const sessionId = getOrCreateSessionId();

// Strip session hash prefix from displayed filename (e.g. "abc123_EXPORT.XLSX" -> "EXPORT.XLSX")
function cleanFileName(name) {
    if (!name) return name;
    // If filename starts with the session id prefix, strip it
    const prefix = sessionId + "_";
    if (name.startsWith(prefix)) {
        return name.substring(prefix.length);
    }
    // Also strip any generic 32-char hex prefix (md5-style hash)
    const hexMatch = name.match(/^[0-9a-f]{32}_(.+)$/i);
    if (hexMatch) {
        return hexMatch[1];
    }
    return name;
}

// DOM Elements
const connectionDot = document.getElementById("connection-status-dot");
const connectionText = document.getElementById("connection-status-text");
const activeFileBadge = document.getElementById("active-file-badge");
const statTotal = document.getElementById("stat-total");
const statCompleted = document.getElementById("stat-completed");
const statPending = document.getElementById("stat-pending");
const statNotClosed = document.getElementById("stat-not-closed");
const statErrors = document.getElementById("stat-errors");
const statCooldown = document.getElementById("stat-cooldown");
const btnStart = document.getElementById("btn-start");
const btnStop = document.getElementById("btn-stop");
const btnDownload = document.getElementById("btn-download");
const progressBarFill = document.getElementById("progress-bar-fill");
const progressPercent = document.getElementById("progress-percent");
const terminal = document.getElementById("terminal");
const btnClearTerminal = document.getElementById("btn-clear-terminal");
const searchInput = document.getElementById("search-input");
const tableBody = document.getElementById("table-body");
const btnResetList = document.getElementById("btn-reset-list");

// Text Paste & Upload Elements
const rawInput = document.getElementById("raw-input");
const btnParseQuery = document.getElementById("btn-parse-query");
const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");

// Tab Selectors
const tabButtons = document.querySelectorAll(".tab-btn");

// Init
document.addEventListener("DOMContentLoaded", () => {
    loadExcelData();
    connectWebSocket();
    setupFileUpload();
    
    // Set the session ID on download button
    if (btnDownload) {
        btnDownload.href = `/api/download?session_id=${sessionId}`;
    }
    
    // Event listeners
    btnStart.addEventListener("click", startAutomation);
    btnStop.addEventListener("click", stopAutomation);
    btnParseQuery.addEventListener("click", startCustomListAutomation);
    btnResetList.addEventListener("click", resetExcelTable);
    
    // Tab switching event binding
    tabButtons.forEach(btn => {
        btn.addEventListener("click", () => switchTab(btn.dataset.tab));
    });
    
    btnClearTerminal.addEventListener("click", () => {
        terminal.innerHTML = '<div class="terminal-line system">[SİSTEM] Terminal temizlendi.</div>';
    });
    searchInput.addEventListener("input", filterTable);
    document.getElementById("filter-status").addEventListener("change", function() {
        this.classList.toggle("filter-active", this.value !== "all");
        filterTable();
    });
    document.getElementById("filter-type").addEventListener("change", function() {
        this.classList.toggle("filter-active", this.value !== "all");
        filterTable();
    });
    
    // Input mode switching (Excel Upload / Serbest Metin)
    const btnModeExcel = document.getElementById("btn-mode-excel");
    const btnModeText = document.getElementById("btn-mode-text");
    const sectionExcel = document.getElementById("section-mode-excel");
    const sectionText = document.getElementById("section-mode-text");
    
    if (btnModeExcel && btnModeText && sectionExcel && sectionText) {
        btnModeExcel.addEventListener("click", () => {
            btnModeExcel.classList.add("active");
            btnModeText.classList.remove("active");
            sectionExcel.classList.add("active");
            sectionText.classList.remove("active");
        });
        btnModeText.addEventListener("click", () => {
            btnModeText.classList.add("active");
            btnModeExcel.classList.remove("active");
            sectionText.classList.add("active");
            sectionExcel.classList.remove("active");
        });
    }
    
    // Start the cooldown ticker
    setInterval(updateCooldowns, 1000);
});

// UI tab switching helper
function switchTab(tabId) {
    document.querySelectorAll(".tab-btn").forEach(btn => btn.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach(content => content.classList.remove("active"));
    
    const activeBtn = document.querySelector(`.tab-btn[data-tab="${tabId}"]`);
    const activeContent = document.getElementById(tabId);
    if (activeBtn) activeBtn.classList.add("active");
    if (activeContent) activeContent.classList.add("active");
}

// Load Excel data from backend API
async function loadExcelData() {
    try {
        const res = await fetch(`/api/data?session_id=${sessionId}`);
        const json = await res.json();
        if (json.success) {
            allRows = json.data;
            activeHeaders = json.headers || [];
            activeGcbColIdx = json.gcb_col_idx || 9;
            activeDateColIdx = json.date_col_idx || 12;
            activeFaturaColIdx = json.fatura_col_idx || 1;
            activeFirmaColIdx = json.firma_col_idx || 3;
            
            if (json.active_file) {
                activeFileBadge.innerText = `Aktif Dosya: ${cleanFileName(json.active_file)}`;
                activeFileBadge.className = "active-file-text loaded";
                btnDownload.classList.remove("disabled-btn");
            } else {
                activeFileBadge.innerText = "Aktif Dosya: Yok (Yeni Görev Bekleniyor)";
                activeFileBadge.className = "active-file-text empty";
                btnDownload.classList.add("disabled-btn");
            }
            renderTable(allRows);
            updateStats();
        } else {
            addTerminalLine("HATA: Excel verisi yüklenemedi: " + json.message, "error");
        }
    } catch (e) {
        addTerminalLine("HATA: Sunucuya bağlanılamadı: " + e.message, "error");
    }
}

// Connect to WebSocket backend
function connectWebSocket() {
    const loc = window.location;
    const wsUri = (loc.protocol === "https:" ? "wss://" : "ws://") + loc.host + `/ws?session_id=${sessionId}`;
    
    addTerminalLine("[SİSTEM] WebSocket sunucusuna bağlanılıyor...", "system");
    socket = new WebSocket(wsUri);
    
    socket.onopen = () => {
        connectionDot.className = "status-dot online";
        connectionText.innerText = "Bağlı";
        btnStart.removeAttribute("disabled");
        btnParseQuery.removeAttribute("disabled");
        addTerminalLine("[SİSTEM] Sunucu bağlantısı sağlandı.", "success");
        
        // Start heartbeat ping
        if (heartbeatInterval) clearInterval(heartbeatInterval);
        heartbeatInterval = setInterval(() => {
            if (socket && socket.readyState === WebSocket.OPEN) {
                socket.send(JSON.stringify({ action: "ping" }));
            }
        }, 25000);
    };
    
    socket.onclose = () => {
        if (heartbeatInterval) {
            clearInterval(heartbeatInterval);
            heartbeatInterval = null;
        }
        connectionDot.className = "status-dot offline animate-pulse";
        connectionText.innerText = "Çevrimdışı";
        btnStart.setAttribute("disabled", "true");
        btnStop.setAttribute("disabled", "true");
        btnParseQuery.setAttribute("disabled", "true");
        addTerminalLine("[SİSTEM] Sunucu bağlantısı koptu. 5 saniye içinde tekrar denenecek...", "error");
        setTimeout(connectWebSocket, 5000);
    };
    
    socket.onerror = (err) => {
        addTerminalLine("[HATA] WebSocket hatası oluştu.", "error");
    };

    socket.onmessage = (event) => {
        const payload = JSON.parse(event.data);
        handleWebSocketMessage(payload);
    };
}

// Handle real-time WebSocket messages
function handleWebSocketMessage(msg) {
    switch (msg.type) {
        case "pong":
            // Heartbeat response, ignore
            break;
            
        case "init_state":
            // Load excel metadata and rows
            allRows = msg.data || [];
            activeHeaders = msg.headers || [];
            activeGcbColIdx = msg.gcb_col_idx || 9;
            activeDateColIdx = msg.date_col_idx || 12;
            activeFaturaColIdx = msg.fatura_col_idx || 1;
            activeFirmaColIdx = msg.firma_col_idx || 3;
            
            if (msg.active_file) {
                activeFileBadge.innerText = `Aktif Dosya: ${cleanFileName(msg.active_file)}`;
                activeFileBadge.className = "active-file-text loaded";
                btnDownload.classList.remove("disabled-btn");
            } else {
                activeFileBadge.innerText = "Aktif Dosya: Yok (Yeni Görev Bekleniyor)";
                activeFileBadge.className = "active-file-text empty";
                btnDownload.classList.add("disabled-btn");
            }
            filterTable();
            updateStats();
            
            // Re-populate terminal with log history
            if (msg.log_history && msg.log_history.length > 0) {
                terminal.innerHTML = "";
                msg.log_history.forEach(logText => {
                    let cls = "";
                    if (logText.includes("HATA") || logText.includes("başarısız")) {
                        cls = "error";
                    } else if (logText.includes("Bulundu") || logText.includes("başarıyla") || logText.includes("çözüldü") || logText.includes("Ayrıştırma başarılı")) {
                        cls = "success";
                    } else if (logText.includes("henüz kapanmamış") || logText.includes("uyarı")) {
                        cls = "warning";
                    }
                    addTerminalLine(logText, cls);
                });
            }
            
            // Re-sync automation button states and progress bar
            if (msg.is_running) {
                btnStart.setAttribute("disabled", "true");
                btnParseQuery.setAttribute("disabled", "true");
                btnStop.removeAttribute("disabled");
                
                const percent = Math.round((msg.completed / msg.total) * 100) || 0;
                progressBarFill.style.width = percent + "%";
                progressPercent.innerText = `${percent}% (${msg.completed}/${msg.total})`;
                startQueryTimer();
            } else {
                btnStart.removeAttribute("disabled");
                btnParseQuery.removeAttribute("disabled");
                btnStop.setAttribute("disabled", "true");
            }
            break;
            
        case "log":
            let logCls = "";
            const logText = msg.message;
            if (logText.includes("HATA") || logText.includes("başarısız")) {
                logCls = "error";
            } else if (logText.includes("Bulundu") || logText.includes("başarıyla") || logText.includes("çözüldü") || logText.includes("Ayrıştırma başarılı")) {
                logCls = "success";
            } else if (logText.includes("henüz kapanmamış") || logText.includes("uyarı")) {
                logCls = "warning";
            }
            addTerminalLine(logText, logCls);
            break;
            
        case "custom_list_loaded":
            // Reset queryStatus and reload the freshly loaded custom list rows
            queryStatus = {};
            allRows = msg.data;
            activeHeaders = msg.headers || [];
            activeGcbColIdx = msg.gcb_col_idx || 9;
            activeDateColIdx = msg.date_col_idx || 12;
            activeFaturaColIdx = msg.fatura_col_idx || 1;
            activeFirmaColIdx = msg.firma_col_idx || 3;
            
            if (msg.active_file) {
                activeFileBadge.innerText = `Aktif Dosya: ${cleanFileName(msg.active_file)}`;
                activeFileBadge.className = "active-file-text loaded";
                btnDownload.classList.remove("disabled-btn");
            } else {
                activeFileBadge.innerText = "Aktif Dosya: Yok (Yeni Görev Bekleniyor)";
                activeFileBadge.className = "active-file-text empty";
                btnDownload.classList.add("disabled-btn");
            }
            filterTable();
            updateStats();
            break;
            
        case "row_start":
            queryStatus[msg.row] = { intac: "", status: "Sorgulanıyor..." };
            addTerminalLine(`[SORGULAMA] Satır ${msg.row} için sorgu başlatıldı. GCB: ${msg.gcb}`, "system");
            startQueryTimer();
            updateRowUIStatus(msg.row, "Sorgulanıyor...", "badge-running", null, msg.gcb);
            updateStats();
            
            if (searchInput.value.trim() !== "" || document.getElementById("filter-status").value !== "all" || document.getElementById("filter-type").value !== "all") {
                filterTable();
            }
            break;
            
        case "row_success":
            queryStatus[msg.row] = { intac: msg.date, status: "İntaç Tarihi Var" };
            addTerminalLine(`[BAŞARILI] Satır ${msg.row} güncellendi: İntaç Tarihi = ${msg.date}`, "success");
            soundEngine.playSuccess();
            
            lastUpdatedRow = msg.row;
            lastUpdatedTime = Date.now();
            
            const cacheRowIdx = allRows.findIndex(r => Number(r.row) === Number(msg.row));
            if (cacheRowIdx !== -1) {
                allRows[cacheRowIdx].intac = msg.date;
                allRows[cacheRowIdx].status = "İntaç Tarihi Var";
            }
            updateRowUIStatus(msg.row, "İntaç Tarihi Var", "badge-success", msg.date, msg.gcb);
            updateStats();
            
            if (searchInput.value.trim() !== "" || document.getElementById("filter-status").value !== "all" || document.getElementById("filter-type").value !== "all") {
                filterTable();
            }
            break;
            
        case "row_not_closed":
            queryStatus[msg.row] = { intac: "", status: "Kapanmamış" };
            startCooldown(msg.gcb);
            addTerminalLine(`[UYARI] Satır ${msg.row} beyannamesi henüz kapanmamış. 5 dakika sorgu soğuma süresi başlatıldı.`, "warning");
            soundEngine.playNotClosed();
            
            lastUpdatedRow = msg.row;
            lastUpdatedTime = Date.now();
            
            const cacheRowIdxWarning = allRows.findIndex(r => Number(r.row) === Number(msg.row));
            if (cacheRowIdxWarning !== -1) {
                allRows[cacheRowIdxWarning].status = "Kapanmamış";
            }
            updateRowUIStatus(msg.row, "Kapanmamış", "badge-warning", null, msg.gcb);
            updateStats();
            
            if (searchInput.value.trim() !== "" || document.getElementById("filter-status").value !== "all" || document.getElementById("filter-type").value !== "all") {
                filterTable();
            }
            break;
            
        case "row_cooldown":
            queryStatus[msg.row] = { intac: "", status: "Soğumada" };
            startCooldown(msg.gcb);
            addTerminalLine(`[UYARI] Satır ${msg.row} sorgulama limitine takıldı. Cooldown süresi bitince otomatik tekrar denenecek.`, "warning");
            soundEngine.playNotClosed();
            
            lastUpdatedRow = msg.row;
            lastUpdatedTime = Date.now();
            
            const cacheRowIdxCooldown = allRows.findIndex(r => Number(r.row) === Number(msg.row));
            if (cacheRowIdxCooldown !== -1) {
                allRows[cacheRowIdxCooldown].status = "Soğumada";
            }
            updateRowUIStatus(msg.row, "Soğumada", "badge-cooldown", null, msg.gcb);
            updateStats();
            
            if (searchInput.value.trim() !== "" || document.getElementById("filter-status").value !== "all" || document.getElementById("filter-type").value !== "all") {
                filterTable();
            }
            break;
            
        case "row_fail":
            queryStatus[msg.row] = { intac: "", status: "Başarısız" };
            startCooldown(msg.gcb);
            addTerminalLine(`[BAŞARISIZ] Satır ${msg.row} sorgulama başarısız oldu: ${msg.message}. 5 dakika sorgu soğuma süresi başlatıldı.`, "error");
            soundEngine.playError();
            
            lastUpdatedRow = msg.row;
            lastUpdatedTime = Date.now();
            
            const cacheRowIdxFail = allRows.findIndex(r => Number(r.row) === Number(msg.row));
            if (cacheRowIdxFail !== -1) {
                allRows[cacheRowIdxFail].status = "Başarısız";
            }
            updateRowUIStatus(msg.row, "Başarısız", "badge-fail", null, msg.gcb);
            updateStats();
            
            if (searchInput.value.trim() !== "" || document.getElementById("filter-status").value !== "all" || document.getElementById("filter-type").value !== "all") {
                filterTable();
            }
            break;
            
        case "progress":
            const percent = Math.round((msg.completed / msg.total) * 100);
            progressBarFill.style.width = percent + "%";
            progressPercent.innerText = `${percent}% (${msg.completed}/${msg.total})`;
            break;
            
        case "finished":
            btnStart.removeAttribute("disabled");
            btnParseQuery.removeAttribute("disabled");
            btnStop.setAttribute("disabled", "true");
            addTerminalLine("[SİSTEM] Sorgulama işlemi tamamlandı. Güncel Excel dosyasını indirebilirsiniz.", "success");
            stopQueryTimer(true); // Stop timer and trigger modal report popup
            loadExcelData();
            break;
            
        case "stopped":
            btnStart.removeAttribute("disabled");
            btnParseQuery.removeAttribute("disabled");
            btnStop.setAttribute("disabled", "true");
            addTerminalLine("[SİSTEM] Sorgulama durduruldu. Kısmi sonuçlar Excel'e kaydedildi.", "warning");
            stopQueryTimer(true); // Stop timer and trigger modal report popup
            loadExcelData();
            break;
            
        case "error":
            addTerminalLine("HATA: " + msg.message, "error");
            btnStart.removeAttribute("disabled");
            btnParseQuery.removeAttribute("disabled");
            btnStop.setAttribute("disabled", "true");
            stopQueryTimer(false); // Stop timer silently
            break;
    }
}

// Add line to terminal panel with premium SVG icon styling
function addTerminalLine(text, className = "") {
    const now = new Date();
    const timeStr = now.toTimeString().split(" ")[0]; // Get HH:MM:SS format
    
    // Clean up duplicated prefixes if any
    let cleanText = text;
    let type = className;
    
    if (cleanText.startsWith("[SİSTEM]")) { cleanText = cleanText.replace("[SİSTEM]", "").trim(); type = "system"; }
    else if (cleanText.startsWith("[BAŞARILI]")) { cleanText = cleanText.replace("[BAŞARILI]", "").trim(); type = "success"; }
    else if (cleanText.startsWith("[SORGULAMA]")) { cleanText = cleanText.replace("[SORGULAMA]", "").trim(); type = "system"; }
    else if (cleanText.startsWith("[UYARI]")) { cleanText = cleanText.replace("[UYARI]", "").trim(); type = "warning"; }
    else if (cleanText.startsWith("[BAŞARISIZ]")) { cleanText = cleanText.replace("[BAŞARISIZ]", "").trim(); type = "error"; }
    else if (cleanText.startsWith("[HATA]")) { cleanText = cleanText.replace("[HATA]", "").trim(); type = "error"; }
    
    if (!type) type = "info";
    
    let iconSvg = "";
    if (type === "success") {
        iconSvg = `<svg class="log-icon log-icon-success" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width:15px; height:15px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>`;
    } else if (type === "error") {
        iconSvg = `<svg class="log-icon log-icon-error" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width:15px; height:15px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M12 9v3.75m9-.75a9 9 0 11-18 0 9 9 0 0118 0zm-9 3.75h.008v.008H12v-.008z" /></svg>`;
    } else if (type === "warning") {
        iconSvg = `<svg class="log-icon log-icon-warning" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width:15px; height:15px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126zM12 15.75h.007v.008H12v-.008z" /></svg>`;
    } else if (type === "system") {
        iconSvg = `<svg class="log-icon log-icon-system" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width:15px; height:15px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 17.25v1.007a3 3 0 01-.879 2.122L7.5 21h9l-.621-.621A3 3 0 0115 18.257V17.25m6-12V15a2.25 2.25 0 01-2.25 2.25H5.25A2.25 2.25 0 013 15V5.25A2.25 2.25 0 015.25 3h13.5A2.25 2.25 0 0121 5.25zM16.5 7.5h.008v.008H16.5V7.5z" /></svg>`;
    } else {
        iconSvg = `<svg class="log-icon log-icon-info" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" style="width:15px; height:15px;"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11.25 11.25l.041-.02a.75.75 0 111.063.852l-.708.283a.75.75 0 00-.475.694V15.75m0-6h.008v.008H12V9.75zM21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>`;
    }
    
    const line = document.createElement("div");
    line.className = "terminal-line " + type;
    line.innerHTML = `<span class="log-time">[${timeStr}]</span> <span class="log-icon-wrapper">${iconSvg}</span> <span class="log-text">${cleanText}</span>`;
    terminal.appendChild(line);
    terminal.scrollTop = terminal.scrollHeight;
}

// Trigger all automated queries
function startAutomation() {
    if (socket && socket.readyState === WebSocket.OPEN) {
        socket.send(JSON.stringify({ 
            action: "start_all"
        }));
        btnStart.setAttribute("disabled", "true");
        btnParseQuery.setAttribute("disabled", "true");
        btnStop.removeAttribute("disabled");
        progressBarFill.style.width = "0%";
        progressPercent.innerText = "0%";
        switchTab("tab-terminal"); // Switch to system logs tab automatically
        startQueryTimer();
    }
}

// Trigger custom GCB text list queries
function startCustomListAutomation() {
    const text = rawInput.value.trim();
    if (!text) {
        addTerminalLine("UYARI: Sorgulanacak metin listesi girilmemiş.", "warning");
        return;
    }
    
    if (socket && socket.readyState === WebSocket.OPEN) {
        socket.send(JSON.stringify({ 
            action: "start_custom_list",
            raw_text: text
        }));
        btnStart.setAttribute("disabled", "true");
        btnParseQuery.setAttribute("disabled", "true");
        btnStop.removeAttribute("disabled");
        progressBarFill.style.width = "0%";
        progressPercent.innerText = "0%";
        switchTab("tab-terminal"); // Switch to system logs tab automatically
        startQueryTimer();
    }
}

// Stop automation
function stopAutomation() {
    if (socket && socket.readyState === WebSocket.OPEN) {
        socket.send(JSON.stringify({ action: "stop" }));
    }
}

// Render data table
function renderTable(rows) {
    const tableHeaders = document.getElementById("table-headers");
    
    // Render dynamic headers
    if (activeHeaders && activeHeaders.length > 0) {
        let headerHtml = "<tr>";
        headerHtml += "<th>Satır</th>";
        activeHeaders.forEach(h => {
            headerHtml += `<th>${h}</th>`;
        });
        headerHtml += "<th>Durum</th>";
        headerHtml += "<th>Eylem</th>";
        headerHtml += "</tr>";
        tableHeaders.innerHTML = headerHtml;
    } else {
        // Default static headers
        tableHeaders.innerHTML = `
            <tr>
                <th>Satır</th>
                <th>Fatura No</th>
                <th>Firma Adı</th>
                <th>Beyanname (GCB) No</th>
                <th>Gümrük İntaç Tarihi</th>
                <th>Durum</th>
                <th>Eylem</th>
            </tr>
        `;
    }
    
    if (rows.length === 0) {
        const colSpan = (activeHeaders && activeHeaders.length > 0) ? (activeHeaders.length + 3) : 7;
        tableBody.innerHTML = `<tr><td colspan="${colSpan}" class="loading-cell">Sistem boşta. Yeni bir görev başlatmak için yukarıdan veri girişi yapın veya yeni bir Excel dosyası yükleyin.</td></tr>`;
        return;
    }
    
    let html = "";
    rows.forEach(item => {
        // Determine display state: prioritize actual intac data, then queryStatus, then item.status
        let state;
        if (item.intac && item.status === "İntaç Tarihi Var") {
            state = { intac: item.intac, status: "İntaç Tarihi Var" };
        } else if (queryStatus[item.row]) {
            state = queryStatus[item.row];
            // If queryStatus says intac found, also use it
            if (state.intac) {
                state = { intac: state.intac, status: "İntaç Tarihi Var" };
            }
        } else {
            state = { intac: item.intac, status: item.status };
        }
        
        let badgeClass = "badge-pending";
        if (state.status === "İntaç Tarihi Var") {
            badgeClass = "badge-success";
        } else if (state.status === "Kapanmamış") {
            badgeClass = "badge-warning";
        } else if (state.status === "Soğumada") {
            badgeClass = "badge-cooldown";
        } else if (state.status === "Başarısız" || state.status === "Hatalı") {
            badgeClass = "badge-fail";
        } else if (state.status === "Sorgulanıyor...") {
            badgeClass = "badge-running";
        }
        
        const isCompleted = state.status === "İntaç Tarihi Var";
        const cooldown = getCooldownRemaining(item.gcb);
        const gcbStr = (item.gcb || "").trim().toUpperCase();
        const isETGB = gcbStr.length === 16;
        let typeBadge = "";
        if (isETGB) {
            typeBadge = ' <span class="tag-etgb">ETGB</span>';
        } else if (gcbStr.length >= 18) {
            const typeCode = gcbStr.substring(8, 10);
            if (typeCode === "EX") {
                typeBadge = ' <span class="tag-ihracat">EX</span>';
            } else if (typeCode === "IM") {
                typeBadge = ' <span class="tag-ithalat">IM</span>';
            } else if (typeCode === "AN") {
                typeBadge = ' <span class="tag-antrepo">AN</span>';
            } else if (typeCode === "TR") {
                typeBadge = ' <span class="tag-transit">TR</span>';
            } else if (typeCode === "EU") {
                typeBadge = ' <span class="tag-transit-eu">EU</span>';
            }
        }
        
        let actionBtn = "";
        if (isCompleted) {
            actionBtn = `<button class="btn btn-inline btn-success" disabled><svg class="btn-icon-svg" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="3" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M4.5 12.75l6 6 9-13.5" /></svg> Tamam</button>`;
        } else if (state.status === "Sorgulanıyor...") {
            actionBtn = `<button class="btn btn-inline btn-primary" disabled>Sorgulanıyor...</button>`;
        } else if (cooldown > 0) {
            const seconds = Math.ceil(cooldown / 1000);
            const m = Math.floor(seconds / 60);
            const s = seconds % 60;
            const timeStr = `${m}:${s < 10 ? '0' : ''}${s}`;
            actionBtn = `<button class="btn btn-inline btn-secondary btn-cooldown" data-gcb="${item.gcb}" data-row="${item.row}" disabled>Sorgula (${timeStr})</button>`;
        } else {
            actionBtn = `<button class="btn btn-inline btn-primary" data-action-btn="${item.row}" onclick="querySingleRow(${item.row}, '${item.gcb}')">Sorgula</button>`;
        }
        
        let rowHtml = `<tr id="row-${item.row}">`;
        // 1. Row Index
        rowHtml += `<td class="font-mono">${item.row}</td>`;
        
        // 2. Dynamic values from spreadsheet
        if (item.values && item.values.length > 0) {
            item.values.forEach((val, colIdx) => {
                const isGcbCol = (colIdx + 1) === activeGcbColIdx;
                const isDateCol = (colIdx + 1) === activeDateColIdx;
                const isFaturaCol = (colIdx + 1) === activeFaturaColIdx;
                const isFirmaCol = (colIdx + 1) === activeFirmaColIdx;
                
                const headerName = (activeHeaders && activeHeaders[colIdx]) ? activeHeaders[colIdx].toLowerCase() : "";
                const isAnyDateCol = headerName.includes("tarih") || headerName.includes("date");
                
                if (isGcbCol) {
                    rowHtml += `<td class="font-mono"><strong>${val || "-"}</strong>${typeBadge}</td>`;
                } else if (isDateCol) {
                    rowHtml += `<td id="date-${item.row}" class="font-mono">${formatDate(state.intac || val) || "-"}</td>`;
                } else if (isAnyDateCol) {
                    rowHtml += `<td class="font-mono">${formatDate(val) || "-"}</td>`;
                } else if (isFaturaCol) {
                    rowHtml += `<td class="font-mono">${val || "-"}</td>`;
                } else if (isFirmaCol) {
                    rowHtml += `<td class="truncate" title="${val}">${val ? truncate(val, 35) : "-"}</td>`;
                } else {
                    rowHtml += `<td>${val || "-"}</td>`;
                }
            });
        } else {
            // Fallback to static columns if row values array is missing
            rowHtml += `<td class="font-mono">${item.fatura || "-"}</td>`;
            rowHtml += `<td title="${item.firma}">${item.firma ? truncate(item.firma, 35) : "-"}</td>`;
            rowHtml += `<td class="font-mono"><strong>${item.gcb || "-"}</strong>${typeBadge}</td>`;
            rowHtml += `<td id="date-${item.row}" class="font-mono">${formatDate(state.intac) || "-"}</td>`;
        }
        
        // 3. Status and Action buttons
        rowHtml += `<td><span id="badge-${item.row}" class="badge ${badgeClass}">${state.status}</span></td>`;
        rowHtml += `<td id="action-cell-${item.row}">${item.gcb ? actionBtn : "-"}</td>`;
        rowHtml += `</tr>`;
        
        html += rowHtml;
    });
    tableBody.innerHTML = html;
}

// Query single row manually
function querySingleRow(row, gcb) {
    if (socket && socket.readyState === WebSocket.OPEN) {
        socket.send(JSON.stringify({
            action: "query_single",
            row: row,
            gcb: gcb
        }));
        switchTab("tab-terminal"); // Switch to system logs tab automatically
    }
}

// Update single row cells in real-time
function updateRowUIStatus(row, statusText, badgeClass, date = null, gcb = null) {
    const badge = document.getElementById(`badge-${row}`);
    if (badge) {
        badge.className = "badge " + badgeClass;
        badge.innerText = statusText;
    }
    
    if (date !== null) {
        const dateCell = document.getElementById(`date-${row}`);
        if (dateCell) {
            dateCell.innerText = formatDate(date);
        }
    }
    
    const actionCell = document.getElementById(`action-cell-${row}`);
    if (actionCell && gcb) {
        if (statusText === "İntaç Tarihi Var") {
            actionCell.innerHTML = `<button class="btn btn-inline btn-success" disabled><svg class="btn-icon-svg" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="3" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M4.5 12.75l6 6 9-13.5" /></svg> Tamam</button>`;
        } else if (statusText === "Sorgulanıyor...") {
            actionCell.innerHTML = `<button class="btn btn-inline btn-primary" disabled>Sorgulanıyor...</button>`;
        } else {
            const cooldown = getCooldownRemaining(gcb);
            if (cooldown > 0) {
                const seconds = Math.ceil(cooldown / 1000);
                const m = Math.floor(seconds / 60);
                const s = seconds % 60;
                const timeStr = `${m}:${s < 10 ? '0' : ''}${s}`;
                actionCell.innerHTML = `<button class="btn btn-inline btn-secondary btn-cooldown" data-gcb="${gcb}" data-row="${row}" disabled>Sorgula (${timeStr})</button>`;
            } else {
                actionCell.innerHTML = `<button class="btn btn-inline btn-primary" data-action-btn="${row}" onclick="querySingleRow(${row}, '${gcb}')">Sorgula</button>`;
            }
        }
    }
    const tr = document.getElementById(`row-${row}`);
    if (tr) {
        tr.classList.remove("row-updating");
        void tr.offsetWidth; // force reflow
        tr.classList.add("row-updating");
    }
}

// Setup Excel drag & drop upload
function setupFileUpload() {
    // Click trigger
    dropZone.addEventListener("click", () => fileInput.click());
    
    fileInput.addEventListener("change", (e) => {
        if (fileInput.files.length > 0) {
            uploadExcelFile(fileInput.files[0]);
        }
    });
    
    // Drag/drop triggers
    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    });
    
    dropZone.addEventListener("dragleave", () => {
        dropZone.classList.remove("dragover");
    });
    
    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        if (e.dataTransfer.files.length > 0) {
            uploadExcelFile(e.dataTransfer.files[0]);
        }
    });
}

// Drop-zone text status helper
function setDropZoneStatus(text, isError = false) {
    const textEl = dropZone.querySelector(".drop-zone-text");
    if (textEl) {
        textEl.innerText = text;
        if (isError) {
            dropZone.style.borderColor = "var(--red-border)";
            textEl.style.color = "var(--red-text)";
        } else {
            dropZone.style.borderColor = "var(--border-medium)";
            textEl.style.color = "var(--text-secondary)";
        }
    }
}

// Perform file upload POST request
async function uploadExcelFile(file) {
    const ext = file.name.toLowerCase();
    if (!ext.endsWith(".xlsx")) {
        alert("HATA: Sadece .xlsx uzantılı modern Excel dosyaları desteklenmektedir.\n\nEğer dosyanız eski .xls formatındaysa, lütfen Microsoft Excel ile açıp sol üstten 'Dosya -> Farklı Kaydet' seçeneğini seçin. Dosya türünü 'Excel Çalışma Kitabı (*.xlsx)' olarak değiştirip kaydedin ve yeni .xlsx dosyasını buraya yükleyin.");
        addTerminalLine("HATA: .xls formatı desteklenmemektedir. Lütfen .xlsx formatına çevirip tekrar deneyin.", "error");
        fileInput.value = "";
        return;
    }
    
    addTerminalLine(`[SİSTEM] Excel dosyası yükleniyor: ${file.name}...`, "system");
    setDropZoneStatus(`YÜKLENİYOR: ${file.name.toUpperCase()}...`, false);
    dropZone.classList.add("upload-loading");
    
    const formData = new FormData();
    formData.append("file", file);
    
    try {
        const res = await fetch(`/api/upload?session_id=${sessionId}`, {
            method: "POST",
            body: formData
        });
        const json = await res.json();
        
        dropZone.classList.remove("upload-loading");
        
        if (json.success) {
            // Reset queries session
            queryStatus = {};
            allRows = json.data;
            activeHeaders = json.headers || [];
            activeGcbColIdx = json.gcb_col_idx || 9;
            activeDateColIdx = json.date_col_idx || 12;
            activeFaturaColIdx = json.fatura_col_idx || 1;
            activeFirmaColIdx = json.firma_col_idx || 3;
            
            const activeFile = cleanFileName(json.active_file) || "EXPORT.XLSX";
            activeFileBadge.innerText = `Aktif Dosya: ${activeFile}`;
            activeFileBadge.className = "active-file-text loaded";
            btnDownload.classList.remove("disabled-btn");
            
            renderTable(allRows);
            updateStats();
            
            setDropZoneStatus("Yeni Dosyayı Sürükleyin veya Dosya Seçin", false);
            addTerminalLine(`[BAŞARILI] Excel dosyası '${activeFile}' yüklendi ve tablo güncellendi.`, "success");
            
            // Premium sound and visual effects
            soundEngine.playUploadSuccess();
            dropZone.classList.add("upload-success");
            setTimeout(() => dropZone.classList.remove("upload-success"), 1500);
        } else {
            setDropZoneStatus("YÜKLEME BAŞARISIZ! Tekrar Deneyin.", true);
            alert("Excel Yükleme Hatası:\n" + json.message);
            addTerminalLine("HATA: Excel yükleme başarısız: " + json.message, "error");
            
            soundEngine.playError();
            dropZone.classList.add("upload-fail");
            setTimeout(() => dropZone.classList.remove("upload-fail"), 1500);
        }
    } catch (e) {
        dropZone.classList.remove("upload-loading");
        setDropZoneStatus("SUNUCU BAĞLANTI HATASI! Tekrar Deneyin.", true);
        alert("Sunucu Bağlantı Hatası:\n" + e.message);
        addTerminalLine("HATA: Sunucuya bağlanırken hata oluştu: " + e.message, "error");
        
        soundEngine.playError();
        dropZone.classList.add("upload-fail");
        setTimeout(() => dropZone.classList.remove("upload-fail"), 1500);
    } finally {
        fileInput.value = ""; // Always reset file input value to allow re-upload of the same file
    }
}

// Helper: Truncate long strings
function truncate(str, n) {
    return (str.length > n) ? str.substr(0, n - 1) + '&hellip;' : str;
}

// Recalculate stats cards
function updateStats() {
    // Group rows by GCB
    const gcbGroups = {};
    allRows.forEach(item => {
        const gcb = item.gcb ? item.gcb.trim() : "";
        if (!gcb) return;
        if (!gcbGroups[gcb]) {
            gcbGroups[gcb] = [];
        }
        gcbGroups[gcb].push(item);
    });
    
    // Determine consolidated status for each unique GCB
    let uniqueTotal = 0;
    let uniqueCompleted = 0;
    let uniquePending = 0;
    let uniqueNotClosed = 0;
    let uniqueErrors = 0;
    let uniqueCooldown = 0;
    
    let rowTotal = allRows.length;
    let rowCompleted = 0;
    let rowPending = 0;
    let rowNotClosed = 0;
    let rowErrors = 0;
    let rowCooldown = 0;
    
    for (const gcb in gcbGroups) {
        uniqueTotal++;
        const rows = gcbGroups[gcb];
        const status = getGcbConsolidatedStatus(rows);
        
        if (status === "İntaç Tarihi Var") {
            uniqueCompleted++;
        } else if (status === "Kapanmamış") {
            uniqueNotClosed++;
        } else if (status === "Başarısız" || status === "Hatalı") {
            uniqueErrors++;
        } else if (status === "Soğumada") {
            uniqueCooldown++;
        } else {
            uniquePending++;
        }
        
        // Count rows for each category
        rows.forEach(item => {
            let rowStatus;
            const qs = queryStatus[item.row];
            if (item.intac && item.status === "İntaç Tarihi Var") {
                rowStatus = "İntaç Tarihi Var";
            } else if (qs) {
                rowStatus = qs.status;
            } else {
                rowStatus = item.status || "Bekliyor";
            }
            
            if (rowStatus === "İntaç Tarihi Var") {
                rowCompleted++;
            } else if (rowStatus === "Kapanmamış") {
                rowNotClosed++;
            } else if (rowStatus === "Başarısız" || rowStatus === "Hatalı") {
                rowErrors++;
            } else if (rowStatus === "Soğumada") {
                rowCooldown++;
            } else {
                rowPending++;
            }
        });
    }
    
    // Update the main counter and the subtitle text
    if (statTotal) animateCounter(statTotal, uniqueTotal);
    if (statCompleted) animateCounter(statCompleted, uniqueCompleted);
    if (statPending) animateCounter(statPending, uniquePending);
    if (statNotClosed) animateCounter(statNotClosed, uniqueNotClosed);
    if (statErrors) animateCounter(statErrors, uniqueErrors);
    if (statCooldown) animateCounter(statCooldown, uniqueCooldown);
    
    // Update subtitle elements
    updateSubtitle("stat-total-rows", rowTotal);
    updateSubtitle("stat-completed-rows", rowCompleted);
    updateSubtitle("stat-pending-rows", rowPending);
    updateSubtitle("stat-not-closed-rows", rowNotClosed);
    updateSubtitle("stat-errors-rows", rowErrors);
    updateSubtitle("stat-cooldown-rows", rowCooldown);
    
    if (allRows.length === 0) {
        btnStart.setAttribute("disabled", "true");
    } else if (socket && socket.readyState === WebSocket.OPEN) {
        btnStart.removeAttribute("disabled");
    }
}

function updateSubtitle(elementId, rowCount) {
    const el = document.getElementById(elementId);
    if (el) {
        el.innerText = `${rowCount} Satır`;
    }
}

// Timer and Analysis Modal Helpers
function startQueryTimer() {
    if (queryTimerInterval) return; // already running
    
    queryStartTime = Date.now();
    const timerLabel = document.getElementById("query-timer-label");
    if (timerLabel) {
        timerLabel.style.display = "block";
        timerLabel.innerHTML = `<svg class="timer-icon-svg" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2.2" stroke="currentColor" style="width:16px; height:16px; display:inline-block; vertical-align:middle; margin-right:4px;"><path stroke-linecap="round" stroke-linejoin="round" d="M12 6v6h4.5m4.5 0a9 9 0 11-18 0 9 9 0 0118 0z" /></svg> Geçen Süre: <strong>00:00</strong> (Kalan: <strong>Hesaplanıyor...</strong>)`;
    }
    
    queryTimerInterval = setInterval(updateQueryTimerUI, 1000);
}

function updateQueryTimerUI() {
    if (!queryStartTime) return;
    
    const elapsedMs = Date.now() - queryStartTime;
    const elapsedSeconds = Math.floor(elapsedMs / 1000);
    
    // Calculate unique GCB progress
    const allUniqueGcbs = [...new Set(allRows.map(item => item.gcb).filter(Boolean))];
    const totalUnique = allUniqueGcbs.length;
    
    let finishedUnique = 0;
    allUniqueGcbs.forEach(gcb => {
        const rows = allRows.filter(r => r.gcb === gcb);
        const status = getGcbConsolidatedStatus(rows);
        if (status !== "Bekliyor" && status !== "Sorgulanıyor...") {
            finishedUnique++;
        }
    });
    
    let elapsedStr = formatSeconds(elapsedSeconds);
    let remainingStr = "--:--";
    
    if (finishedUnique > 0 && totalUnique > finishedUnique) {
        const avgTimePerGcb = elapsedSeconds / finishedUnique;
        const remainingGcbs = totalUnique - finishedUnique;
        const etaSeconds = Math.round(remainingGcbs * avgTimePerGcb);
        remainingStr = formatSeconds(etaSeconds);
    } else if (totalUnique === finishedUnique && totalUnique > 0) {
        remainingStr = "00:00";
    } else if (finishedUnique === 0 && totalUnique > 0) {
        const etaSeconds = totalUnique * 12; // Fallback estimate
        remainingStr = formatSeconds(etaSeconds);
    }
    
    const timerLabel = document.getElementById("query-timer-label");
    if (timerLabel) {
        timerLabel.innerHTML = `<svg class="timer-icon-svg" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2.2" stroke="currentColor" style="width:16px; height:16px; display:inline-block; vertical-align:middle; margin-right:4px;"><path stroke-linecap="round" stroke-linejoin="round" d="M12 6v6h4.5m4.5 0a9 9 0 11-18 0 9 9 0 0118 0z" /></svg> Geçen Süre: <strong>${elapsedStr}</strong> (Kalan: <strong>${remainingStr}</strong>)`;
    }
}

function stopQueryTimer(showAnalysis = false) {
    if (queryTimerInterval) {
        clearInterval(queryTimerInterval);
        queryTimerInterval = null;
    }
    
    const elapsedMs = queryStartTime ? (Date.now() - queryStartTime) : 0;
    const elapsedSeconds = Math.floor(elapsedMs / 1000);
    
    if (showAnalysis && elapsedSeconds > 1) {
        showQueryAnalysisModal(elapsedSeconds);
    }
    
    queryStartTime = null;
}

function showQueryAnalysisModal(elapsedSeconds) {
    const allUniqueGcbs = [...new Set(allRows.map(item => item.gcb).filter(Boolean))];
    const totalUnique = allUniqueGcbs.length;
    
    let finishedUnique = 0;
    let successUnique = 0;
    let notClosedUnique = 0;
    
    allUniqueGcbs.forEach(gcb => {
        const rows = allRows.filter(r => r.gcb === gcb);
        const status = getGcbConsolidatedStatus(rows);
        if (status !== "Bekliyor" && status !== "Sorgulanıyor...") {
            finishedUnique++;
            if (status === "İntaç Tarihi Var") successUnique++;
            if (status === "Kapanmamış") notClosedUnique++;
        }
    });
    
    const avgTime = finishedUnique > 0 ? (elapsedSeconds / finishedUnique).toFixed(1) + " sn" : "0 sn";
    // Success rate is based on successfully resolved queries (both Closed and Not Closed are successful portal results)
    const totalSuccessfulResolved = successUnique + notClosedUnique;
    const successRate = finishedUnique > 0 ? Math.round((totalSuccessfulResolved / finishedUnique) * 100) : 0;
    
    const totalTimeEl = document.getElementById("analysis-total-time");
    const avgTimeEl = document.getElementById("analysis-avg-time");
    const successCountEl = document.getElementById("analysis-success-count");
    const notClosedEl = document.getElementById("analysis-not-closed");
    const efficiencyBar = document.getElementById("analysis-efficiency-bar");
    const efficiencyText = document.getElementById("analysis-efficiency-text");
    
    if (totalTimeEl) totalTimeEl.innerText = formatSeconds(elapsedSeconds);
    if (avgTimeEl) avgTimeEl.innerText = avgTime;
    if (successCountEl) successCountEl.innerText = successUnique;
    if (notClosedEl) notClosedEl.innerText = notClosedUnique;
    if (efficiencyBar) efficiencyBar.style.width = successRate + "%";
    if (efficiencyText) efficiencyText.innerText = `%${successRate} Başarı Oranı (${totalSuccessfulResolved} / ${finishedUnique} GCB)`;
    
    const modal = document.getElementById("analysis-modal");
    if (modal) {
        modal.style.display = "flex";
        // Trigger reflow for transition
        void modal.offsetWidth;
        modal.classList.add("fade-in");
    }
}

function closeAnalysisModal() {
    const modal = document.getElementById("analysis-modal");
    if (modal) {
        modal.classList.remove("fade-in");
        setTimeout(() => {
            modal.style.display = "none";
        }, 300);
    }
}

function formatSeconds(totalSeconds) {
    const m = Math.floor(totalSeconds / 60);
    const s = totalSeconds % 60;
    return `${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
}

window.closeAnalysisModal = closeAnalysisModal;


// Animated numerical count-up helper
// Uses element._currentValue to track the real numeric value instead of parsing textContent,
// which prevents corrupted/negative numbers from race conditions during rapid WebSocket updates.
function animateCounter(element, targetValue) {
    // Ensure targetValue is a valid non-negative integer
    targetValue = Math.max(0, Math.round(targetValue)) || 0;
    
    // Cancel any ongoing animation for this element
    if (element._animId) {
        cancelAnimationFrame(element._animId);
        element._animId = null;
    }
    
    const previousValue = element._currentValue || 0;
    
    // Set value IMMEDIATELY (no animation delay) for perfect sync
    element.textContent = targetValue;
    element._currentValue = targetValue;
    
    // Pulse card animation (only if value actually changed)
    if (previousValue !== targetValue) {
        const card = element.closest(".stat-card");
        if (card) {
            card.classList.remove("stat-card-pulse");
            void card.offsetWidth; // force layout reflow
            card.classList.add("stat-card-pulse");
        }
    }
}

// Table search and column filter combined
function filterTable() {
    const query = searchInput.value.toLowerCase().trim();
    const statusFilter = document.getElementById("filter-status").value;
    const typeFilter = document.getElementById("filter-type").value;
    
    const filtered = allRows.filter(item => {
        // 1. Text Search Filter
        const matchesQuery = !query || (
            (item.fatura && item.fatura.toLowerCase().includes(query)) ||
            (item.firma && item.firma.toLowerCase().includes(query)) ||
            (item.gcb && item.gcb.toLowerCase().includes(query))
        );
        
        // 2. Status Filter — use same priority as updateStats
        let state;
        if (item.intac && item.status === "İntaç Tarihi Var") {
            state = { intac: item.intac, status: "İntaç Tarihi Var" };
        } else if (queryStatus[item.row]) {
            state = queryStatus[item.row];
        } else {
            state = { intac: item.intac, status: item.status };
        }
        const matchesStatus = statusFilter === "all" || state.status === statusFilter;
        
        // 3. Type (İhracat / İthalat / Antrepo / Transit / Ortak Transit / ETGB) Filter
        const gcbStr = (item.gcb || "").trim().toUpperCase();
        const isETGB = gcbStr.length === 16;
        const typeCode = (!isETGB && gcbStr.length >= 18) ? gcbStr.substring(8, 10) : "";
        const isIthalat = typeCode === "IM";
        const isAntrepo = typeCode === "AN";
        const isTransit = typeCode === "TR";
        const isTransitEU = typeCode === "EU";
        const isIhracat = typeCode === "EX" || (!isETGB && !isIthalat && !isAntrepo && !isTransit && !isTransitEU);
        const matchesType = typeFilter === "all" || (
            (typeFilter === "etgb" && isETGB) ||
            (typeFilter === "ithalat" && isIthalat) ||
            (typeFilter === "antrepo" && isAntrepo) ||
            (typeFilter === "transit" && isTransit) ||
            (typeFilter === "transit_eu" && isTransitEU) ||
            (typeFilter === "ihracat" && isIhracat)
        );
        
        return matchesQuery && matchesStatus && matchesType;
    });
    
    renderTable(filtered);
}

// Reset table action
function resetExcelTable() {
    if (socket && socket.readyState === WebSocket.OPEN) {
        socket.send(JSON.stringify({ action: "reset_excel" }));
        rawInput.value = ""; // clear GCB text input
        progressBarFill.style.width = "0%";
        progressPercent.innerText = "0%";
        
        // Reset stat card animation state
        [statTotal, statCompleted, statPending, statNotClosed, statErrors, statCooldown].forEach(el => {
            if (el) {
                el._currentValue = 0;
                el.textContent = "0";
                if (el._animId) {
                    cancelAnimationFrame(el._animId);
                    el._animId = null;
                }
            }
        });
        
        soundEngine.playReset();
    }
}

// Date formatter helper (converts YYYY-MM-DD or other formats to Turkish DD.MM.YYYY format)
function formatDate(dateStr) {
    if (!dateStr) return "";
    dateStr = String(dateStr).trim();
    // Match YYYY-MM-DD optionally followed by time
    const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})(?:\s.+)?$/);
    if (match) {
        const year = match[1];
        const month = match[2];
        const day = match[3]; // Keep leading zeros!
        return `${day}.${month}.${year}`;
    }
    // If it's already in D.MM.YYYY or DD.MM.YYYY format
    if (dateStr.includes(".")) {
        const parts = dateStr.split(".");
        if (parts.length === 3) {
            const day = parts[0].padStart(2, '0'); // pad day with leading zero
            const month = parts[1].padStart(2, '0'); // pad month with leading zero
            const year = parts[2].split(" ")[0]; // remove time if any
            return `${day}.${month}.${year}`;
        }
    }
    return dateStr;
}


// Cooldown state managers
function startCooldown(gcb) {
    if (!gcb) return;
    const cooldowns = JSON.parse(localStorage.getItem("gcb_cooldowns") || "{}");
    cooldowns[gcb] = Date.now();
    localStorage.setItem("gcb_cooldowns", JSON.stringify(cooldowns));
}

function getCooldownRemaining(gcb) {
    if (!gcb) return 0;
    const cooldowns = JSON.parse(localStorage.getItem("gcb_cooldowns") || "{}");
    const timestamp = cooldowns[gcb];
    if (!timestamp) return 0;
    const elapsed = Date.now() - timestamp;
    const remaining = Math.max(0, 300000 - elapsed); // 5 minutes (300,000ms)
    return remaining;
}

function updateCooldowns() {
    const buttons = document.querySelectorAll(".btn-cooldown");
    buttons.forEach(btn => {
        const gcb = btn.dataset.gcb;
        const row = btn.dataset.row;
        const cooldown = getCooldownRemaining(gcb);
        if (cooldown > 0) {
            const seconds = Math.ceil(cooldown / 1000);
            const m = Math.floor(seconds / 60);
            const s = seconds % 60;
            const timeStr = `${m}:${s < 10 ? '0' : ''}${s}`;
            btn.innerText = `Sorgula (${timeStr})`;
        } else {
            btn.className = "btn btn-inline btn-primary";
            btn.removeAttribute("disabled");
            btn.innerText = "Sorgula";
            btn.onclick = () => querySingleRow(row, gcb);
        }
    });
}
