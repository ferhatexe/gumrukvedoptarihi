let socket = null;
let heartbeatInterval = null;
let allRows = [];
let queryStatus = {}; // Keep track of current session query status per row index
let activeHeaders = []; // Dynamic headers from active Excel
let activeGcbColIdx = 9; // Index of GCB column (1-based)
let activeDateColIdx = 12; // Index of Date column (1-based)
let activeFaturaColIdx = 1; // Index of Fatura column (1-based)
let activeFirmaColIdx = 3; // Index of Firma column (1-based)

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

// DOM Elements
const connectionDot = document.getElementById("connection-status-dot");
const connectionText = document.getElementById("connection-status-text");
const activeFileBadge = document.getElementById("active-file-badge");
const statTotal = document.getElementById("stat-total");
const statCompleted = document.getElementById("stat-completed");
const statPending = document.getElementById("stat-pending");
const statErrors = document.getElementById("stat-errors");
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
    document.getElementById("filter-status").addEventListener("change", filterTable);
    document.getElementById("filter-type").addEventListener("change", filterTable);
    
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
                activeFileBadge.innerText = `Aktif Dosya: ${json.active_file}`;
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
                activeFileBadge.innerText = `Aktif Dosya: ${msg.active_file}`;
                activeFileBadge.className = "active-file-text loaded";
                btnDownload.classList.remove("disabled-btn");
            } else {
                activeFileBadge.innerText = "Aktif Dosya: Yok (Yeni Görev Bekleniyor)";
                activeFileBadge.className = "active-file-text empty";
                btnDownload.classList.add("disabled-btn");
            }
            renderTable(allRows);
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
            } else {
                btnStart.removeAttribute("disabled");
                btnParseQuery.removeAttribute("disabled");
                btnStop.setAttribute("disabled", "true");
            }
            break;
            
        case "log":
            let cls = "";
            const text = msg.message;
            if (text.includes("HATA") || text.includes("başarısız")) {
                cls = "error";
            } else if (text.includes("Bulundu") || text.includes("başarıyla") || text.includes("çözüldü") || text.includes("Ayrıştırma başarılı")) {
                cls = "success";
            } else if (text.includes("henüz kapanmamış") || text.includes("uyarı")) {
                cls = "warning";
            }
            addTerminalLine(text, cls);
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
                activeFileBadge.innerText = `Aktif Dosya: ${msg.active_file}`;
                activeFileBadge.className = "active-file-text loaded";
                btnDownload.classList.remove("disabled-btn");
            } else {
                activeFileBadge.innerText = "Aktif Dosya: Yok (Yeni Görev Bekleniyor)";
                activeFileBadge.className = "active-file-text empty";
                btnDownload.classList.add("disabled-btn");
            }
            renderTable(allRows);
            updateStats();
            break;
            
        case "row_start":
            updateRowUIStatus(msg.row, "Sorgulanıyor...", "badge-running", null, msg.gcb);
            addTerminalLine(`[SORGULAMA] Satır ${msg.row} için sorgu başlatıldı. GCB: ${msg.gcb}`, "system");
            break;
            
        case "row_success":
            queryStatus[msg.row] = { intac: msg.date, status: "İntaç Tarihi Var" };
            updateRowUIStatus(msg.row, "İntaç Tarihi Var", "badge-success", msg.date, msg.gcb);
            addTerminalLine(`[BAŞARILI] Satır ${msg.row} güncellendi: İntaç Tarihi = ${msg.date}`, "success");
            
            const tr = document.getElementById(`row-${msg.row}`);
            if (tr) {
                tr.classList.add("row-updating");
                setTimeout(() => tr.classList.remove("row-updating"), 2000);
            }
            const cacheRowIdx = allRows.findIndex(r => r.row === msg.row);
            if (cacheRowIdx !== -1) {
                allRows[cacheRowIdx].intac = msg.date;
                allRows[cacheRowIdx].status = "İntaç Tarihi Var";
            }
            updateStats();
            break;
            
        case "row_not_closed":
            queryStatus[msg.row] = { intac: "", status: "Kapanmamış" };
            startCooldown(msg.gcb);
            updateRowUIStatus(msg.row, "Kapanmamış", "badge-fail", null, msg.gcb);
            addTerminalLine(`[UYARI] Satır ${msg.row} beyannamesi henüz kapanmamış. 5 dakika sorgu soğuma süresi başlatıldı.`, "warning");
            
            const cacheRowIdxWarning = allRows.findIndex(r => r.row === msg.row);
            if (cacheRowIdxWarning !== -1) {
                allRows[cacheRowIdxWarning].status = "Kapanmamış";
            }
            updateStats();
            break;
            
        case "row_fail":
            queryStatus[msg.row] = { intac: "", status: "Başarısız" };
            startCooldown(msg.gcb);
            updateRowUIStatus(msg.row, "Başarısız", "badge-fail", null, msg.gcb);
            addTerminalLine(`[BAŞARISIZ] Satır ${msg.row} sorgulama başarısız oldu: ${msg.message}. 5 dakika sorgu soğuma süresi başlatıldı.`, "error");
            
            const cacheRowIdxFail = allRows.findIndex(r => r.row === msg.row);
            if (cacheRowIdxFail !== -1) {
                allRows[cacheRowIdxFail].status = "Başarısız";
            }
            updateStats();
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
            loadExcelData();
            break;
            
        case "stopped":
            btnStart.removeAttribute("disabled");
            btnParseQuery.removeAttribute("disabled");
            btnStop.setAttribute("disabled", "true");
            addTerminalLine("[SİSTEM] Sorgulama durduruldu. Kısmi sonuçlar Excel'e kaydedildi.", "warning");
            loadExcelData();
            break;
            
        case "error":
            addTerminalLine("HATA: " + msg.message, "error");
            btnStart.removeAttribute("disabled");
            btnParseQuery.removeAttribute("disabled");
            btnStop.setAttribute("disabled", "true");
            break;
    }
}

// Add line to terminal panel
function addTerminalLine(text, className = "") {
    const line = document.createElement("div");
    line.className = "terminal-line " + className;
    line.innerText = text;
    terminal.appendChild(line);
    terminal.scrollTop = terminal.scrollHeight;
}

// Trigger all automated queries
function startAutomation() {
    if (socket && socket.readyState === WebSocket.OPEN) {
        socket.send(JSON.stringify({ action: "start_all" }));
        btnStart.setAttribute("disabled", "true");
        btnParseQuery.setAttribute("disabled", "true");
        btnStop.removeAttribute("disabled");
        progressBarFill.style.width = "0%";
        progressPercent.innerText = "0%";
        switchTab("tab-terminal"); // Switch to system logs tab automatically
    }
}

// Trigger custom GCB text list queries
function startCustomListAutomation() {
    const text = rawInput.value.strip();
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
        const state = queryStatus[item.row] || { intac: item.intac, status: item.status };
        
        let badgeClass = "badge-pending";
        if (state.status === "İntaç Tarihi Var") {
            badgeClass = "badge-success";
        } else if (state.status === "Kapanmamış" || state.status === "Başarısız") {
            badgeClass = "badge-fail";
        } else if (state.status === "Sorgulanıyor...") {
            badgeClass = "badge-running";
        }
        
        const isCompleted = state.status === "İntaç Tarihi Var";
        const cooldown = getCooldownRemaining(item.gcb);
        const isETGB = item.gcb && item.gcb.trim().length === 16;
        const etgbBadge = isETGB ? ' <span class="tag-etgb">ETGB</span>' : '';
        
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
                
                if (isGcbCol) {
                    rowHtml += `<td class="font-mono"><strong>${val || "-"}</strong>${etgbBadge}</td>`;
                } else if (isDateCol) {
                    rowHtml += `<td id="date-${item.row}" class="font-mono">${formatDate(state.intac || val) || "-"}</td>`;
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
            rowHtml += `<td class="font-mono"><strong>${item.gcb || "-"}</strong>${etgbBadge}</td>`;
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
    if (!ext.endsWith(".xlsx") && !ext.endsWith(".xls")) {
        alert("HATA: Lütfen geçerli bir Excel dosyası (.xlsx veya .xls) yükleyin.");
        addTerminalLine("HATA: Lütfen geçerli bir Excel dosyası (.xlsx veya .xls) yükleyin.", "error");
        return;
    }
    
    addTerminalLine(`[SİSTEM] Excel dosyası yükleniyor: ${file.name}...`, "system");
    setDropZoneStatus(`YÜKLENİYOR: ${file.name.toUpperCase()}...`, false);
    
    const formData = new FormData();
    formData.append("file", file);
    
    try {
        const res = await fetch(`/api/upload?session_id=${sessionId}`, {
            method: "POST",
            body: formData
        });
        const json = await res.json();
        
        if (json.success) {
            // Reset queries session
            queryStatus = {};
            allRows = json.data;
            activeHeaders = json.headers || [];
            activeGcbColIdx = json.gcb_col_idx || 9;
            activeDateColIdx = json.date_col_idx || 12;
            activeFaturaColIdx = json.fatura_col_idx || 1;
            activeFirmaColIdx = json.firma_col_idx || 3;
            
            const activeFile = json.active_file || "EXPORT.XLSX";
            activeFileBadge.innerText = `Aktif Dosya: ${activeFile}`;
            activeFileBadge.className = "active-file-text loaded";
            btnDownload.classList.remove("disabled-btn");
            
            renderTable(allRows);
            updateStats();
            
            setDropZoneStatus("Yeni Dosyayı Sürükleyin veya Dosya Seçin", false);
            addTerminalLine(`[BAŞARILI] Excel dosyası '${activeFile}' yüklendi ve tablo güncellendi.`, "success");
        } else {
            setDropZoneStatus("YÜKLEME BAŞARISIZ! Tekrar Deneyin.", true);
            alert("Excel Yükleme Hatası:\n" + json.message);
            addTerminalLine("HATA: Excel yükleme başarısız: " + json.message, "error");
        }
    } catch (e) {
        setDropZoneStatus("SUNUCU BAĞLANTI HATASI! Tekrar Deneyin.", true);
        alert("Sunucu Bağlantı Hatası:\n" + e.message);
        addTerminalLine("HATA: Sunucuya bağlanırken hata oluştu: " + e.message, "error");
    }
}

// Helper: Truncate long strings
function truncate(str, n) {
    return (str.length > n) ? str.substr(0, n - 1) + '&hellip;' : str;
}

// Recalculate stats cards
function updateStats() {
    statTotal.innerText = allRows.length;
    
    let completed = 0;
    let pending = 0;
    let errors = 0;
    
    allRows.forEach(item => {
        const state = queryStatus[item.row] || { intac: item.intac, status: item.status };
        if (state.status === "İntaç Tarihi Var") {
            completed++;
        } else if (state.status === "Bekliyor" || state.status === "Sorgulanıyor...") {
            pending++;
        } else if (state.status === "Kapanmamış" || state.status === "Başarısız" || state.status === "Sistem Uyarısı") {
            errors++;
        }
    });
    
    statCompleted.innerText = completed;
    statPending.innerText = pending;
    statErrors.innerText = errors;
    
    if (allRows.length === 0) {
        btnStart.setAttribute("disabled", "true");
    } else if (socket && socket.readyState === WebSocket.OPEN) {
        btnStart.removeAttribute("disabled");
    }
}

// Table search and column filter combined
function filterTable() {
    const query = searchInput.value.toLowerCase().strip();
    const statusFilter = document.getElementById("filter-status").value;
    const typeFilter = document.getElementById("filter-type").value;
    
    const filtered = allRows.filter(item => {
        // 1. Text Search Filter
        const matchesQuery = !query || (
            (item.fatura && item.fatura.toLowerCase().includes(query)) ||
            (item.firma && item.firma.toLowerCase().includes(query)) ||
            (item.gcb && item.gcb.toLowerCase().includes(query))
        );
        
        // 2. Status Filter
        const state = queryStatus[item.row] || { intac: item.intac, status: item.status };
        const matchesStatus = statusFilter === "all" || state.status === statusFilter;
        
        // 3. Type (ETGB vs Beyanname) Filter
        const isETGB = item.gcb && item.gcb.trim().length === 16;
        const matchesType = typeFilter === "all" || (
            (typeFilter === "etgb" && isETGB) ||
            (typeFilter === "beyanname" && !isETGB)
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
    }
}

// Date formatter helper (converts YYYY-MM-DD to D.MM.YYYY e.g. 6.06.2026)
function formatDate(dateStr) {
    if (!dateStr) return "";
    dateStr = String(dateStr).trim();
    // Match YYYY-MM-DD optionally followed by time
    const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})(?:\s.+)?$/);
    if (match) {
        const year = match[1];
        const month = match[2];
        const day = parseInt(match[3], 10); // strip leading zeros, e.g. 06 -> 6
        return `${day}.${month}.${year}`;
    }
    // If it's already in D.MM.YYYY or DD.MM.YYYY format
    if (dateStr.includes(".")) {
        const parts = dateStr.split(".");
        if (parts.length === 3) {
            const day = parseInt(parts[0], 10); // strip leading zeros
            const month = parts[1];
            const year = parts[2].split(" ")[0]; // remove time if any
            return `${day}.${month}.${year}`;
        }
    }
    return dateStr;
}

// Clean prototype helper
String.prototype.strip = function() {
    return this.replace(/^\s+|\s+$/g, '');
};

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
