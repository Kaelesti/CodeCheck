// Register Service Worker for PWA
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js')
      .then(reg => console.log('Service Worker registrado con éxito', reg))
      .catch(err => console.error('Error al registrar Service Worker', err));
  });
}

// Global Application State
const state = {
  excelData: [],
  excelHeaders: [],
  scannedCodes: new Set(),
  scanHistory: [],
  html5QrcodeScanner: null,
  activeTab: 'audit'
};

// DOM Elements
document.addEventListener('DOMContentLoaded', () => {
  initTabs();
  initExcelUpload();
  initManualEntry();
  initExport();
});

// Tab Switching Logic
function initTabs() {
  const tabBtns = document.querySelectorAll('.tab-btn');
  const tabContents = document.querySelectorAll('.tab-content');

  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetTab = btn.dataset.tab;
      
      tabBtns.forEach(b => {
        b.classList.remove('active', 'border-sky-600', 'text-sky-600');
        b.classList.add('border-transparent', 'text-slate-500');
      });
      
      btn.classList.add('active', 'border-sky-600', 'text-sky-600');
      btn.classList.remove('border-transparent', 'text-slate-500');

      tabContents.forEach(content => {
        if (content.id === `${targetTab}-tab`) {
          content.classList.remove('hidden');
        } else {
          content.classList.add('hidden');
        }
      });

      state.activeTab = targetTab;

      if (targetTab === 'scanner') {
        startScanner();
      } else {
        stopScanner();
      }
    });
  });
}

// Excel File Upload Handler
function initExcelUpload() {
  const excelInput = document.getElementById('excel-file-input');
  if (!excelInput) return;

  excelInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      if (json.length > 0) {
        state.excelHeaders = json[0];
        state.excelData = json.slice(1).filter(row => row.length > 0);
        renderExcelTable();
        updateSummaryMetrics();
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

// Render Loaded Excel Table
function renderExcelTable() {
  const tbody = document.getElementById('excel-table-body');
  if (!tbody) return;

  tbody.innerHTML = '';

  state.excelData.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50 border-b border-slate-100 text-xs';
    
    const codeValue = String(row[0] || '').trim();
    const isScanned = state.scannedCodes.has(codeValue);
    
    let statusBadge = isScanned 
      ? '<span class="px-2 py-0.5 rounded-full text-xs font-medium bg-emerald-100 text-emerald-800">Verificado</span>'
      : '<span class="px-2 py-0.5 rounded-full text-xs font-medium bg-amber-100 text-amber-800">Pendiente</span>';

    tr.innerHTML = `
      <td class="px-3 py-2 text-slate-500 font-mono">${index + 1}</td>
      <td class="px-3 py-2 font-medium text-slate-900 font-mono">${codeValue}</td>
      <td class="px-3 py-2 text-slate-600">${row[1] || '-'}</td>
      <td class="px-3 py-2 text-slate-600">${row[2] || '-'}</td>
      <td class="px-3 py-2 text-center">${statusBadge}</td>
    `;
    tbody.appendChild(tr);
  });
}

// Manual Code Entry Handler
function initManualEntry() {
  const addBtn = document.getElementById('add-code-btn');
  const codeInput = document.getElementById('manual-code-input');

  if (addBtn && codeInput) {
    const processCode = () => {
      const val = codeInput.value.trim();
      if (val) {
        registerCodeScan(val);
        codeInput.value = '';
      }
    };

    addBtn.addEventListener('click', processCode);
    codeInput.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') processCode();
    });
  }
}

// Register Scanned/Entered Code
function registerCodeScan(code) {
  state.scannedCodes.add(code);
  state.scanHistory.unshift({
    code: code,
    time: new Date().toLocaleTimeString()
  });

  renderExcelTable();
  renderHistoryTable();
  updateSummaryMetrics();
}

// Render History Table
function renderHistoryTable() {
  const tbody = document.getElementById('history-table-body');
  if (!tbody) return;

  tbody.innerHTML = '';

  state.scanHistory.forEach(item => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50 border-b border-slate-100 text-xs';
    tr.innerHTML = `
      <td class="px-3 py-2 text-slate-500">${item.time}</td>
      <td class="px-3 py-2 font-mono font-medium text-slate-900">${item.code}</td>
    `;
    tbody.appendChild(tr);
  });
}

// Update Dashboard Cards
function updateSummaryMetrics() {
  const totalEl = document.getElementById('total-items-count');
  const verifiedEl = document.getElementById('verified-items-count');
  const pendingEl = document.getElementById('pending-items-count');

  const total = state.excelData.length;
  let verifiedCount = 0;

  state.excelData.forEach(row => {
    const code = String(row[0] || '').trim();
    if (state.scannedCodes.has(code)) verifiedCount++;
  });

  if (totalEl) totalEl.textContent = total;
  if (verifiedEl) verifiedEl.textContent = verifiedCount;
  if (pendingEl) pendingEl.textContent = Math.max(0, total - verifiedCount);
}

// Camera Scanner Handling
function startScanner() {
  if (state.html5QrcodeScanner) return;

  const readerEl = document.getElementById('reader');
  if (!readerEl) return;

  state.html5QrcodeScanner = new Html5Qrcode("reader");
  
  const config = { fps: 10, qrbox: { width: 250, height: 250 } };
  
  state.html5QrcodeScanner.start(
    { facingMode: "environment" },
    config,
    (decodedText) => {
      registerCodeScan(decodedText);
      if (navigator.vibrate) navigator.vibrate(100);
    },
    (errorMessage) => {
      // Camera scan frame errors (normal)
    }
  ).catch(err => console.error("Error iniciando cámara", err));
}

function stopScanner() {
  if (state.html5QrcodeScanner) {
    state.html5QrcodeScanner.stop().then(() => {
      state.html5QrcodeScanner.clear();
      state.html5QrcodeScanner = null;
    }).catch(err => console.error("Error al detener cámara", err));
  }
}

// Export Results to Excel
function initExport() {
  const exportBtn = document.getElementById('export-excel-btn');
  if (!exportBtn) return;

  exportBtn.addEventListener('click', () => {
    if (state.excelData.length === 0) {
      alert('No hay datos para exportar.');
      return;
    }

    const exportRows = state.excelData.map(row => {
      const code = String(row[0] || '').trim();
      const status = state.scannedCodes.has(code) ? 'Verificado' : 'Pendiente';
      return [...row, status];
    });

    const headers = [...state.excelHeaders, 'Estado Verificación'];
    const worksheetData = [headers, ...exportRows];

    const ws = XLSX.utils.aoa_to_sheet(worksheetData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Auditoria_CodeCheck");

    XLSX.writeFile(wb, `CodeCheck_Auditoria_${new Date().toISOString().slice(0,10)}.xlsx`);
  });
}
