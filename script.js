const translations = {
  "mr": {
    "description": "ही संकेतस्थळ छत्रपती संभाजीनगर महानगरपालिकेच्या सेवांसाठी तयार करण्यात आलेली आहे.",
    "dashTitle": "उपलब्ध TDR जागा",
    "srNo": "अनुक्रमांक",
    "nameAddress": "टि.डी.आर. धारकांचे नाव व पत्ता",
    "location": "मोहल्ला",
    "tdr": "क्षेत्र चौ.मी.",
    "analyticsTitle": "टीडीआरची मोजणी",
    "analyticsAvailableLabel": "बाकी टीडीआर क्षेत्र",
    "analyticsTotalLabel": "एकूण टीडीआर क्षेत्र",
    "analyticsUsedLabel": "वापरलेले टीडीआर क्षेत्र",
    "remainingTdr": "बाकी क्षेत्र चौ. मी.",
    "rr": "आर आर रेट",
    "formTitle1": "उपलब्ध TDR धारकांची यादी डाउनलोड करा :",
    "formTitle2": "टीडीआर संबंधित माहिती पाहण्यासाठी मोहल्ला निवडा",
    "srLabel": "अनुक्रमांक",
    "mohallaLabel": "मोहल्ला",
    "submitBtn": "सबमिट",
    "downloadPdf": "PDF डाउनलोड करा",
    // "downloadExcel": "Excel डाउनलोड करा"
  },
  "en": {
    "description": "This website is designed for the services of Chhatrapati Sambhajinagar Municipal Corporation.",
    "dashTitle": "Available TDR Spaces",
    "srNo": "Serial Number",
    "nameAddress": "Name & Address of TDR Holder",
    "location": "Locality",
    "tdr": "Area (sq. m.)",
    "analyticsTitle": "TDR Count",
    "analyticsAvailableLabel": "Available TDR Area",
    "analyticsTotalLabel": "Total TDR Area",
    "analyticsUsedLabel": "Used TDR Area",
    "remainingTdr": "Available Area (sq. m.)",
    "rr": "RR Rate",
    "formTitle1": "Download the list of available TDR holders:",
    "formTitle2": "Select Mohalla to view TDR details",
    "srLabel": "Serial Number(Sr. No.)",
    "mohallaLabel": "Mohalla",
    "submitBtn": "Submit",
    "downloadPdf": "Download PDF",
    // "downloadExcel": "Download Excel"
  }
};

function setLanguage(lang) {
  currentLang = lang;
  document.documentElement.lang = lang;
  document.querySelectorAll("[data-key]").forEach(el => {
    el.textContent = translations[lang][el.getAttribute("data-key")];
  });

  // Update the first option text of the mohalla select if present
  const sel = document.getElementById("mohallaSelect");
  if (sel) {
    const opt = sel.querySelector('option');
    if (opt) opt.textContent = (lang === 'mr') ? '-- मोहल्ला निवडा --' : '-- Select Mohalla --';
  }

// update placeholders safely (guard against commented-out / missing form elements)
const namePlaceholderEl = document.getElementById("nameInput");
if (namePlaceholderEl) namePlaceholderEl.placeholder = lang === "mr" ? "नाव व पत्ता" : "Name & Address";

const locationPlaceholderEl = document.getElementById("locationInput");
if (locationPlaceholderEl) locationPlaceholderEl.placeholder = lang === "mr" ? "मोहल्ला" : "Location";

const tdrPlaceholderEl = document.getElementById("tdrInput");
if (tdrPlaceholderEl) tdrPlaceholderEl.placeholder = lang === "mr" ? "क्षेत्र चौ.मी." : "Area (sq. m.)";

const remainingPlaceholderEl = document.getElementById("remainingInput");
if (remainingPlaceholderEl) remainingPlaceholderEl.placeholder = lang === "mr" ? "बाकी क्षेत्र चौ. मी." : "Remaining Area (sq. m.)";

const rrPlaceholderEl = document.getElementById("rrInput");
if (rrPlaceholderEl) rrPlaceholderEl.placeholder = lang === "mr" ? "आर आर रेट" : "RR Rate";
    
  document.getElementById("title").textContent =
    lang === "mr"
      ? "छत्रपती संभाजीनगर महानगरपालिका"
      : "Chhatrapati Sambhajinagar Municipal Corporation";
  // re-render analytics to pick up language-specific labels
  try { loadAnalyticsChart(); } catch(e) { /* ignore if not ready */ }
  // Update ticker text if present — prefer element's data-<lang> attribute, fallback to translation description
  const ticker = document.getElementById('tickerText');
  if (ticker) {
    const t = ticker.dataset && ticker.dataset[lang];
    ticker.textContent = t || translations[lang] && translations[lang].description || '';
  }
  
  // Update footer text
  const footerText = document.getElementById('footerText');
  if (footerText) {
    const footerT = footerText.dataset && footerText.dataset[lang];
    footerText.textContent = footerT || '';
  }
  
  // Update copyright text
  const copyrightText = document.getElementById('copyrightText');
  if (copyrightText) {
    const copyrightT = copyrightText.dataset && copyrightText.dataset[lang];
    copyrightText.textContent = copyrightT || '';
  }
}

const SHEET_URL_FALLBACK_NOTE = "PLEASE_SET_SHEET_URL"; // internal check
// SHEET_URL and SHEET_ID are set from HTML <script> block. (Replace SHEET_URL with deployment URL)
if (typeof SHEET_URL === "https://script.google.com/macros/s/AKfycbz_k3x8xyJB8pj2ok9RfqEQv1iBY0PAP2Mmdf1FfplKlYXBD7yqw4GfqDhveZpq7_9b5Q/exec" || !SHEET_URL || SHEET_URL.indexOf("http") !== 0) {
  console.warn("SHEET_URL is not set. Please paste your Apps Script deployment URL into the HTML file's SHEET_URL constant.");
}
const SHEET_URL_ACTIVE = (typeof SHEET_URL !== "undefined" && SHEET_URL && SHEET_URL.indexOf("http") === 0) ? SHEET_URL : null;
const SHEET_ID_ACTIVE = (typeof SHEET_ID !== "undefined") ? SHEET_ID : "1Ygdkm88cgryfK8yD1YjUcyTxHjTcnDabwrv1LXohYjE";

// Optional: put your Google Spreadsheet ID here to enable direct export of the current sheet
// If SHEET_ID is provided and the sheet's sharing allows access, the download will be the current live sheet.
// Example: const SHEET_ID = '1aBcD...';
// const SHEET_ID = ""; // set in HTML

let sheetData = [];
let loaded = false;

let analyticsChartInstance = null;
let currentRow = null; 
let currentLang = document.documentElement.lang || 'mr';
// index of header row in sheetData (0-based). Your sheet uses headers in row 3, so we default to index 2.
let headerIndex = 2;
let colIndex = { sr: 0, name: 1, location: 2, tdr: 3, remaining: 4, rr: 5 };

function sanitizeNumericInfo(value) {
  if (value === null || value === undefined) return null;
  const str = String(value).replace(/,/g, '').trim();
  if (!str || str === '.' || str === '-' || str === '-.') return null;
  const num = Number(str);
  if (!Number.isFinite(num)) return null;
  return { raw: str, num };
}

function applyIndianGrouping(numString) {
  if (!numString) return '';
  const negative = numString.startsWith('-');
  const payload = negative ? numString.slice(1) : numString;
  const [intRaw, decimalPart] = payload.split('.');
  let intPart = intRaw || '0';
  intPart = intPart.replace(/^0+(?=\d)/, '');
  if (intPart.length > 3) {
    const lastThree = intPart.slice(-3);
    const rest = intPart.slice(0, -3);
    intPart = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + lastThree;
  }
  const prefix = negative ? '-' : '';
  return prefix + intPart + (decimalPart ? '.' + decimalPart : '');
}

function formatNumberIndian(value, decimals = null) {
  const info = sanitizeNumericInfo(value);
  if (!info) return typeof value === 'string' ? value : (value ?? '');
  const base = (decimals !== null && decimals !== undefined)
    ? info.num.toFixed(decimals)
    : info.raw;
  return applyIndianGrouping(base);
}

function formatNumberIndianWithUnit(value, decimals = null, unit = '') {
  const info = sanitizeNumericInfo(value);
  if (!info) return typeof value === 'string' ? value : (value ?? '');
  const formatted = formatNumberIndian(value, decimals);
  return unit ? `${formatted} ${unit}` : formatted;
}

function formatValueIfNumeric(value, decimals = null) {
  const info = sanitizeNumericInfo(value);
  if (!info) return value ?? '';
  return formatNumberIndian(value, decimals);
}

// Helper: find header row and map column indices by keyword matching
function detectHeaderAndColumns() {
  // Your sheet has the header in row 3 (1-based), so force the header row to index 2 (0-based)
  headerIndex = 2;

  const headers = sheetData[headerIndex] || [];
  const find = (variants) => {
    const vLower = variants.map(x => x.toLowerCase());
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i] || '').toLowerCase();
      for (const v of vLower) if (h.includes(v)) return i;
    }
    return -1;
  };

  // map common header names to columns (fallback to existing colIndex if not found)
  const srIdx = find(['sr','serial']); if (srIdx >= 0) colIndex.sr = srIdx;
  const nameIdx = find(['name','address','नाव']); if (nameIdx >= 0) colIndex.name = nameIdx;
  const locIdx = find(['location','place','ठिकाण','मोहल्ला']); if (locIdx >= 0) colIndex.location = locIdx;
  const tdrIdx = find(['tdr','area','क्षेत्र']); if (tdrIdx >= 0) colIndex.tdr = tdrIdx;
  const remIdx = find(['remaining','available','उर्वरित','बाकी']); if (remIdx >= 0) colIndex.remaining = remIdx;
  const rrIdx = find(['rr','rate','आर आर']); if (rrIdx >= 0) colIndex.rr = rrIdx;

  return { headerIndex, headers };
}

function loadDashboard() {
  const table = document.getElementById("tdrTable");
  if (!table) return;

  // detect header row and columns (handles sheets with unused rows at top)
  const info = detectHeaderAndColumns();
  const headers = info.headers || [];
  const rows = sheetData.slice(headerIndex + 1);

  let thead = '<thead><tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr></thead>';
  let tbody = '<tbody>' + rows.map(r => `<tr>${r.map(v => `<td>${formatValueIfNumeric(v)}</td>`).join('')}</tr>`).join('') + '</tbody>';

  table.innerHTML = thead + tbody;
}
function showResult(row) {
  // Ensure header mapping updated
  detectHeaderAndColumns();
  const headers = sheetData[headerIndex] || [];
    let html = `
        <table id="resultTable">
            <thead>
                <tr>
                    ${headers.map(h => `<th>${h}</th>`).join("")}
                </tr>
            </thead>
            <tbody>
                <tr>
                    ${row.map(v => `<td>${formatValueIfNumeric(v)}</td>`).join("")}
                </tr>
            </tbody>
        </table>
    `;

    const box = document.getElementById("resultBox");
    if (box) {
      box.innerHTML = html;
      box.style.display = "block";
    } else {
      console.warn('showResult: #resultBox not found; skipping rendering.');
    }
    currentRow = row;
    showResultActions(true);
}

function findRow(sr) {
  if (!sheetData || sheetData.length === 0) return null;
  const rows = sheetData.slice(headerIndex + 1);
  const srCol = colIndex.sr || 0;
  return rows.find(r => {
    const cell = String(r[srCol] ?? '').trim();
    // compare normalized sr values (keep alphanumeric) but trim spaces
    const norm = s => String(s ?? '').toLowerCase().replace(/[^a-z0-9]/gi, '');
    return norm(cell) === norm(sr);
  });
}

function fetchBySr() {
  const el = document.getElementById("srInput");
  if (!el) return; // SR input removed — keep function safe if called elsewhere
  const sr = el.value.trim();

  if (!sr) {
    clearAutofill();
    hideResult();
    return;
  }

  if (!loaded) {
    console.warn("Sheet data not loaded yet — will try again in 300ms");
    setTimeout(fetchBySr, 300);
    return;
  }

  const row = findRow(sr);
  if (row) {
    showResult(row);
    autofillFields(row);
  } else {
    hideResult();
    clearAutofill();
  }
}

// -- New location-based filtering helpers --
function normalizeForCompare(s) {
  return String(s ?? '').toLowerCase().replace(/[^a-z0-9\u0900-\u097F\s]/gi, '').trim();
}

function findRowsByLocation(location) {
  if (!sheetData || sheetData.length === 0) return [];
  if (!location) return [];
  const rows = sheetData.slice(headerIndex + 1);
  const locCol = colIndex.location || 2;
  const needle = normalizeForCompare(location);
  // special-case: return all rows when user selects the __ALL__ sentinel
  if (location === '__ALL__') return rows;

  return rows.filter(r => {
    const cell = String(r[locCol] ?? '').trim();
    return normalizeForCompare(cell) === needle;
  });
}

function populateMohallaDropdown() {
  if (!sheetData || sheetData.length <= headerIndex) return;
  detectHeaderAndColumns();
  const rows = sheetData.slice(headerIndex + 1);
  const locCol = colIndex.location || 2;

  const set = new Set();
  rows.forEach(r => {
    const v = String(r[locCol] ?? '').trim();
    if (v) set.add(v);
  });

  const arr = Array.from(set).sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));
  const sel = document.getElementById('mohallaSelect');
  if (!sel) return;

  const allLabel = currentLang === 'mr' ? 'सर्व मोहल्ले' : 'All Mohallas';
  const defaultText = currentLang === 'mr' ? '-- मोहल्ला निवडा --' : '-- Select Mohalla --';

  sel.innerHTML = '';
  const optDefault = document.createElement('option');
  optDefault.value = '';
  optDefault.textContent = defaultText;
  sel.appendChild(optDefault);

  // optional "Show all" choice
  const optAll = document.createElement('option');
  optAll.value = '__ALL__';
  optAll.textContent = allLabel;
  sel.appendChild(optAll);

  arr.forEach(v => {
    const o = document.createElement('option');
    o.value = v;
    o.textContent = v;
    sel.appendChild(o);
  });
}

function fetchByLocation() {
  const sel = document.getElementById('mohallaSelect');
  if (!sel) return;
  const loc = sel.value.trim();

  if (!loc) {
    clearAutofill();
    hideResult();
    return;
  }

  if (!loaded) {
    console.warn("Sheet data not loaded yet — will try again in 300ms");
    setTimeout(fetchByLocation, 300);
    return;
  }

  const rows = findRowsByLocation(loc);
  if (rows && rows.length) {
    showResults(rows);
    if (rows.length === 1) {
      autofillFields(rows[0]);
    } else {
      clearAutofill();
    }
  } else {
    hideResult();
    clearAutofill();
  }
}

function showResults(rows) {
  detectHeaderAndColumns();
  const headers = sheetData[headerIndex] || [];
  
  // Get rows before header (rows 0 and 1 if headerIndex is 2)
  // Commented out for website display - will be included in PDF download only
  let preHeaderHtml = '';
  // if (sheetData && headerIndex > 0) {
  //   for (let i = 0; i < headerIndex; i++) {
  //     const row = sheetData[i] || [];
  //     if (row.length > 0 && row.some(cell => cell)) {
  //       // Check if this row has content
  //       preHeaderHtml += `<tr class="pre-header-row"><td colspan="${headers.length}" style="text-align: center; font-weight: bold; padding: 10px; background: #f0f3f7;">${row.join(' ')}</td></tr>`;
  //     }
  //   }
  // }

  // Helper function to check if value is numeric
  const isNumeric = (val) => {
    const info = sanitizeNumericInfo(val);
    return info !== null;
  };

  let html = `
        <table id="resultTable">
            <thead>
                ${preHeaderHtml}
                <tr>
                    ${headers.map(h => `<th>${h}</th>`).join("")}
                </tr>
            </thead>
            <tbody>
                ${rows.map(r => `<tr>${r.map(v => {
                  const formatted = formatValueIfNumeric(v);
                  const numericClass = isNumeric(v) ? ' class="numeric-cell"' : '';
                  return `<td${numericClass}>${formatted}</td>`;
                }).join("")}</tr>`).join('')}
            </tbody>
        </table>
    `;

  const box = document.getElementById("resultBox");
  if (box) {
    box.innerHTML = html;
    box.style.display = "block";
    // Scroll to result table after a short delay to ensure DOM is updated
    setTimeout(() => {
      box.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 150);
  } else {
    console.warn('showResults: #resultBox not found; skipping rendering.');
  }
  // if a single row, set currentRow so result actions / downloads will use it
  currentRow = rows.length === 1 ? rows[0] : null;
  showResultActions(true);
}

function clearAutofill() {
  const nameEl = document.getElementById("nameInput"); if (nameEl) nameEl.value = "";
  const locationEl = document.getElementById("locationInput"); if (locationEl) locationEl.value = "";
  const tdrEl = document.getElementById("tdrInput"); if (tdrEl) tdrEl.value = "";
  const remainingEl = document.getElementById("remainingInput"); if (remainingEl) remainingEl.value = "";
  const rrEl = document.getElementById("rrInput"); if (rrEl) rrEl.value = "";
}

function hideResult() {
  const box = document.getElementById("resultBox");
  if (box) {
    box.style.display = "none";
    box.innerHTML = "";
  }
  currentRow = null;
  showResultActions(false);
}

function showResultActions(visible = true) {
  const actions = document.getElementById('resultActions');
  if (!actions) return;
  actions.style.display = visible ? 'flex' : 'none';
}

async function uploadResult() {
  if (!currentRow) {
    alert('No result to upload. Please select an SR first.');
    return;
  }

  const headers = sheetData[headerIndex] || [];
  const rowObj = {};
  headers.forEach((h, i) => { rowObj[h] = currentRow[i] ?? ''; });

  const payload = { action: 'uploadResult', row: rowObj };

  // Try POST to the same script endpoint (if it supports POST action)
  try {
    const postUrl = (SHEET_URL_ACTIVE ? SHEET_URL_ACTIVE : "") + '?action=upload';
    const res = await fetch(postUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    if (res.ok) {
      const data = await res.json().catch(() => null);
      alert('Upload successful' + (data && data.message ? ': ' + data.message : ''));
      return;
    }
  } catch (err) {
    console.warn('Upload POST failed, falling back to download', err);
  }

  try {
    const csv = `${headers.map(h => '"' + String(h).replace(/"/g,'""') + '"').join(',')}\n${currentRow.map(v => '"' + String(v ?? '').replace(/"/g,'""') + '"').join(',')}`;
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `sr-${String(currentRow[0] ?? 'result')}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    alert('Upload endpoint not available — downloaded CSV instead.');
  } catch (err) {
    console.error('Fallback download failed', err);
    alert('Unable to upload or download result. Check console for details.');
  }
}
async function downloadSheetPdf() {
  // Always export the entire Google Sheet as PDF using the same minimal URL
  // that works when opened directly in the browser.
  if (SHEET_ID_ACTIVE) {
    const exportUrl =
      `https://docs.google.com/spreadsheets/d/${SHEET_ID_ACTIVE}/export?format=pdf`;

    const win = window.open(exportUrl, '_blank');
    if (win) win.opener = null;
    return;
  }

  // Fallback: open a static local PDF if no SHEET_ID is configured
  window.open('TDR - Sheet1.pdf', '_blank');
}

  function downloadResultPDF() {
    if (!currentRow) {
      alert('No result to download. Please select an SR first.');
      return;
    }

    const el = document.getElementById('resultTable');
    if (!el) {
      alert('No result table present to download.');
      return;
    }

    const filename = `SR-${String(currentRow[0] ?? 'result')}.pdf`;
    const opt = {
      margin:       10,
      filename:     filename,
      image:        { type: 'jpeg', quality: 0.98 },
      html2canvas:  { scale: 2 },
      jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    try {
      html2pdf().set(opt).from(el).save();
    } catch (err) {
      console.error('PDF download failed', err);
      alert('PDF download failed — check console for details.');
    }
  }

function autofillFields(row) {
  // Use detected column indices so autofill works even when header isn't first row
  const nameEl = document.getElementById("nameInput"); if (nameEl) nameEl.value = row[colIndex.name] ?? "";
  const locationEl = document.getElementById("locationInput"); if (locationEl) locationEl.value = row[colIndex.location] ?? "";
  const tdrEl = document.getElementById("tdrInput"); if (tdrEl) tdrEl.value = formatValueIfNumeric(row[colIndex.tdr]);
  const remainingEl = document.getElementById("remainingInput"); if (remainingEl) remainingEl.value = formatValueIfNumeric(row[colIndex.remaining]);
  const rrEl = document.getElementById("rrInput"); if (rrEl) rrEl.value = formatValueIfNumeric(row[colIndex.rr]);
}

async function initialize() {
  try {
    if (!SHEET_URL_ACTIVE) {
      // If no Apps Script URL present, attempt direct GET to Google Apps Script path will fail.
      // But if you configured SHEET_URL in HTML, we will use it.
      console.warn("SHEET_URL not configured; please set the deployment URL in HTML.");
    }

    const url = (SHEET_URL_ACTIVE ? (SHEET_URL_ACTIVE + "?action=getData") : ("/not-configured?action=getData"));
    const res = await fetch(url);
    const data = await res.json();

    sheetData = data;
    loaded = true;

    loadDashboard();
    // ensure UI text reflects page language and update charts accordingly
    try { setLanguage(currentLang); } catch(e) { /* ignore */ }

    // Wire Mohalla select for filtering
    populateMohallaDropdown();
    const selEl = document.getElementById('mohallaSelect');
    if (selEl) selEl.addEventListener('change', fetchByLocation);
  } catch (err) {
    console.error("Failed to load sheet data:", err);
  }
}
// start initialization when DOM ready
window.addEventListener("DOMContentLoaded", initialize);

// set CSS variable for header height so #main-bg can start right below header
function updateHeaderHeightVar() {
  const header = document.querySelector('header');
  if (!header) { document.documentElement.style.setProperty('--header-height', '64px'); return; }
  // Use header's bottom relative to viewport so the background top aligns exactly under the visible header
  const rect = header.getBoundingClientRect();
  const px = Math.max(0, Math.round(rect.bottom));
  const val = px + 'px';
  document.documentElement.style.setProperty('--header-height', val);
}

// update on load and resize so the background top always aligns under header
window.addEventListener('load', updateHeaderHeightVar);
window.addEventListener('resize', updateHeaderHeightVar);
// update on scroll as well so background follows header bottom when header changes size/position
let _rafHeader = null;
window.addEventListener('scroll', () => {
  if (_rafHeader) cancelAnimationFrame(_rafHeader);
  _rafHeader = requestAnimationFrame(() => { updateHeaderHeightVar(); _rafHeader = null; });
});
// also run immediately
updateHeaderHeightVar();
// distribution chart removed — only single banner analytics chart is displayed under header

function loadAnalyticsChart() {
  // ensure we detected header row and column mapping
  const info = detectHeaderAndColumns();
  const rows = sheetData.slice(headerIndex + 1);
  const headers = (sheetData[headerIndex] || []).map(h => String(h || "").toLowerCase());

  // Use mapped column indices (colIndex) and extract numeric values from alphanumeric cells
  const remainingIdx = colIndex.remaining;
  const totalIdx = colIndex.tdr;

  // helper to extract first numeric value from a cell's string (supports commas, spaces)
  const extractNumeric = (cell) => {
    if (cell == null) return 0;
    const s = String(cell).replace(/[,\s]/g, '');
    const m = s.match(/-?\d+(?:\.\d+)?/);
    return m ? parseFloat(m[0]) : 0;
  };

  // compute sums using numeric extraction only
  const sumRemaining = remainingIdx >= 0 ? rows.reduce((s, r) => s + extractNumeric(r[remainingIdx]), 0) : 0;
  const sumTotal = totalIdx >= 0 ? rows.reduce((s, r) => s + extractNumeric(r[totalIdx]), 0) : 0;

  // prefer banner chart id (right below header), fallback to old id
  const bannerCanvas = document.getElementById("tdrAnalyticsBannerChart") || document.getElementById("tdrAnalyticsChart");
  if (!bannerCanvas) return;
  const ctx2 = bannerCanvas.getContext("2d");

  const labels = [];
  const values = [];

  const availableLabel = (translations[currentLang] && translations[currentLang]['analyticsAvailableLabel']) || 'Available';
  const usedLabel = (translations[currentLang] && translations[currentLang]['analyticsUsedLabel']) || 'Used';
  const sumUsed = Math.max(0, sumTotal - sumRemaining);
  const unit = currentLang === 'mr' ? 'चौ.मी.' : 'sq.m.';

  if (sumTotal > 0) {
    labels.push(availableLabel, usedLabel);
    values.push(sumRemaining, sumUsed);
  } else if (sumRemaining > 0) {
    labels.push(availableLabel);
    values.push(sumRemaining);
  } else {
    labels.push("No numeric TDR data");
    values.push(1);
  }

  if (analyticsChartInstance) analyticsChartInstance.destroy();
  // create two-tone gradients for slices
  const gradBlue = ctx2.createLinearGradient(0, 0, 0, 160);
  gradBlue.addColorStop(0, '#66ccff');
  gradBlue.addColorStop(1, '#0288d1');

  const gradOrange = ctx2.createLinearGradient(0, 0, 0, 160);
  gradOrange.addColorStop(0, '#ffd699');
  gradOrange.addColorStop(1, '#ff9800');

  // simple plugin to draw central percentage/value in the doughnut
  const centerTextPlugin = {
    id: 'centerText',
    afterDraw(chart) {
      if (!chart || !chart.data || !chart.data.datasets || !chart.data.datasets.length) return;
      const meta = chart.getDatasetMeta(0);
      if (!meta || !meta.data || !meta.data.length) return;
      const { ctx, chartArea: { top, bottom, left, right } } = chart;
      const width = right - left;
      const height = bottom - top;
      const centerX = left + width / 2;
      const centerY = top + height / 2;

      const data = chart.data.datasets[0].data;
      const details = chart.config.options.centerDetails || {};
      const availableValue = details.value != null ? details.value : (data[0] || 0);
      const label = details.label || 'Available';
      const percentage = (typeof details.percentage === 'number')
        ? details.percentage
        : (() => {
            const total = data.reduce((s, v) => s + (Number(v) || 0), 0) || 0;
            if (!total) return 0;
            return (availableValue / total) * 100;
          })();

      ctx.save();
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';

      const baseSize = Math.min(width, height);
      const pctFont = Math.max(16, baseSize / 9.5);
      const labelFont = Math.max(12, baseSize / 18);

      ctx.font = `700 ${pctFont}px Arial`;
      ctx.fillStyle = '#0c1f33';
      ctx.fillText(`${percentage.toFixed(2)}%`, centerX, centerY - (height * 0.025));

      ctx.font = `600 ${labelFont}px Arial`;
      ctx.fillStyle = '#4e5a6a';
      ctx.fillText(label, centerX, centerY + (height * 0.07));
      ctx.restore();
    }
  };

  // compute per-slice offsets: make the first slice (Available) pop out slightly
  const offsets = values.map((v, i) => i === 0 ? 12 : 0);

  analyticsChartInstance = new Chart(ctx2, {
    type: 'doughnut',
    data: {
      labels: labels,
        datasets: [{
        data: values,
        offset: offsets,
        backgroundColor: values.length === 2 ? [gradBlue, gradOrange] : [gradBlue],
        borderColor: "#ffffff",
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      cutout: '62%',
      animation: {
        animateScale: true,
        duration: 900,
        easing: 'easeOutCubic'
      },
      hoverOffset: 0,
      plugins: {
        tooltip: { enabled: false },
        legend: { position: 'bottom' }
      },
      onClick: null,
      centerDetails: {
        label: availableLabel,
        unit,
        value: sumRemaining,
        percentage: (sumTotal > 0 ? (sumRemaining / sumTotal) * 100 : 0)
      }
    },
    plugins: [centerTextPlugin]
  });
  // Set numeric summary below banner chart if present
  const totalEl = document.getElementById('analytics-total');
  const availEl = document.getElementById('analytics-available');
  const usedEl = document.getElementById('analytics-used');
  const zeroDisplay = `0 ${unit}`;
  const formatSummaryValue = (value) => {
    const info = sanitizeNumericInfo(value);
    if (!info) return '';
    return formatNumberIndianWithUnit(value, 2, unit);
  };
  if (totalEl) totalEl.textContent = formatSummaryValue(sumTotal) || zeroDisplay;
  if (usedEl) usedEl.textContent = formatSummaryValue(sumUsed) || zeroDisplay;
  if (availEl) availEl.textContent = formatSummaryValue(sumRemaining) || zeroDisplay;
}


// Download result table as Excel (CSV format)
function downloadResultExcel() {
  const resultTable = document.getElementById('resultTable');
  if (!resultTable) {
    alert(currentLang === 'mr' ? 'डाउनलोड करण्यासाठी कोणताही परिणाम नाही.' : 'No result to download.');
    return;
  }

  let csv = [];
  const rows = resultTable.querySelectorAll('tr');
  
  rows.forEach(row => {
    const cols = row.querySelectorAll('td, th');
    const csvRow = [];
    cols.forEach(col => {
      let text = col.innerText.replace(/"/g, '""'); // Escape quotes
      csvRow.push('"' + text + '"');
    });
    csv.push(csvRow.join(','));
  });

  const csvContent = csv.join('\n');
  const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  
  link.setAttribute('href', url);
  link.setAttribute('download', 'TDR_Result_' + new Date().getTime() + '.csv');
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// Download result table as PDF (existing function - enhanced)
function downloadResultPDF() {
  const resultBox = document.getElementById('resultBox');
  if (!resultBox || resultBox.style.display === 'none') {
    alert(currentLang === 'mr' ? 'डाउनलोड करण्यासाठी कोणताही परिणाम नाही.' : 'No result to download.');
    return;
  }

  // Clone the result box to add pre-header rows and styling for PDF
  const clonedBox = resultBox.cloneNode(true);
  const table = clonedBox.querySelector('#resultTable');
  
  if (table && sheetData && headerIndex > 0) {
    const headers = sheetData[headerIndex] || [];
    const thead = table.querySelector('thead');
    
    // Add pre-header rows at the beginning of thead
    let preHeaderHtml = '';
    for (let i = 0; i < headerIndex; i++) {
      const row = sheetData[i] || [];
      if (row.length > 0 && row.some(cell => cell)) {
        preHeaderHtml += `<tr class="pre-header-row"><td colspan="${headers.length}" style="text-align: center; font-weight: bold; padding: 10px; background: #f0f3f7; border: 1px solid #ddd;">${row.join(' ')}</td></tr>`;
      }
    }
    
    if (preHeaderHtml && thead) {
      thead.insertAdjacentHTML('afterbegin', preHeaderHtml);
    }
  }
  
  // Add CSS to prevent rows from breaking across pages and align numeric cells
  const style = document.createElement('style');
  style.textContent = `
    #resultTable tr { page-break-inside: avoid; }
    #resultTable thead { display: table-header-group; }
    #resultTable tbody { display: table-row-group; }
    #resultTable td.numeric-cell { text-align: right !important; }
    #resultTable td { text-align: left; }
  `;
  clonedBox.appendChild(style);

  const opt = {
    margin: [8, 8, 8, 8],
    filename: 'TDR_Result_' + new Date().getTime() + '.pdf',
    image: { type: 'jpeg', quality: 0.95 },
    html2canvas: { 
      scale: 2, 
      useCORS: true,
      letterRendering: true,
      logging: false
    },
    jsPDF: { 
      unit: 'mm', 
      format: 'a4', 
      orientation: 'landscape',
      compress: true
    },
    pagebreak: { 
      mode: ['avoid-all', 'css', 'legacy'],
      before: '.page-break-before',
      after: '.page-break-after',
      avoid: ['tr', 'td', 'th']
    }
  };

  html2pdf().set(opt).from(clonedBox).save();
}
