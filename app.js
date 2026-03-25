/* ============================================================
   GUILD ELECTRIC — WARRANTY MANAGEMENT SYSTEM
   app.js — all application logic
   ============================================================ */

'use strict';

// ── STATE ──────────────────────────────────────────────────
let warranties = [];
let editingId  = null;
let currentFilter = 'all';
let currentSearch = '';
let pdfWarranty = null;

// ── DOM REFS ───────────────────────────────────────────────
const formPanel      = document.getElementById('formPanel');
const formTitle      = document.getElementById('formTitle');
const warrantyForm   = document.getElementById('warrantyForm');
const btnNew         = document.getElementById('btnNew');
const btnCloseForm   = document.getElementById('btnCloseForm');
const btnCancelForm  = document.getElementById('btnCancelForm');
const btnSave        = document.getElementById('btnSave');
const fileOpen       = document.getElementById('fileOpen');
const searchInput    = document.getElementById('searchInput');
const pdfModal       = document.getElementById('pdfModal');
const btnCloseModal  = document.getElementById('btnCloseModal');
const btnDownloadPDF = document.getElementById('btnDownloadPDF');
const modalContent   = document.getElementById('modalContent');
const modalTitle     = document.getElementById('modalTitle');
const toast          = document.getElementById('toast');

const fProjectName   = document.getElementById('fProjectName');
const fProjectAddress= document.getElementById('fProjectAddress');
const fIssuedTo      = document.getElementById('fIssuedTo');
const fSubject       = document.getElementById('fSubject');
const fDateIssued    = document.getElementById('fDateIssued');
const fEffectiveDate = document.getElementById('fEffectiveDate');
const fWarrantyPeriod= document.getElementById('fWarrantyPeriod');
const fCustomPeriod  = document.getElementById('fCustomPeriod');
const customPeriodGroup = document.getElementById('customPeriodGroup');
const fWO            = document.getElementById('fWO');
const fPM            = document.getElementById('fPM');
const fNotes         = document.getElementById('fNotes');

// ── UTILS ──────────────────────────────────────────────────
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function today() {
  return new Date().toISOString().split('T')[0];
}

function parseYMD(str) {
  if (!str) return null;
  const [y, m, d] = str.split('-').map(Number);
  return new Date(y, m - 1, d);
}

function formatDate(str) {
  const d = parseYMD(str);
  if (!d) return '—';
  return d.toLocaleDateString('en-CA', { year: 'numeric', month: 'short', day: 'numeric' });
}

function calcEndDate(effectiveDate, period) {
  const d = parseYMD(effectiveDate);
  if (!d) return null;
  const p = period || '1 Calendar Year';
  const match = p.match(/^(\d+)/);
  const years = match ? parseInt(match[1]) : 1;
  const end = new Date(d);
  end.setFullYear(end.getFullYear() + years);
  end.setDate(end.getDate() - 1);
  return end.toISOString().split('T')[0];
}

function daysUntil(dateStr) {
  const d = parseYMD(dateStr);
  if (!d) return null;
  const now = new Date(); now.setHours(0,0,0,0);
  return Math.round((d - now) / 86400000);
}

function getStatus(w) {
  const days = daysUntil(w.endDate);
  if (days === null) return 'active';
  if (days < 0) return 'expired';
  if (days <= 30) return 'expiring';
  return 'active';
}

// ── TOAST ──────────────────────────────────────────────────
let toastTimer;
function showToast(msg, type = 'success') {
  toast.textContent = msg;
  toast.className = 'toast show' + (type !== 'success' ? ' toast-' + type : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('show'), 3000);
}

// ── FORM PANEL ─────────────────────────────────────────────
function openForm(warranty = null) {
  editingId = warranty ? warranty.id : null;
  formTitle.textContent = warranty ? 'Edit Warranty' : 'New Warranty Certificate';

  if (warranty) {
    fProjectName.value    = warranty.projectName || '';
    fProjectAddress.value = warranty.projectAddress || '';
    fIssuedTo.value       = warranty.issuedTo || '';
    fSubject.value        = warranty.subject || '';
    fDateIssued.value     = warranty.dateIssued || today();
    fEffectiveDate.value  = warranty.effectiveDate || today();
    fWarrantyPeriod.value = warranty.warrantyPeriod || '1 Calendar Year';
    if (!['1 Calendar Year','2 Calendar Years','3 Calendar Years','5 Calendar Years'].includes(warranty.warrantyPeriod)) {
      fWarrantyPeriod.value = 'Custom';
      fCustomPeriod.value   = warranty.warrantyPeriod;
      customPeriodGroup.style.display = '';
    }
    fWO.value    = warranty.wo || '';
    fPM.value    = warranty.pm || '';
    fNotes.value = warranty.notes || '';
  } else {
    warrantyForm.reset();
    fDateIssued.value    = today();
    fEffectiveDate.value = today();
    fWarrantyPeriod.value = '1 Calendar Year';
    customPeriodGroup.style.display = 'none';
  }

  formPanel.classList.remove('hidden');
  fProjectName.focus();
}

function closeForm() {
  formPanel.classList.add('hidden');
  editingId = null;
}

fWarrantyPeriod.addEventListener('change', () => {
  const isCustom = fWarrantyPeriod.value === 'Custom';
  customPeriodGroup.style.display = isCustom ? '' : 'none';
  if (isCustom) fCustomPeriod.focus();
});

btnNew.addEventListener('click', () => openForm());
btnCloseForm.addEventListener('click', closeForm);
btnCancelForm.addEventListener('click', closeForm);

warrantyForm.addEventListener('submit', e => {
  e.preventDefault();
  if (!fProjectName.value.trim() || !fIssuedTo.value.trim() || !fEffectiveDate.value) {
    showToast('Please fill in all required fields.', 'error');
    return;
  }

  const period = fWarrantyPeriod.value === 'Custom'
    ? (fCustomPeriod.value.trim() || '1 Calendar Year')
    : fWarrantyPeriod.value;

  const endDate = calcEndDate(fEffectiveDate.value, period);

  const data = {
    id:             editingId || uid(),
    projectName:    fProjectName.value.trim(),
    projectAddress: fProjectAddress.value.trim(),
    issuedTo:       fIssuedTo.value.trim(),
    subject:        fSubject.value.trim(),
    dateIssued:     fDateIssued.value,
    effectiveDate:  fEffectiveDate.value,
    warrantyPeriod: period,
    endDate:        endDate,
    wo:             fWO.value.trim(),
    pm:             fPM.value.trim(),
    notes:          fNotes.value.trim(),
    createdAt:      editingId ? (warranties.find(w => w.id === editingId)?.createdAt || Date.now()) : Date.now()
  };

  if (editingId) {
    const idx = warranties.findIndex(w => w.id === editingId);
    if (idx !== -1) warranties[idx] = data;
    showToast('Warranty updated.');
  } else {
    warranties.push(data);
    showToast('Warranty certificate issued.');
  }

  closeForm();
  renderAll();
});

// ── SAVE / OPEN ────────────────────────────────────────────
btnSave.addEventListener('click', () => {
  if (!warranties.length) { showToast('Nothing to save.', 'warn'); return; }
  const blob = new Blob([JSON.stringify({ version: 1, warranties }, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `guild-warranties-${today()}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
  showToast('Database saved.');
});

fileOpen.addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const data = JSON.parse(ev.target.result);
      if (data.warranties && Array.isArray(data.warranties)) {
        warranties = data.warranties;
        renderAll();
        showToast(`Loaded ${warranties.length} warranties.`);
      } else {
        showToast('Invalid file format.', 'error');
      }
    } catch {
      showToast('Could not read file.', 'error');
    }
  };
  reader.readAsText(file);
  fileOpen.value = '';
});

// ── SEARCH & FILTER ────────────────────────────────────────
searchInput.addEventListener('input', () => {
  currentSearch = searchInput.value.toLowerCase();
  renderAll();
});

document.querySelectorAll('.filter-tab').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.filter-tab').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentFilter = btn.dataset.filter;
    renderAll();
  });
});

// ── RENDER ─────────────────────────────────────────────────
function matches(w) {
  if (!currentSearch) return true;
  const haystack = [w.projectName, w.projectAddress, w.issuedTo, w.subject, w.wo, w.pm].join(' ').toLowerCase();
  return haystack.includes(currentSearch);
}

function renderAll() {
  updateStats();

  const active   = warranties.filter(w => getStatus(w) === 'active'    && matches(w));
  const expiring = warranties.filter(w => getStatus(w) === 'expiring'  && matches(w));
  const expired  = warranties.filter(w => getStatus(w) === 'expired'   && matches(w));

  const listActive   = document.getElementById('listActive');
  const listExpiring = document.getElementById('listExpiring');
  const listExpired  = document.getElementById('listExpired');
  const dividerActive   = document.getElementById('dividerActive');
  const dividerExpiring = document.getElementById('dividerExpiring');
  const dividerExpired  = document.getElementById('dividerExpired');
  const emptyState   = document.getElementById('emptyState');

  // Filter logic
  const showActive   = (currentFilter === 'all' || currentFilter === 'active');
  const showExpiring = (currentFilter === 'all' || currentFilter === 'expiring');
  const showExpired  = (currentFilter === 'all' || currentFilter === 'expired');

  const renderList = (container, list, divider, show) => {
    container.innerHTML = '';
    const visible = show && list.length > 0;
    divider.style.display   = visible ? '' : 'none';
    container.style.display = visible ? '' : 'none';
    if (visible) list.forEach(w => container.appendChild(buildCard(w)));
  };

  renderList(listActive,   active,   dividerActive,   showActive);
  renderList(listExpiring, expiring, dividerExpiring, showExpiring);
  renderList(listExpired,  expired,  dividerExpired,  showExpired);

  const total = active.length + expiring.length + expired.length;
  emptyState.style.display = total === 0 ? '' : 'none';
}

function updateStats() {
  const active   = warranties.filter(w => getStatus(w) === 'active').length;
  const expiring = warranties.filter(w => getStatus(w) === 'expiring').length;
  const expired  = warranties.filter(w => getStatus(w) === 'expired').length;
  document.getElementById('statTotal').textContent   = warranties.length;
  document.getElementById('statActive').textContent  = active + expiring;
  document.getElementById('statExpired').textContent = expired;
  document.getElementById('statExpiring').textContent = expiring;
}

function buildCard(w) {
  const status  = getStatus(w);
  const days    = daysUntil(w.endDate);
  const pill    = { active: 'pill-active Active', expiring: 'pill-expiring Expiring', expired: 'pill-expired Expired' }[status];
  const cardCls = { active: 'active-card', expiring: 'expiring-card', expired: 'expired-card' }[status];
  const dateCls = { active: 'date-highlight', expiring: 'date-amber', expired: 'date-gray' }[status];

  let daysLabel = '';
  if (status === 'active')   daysLabel = `<span class="${dateCls}">${days}d remaining</span>`;
  if (status === 'expiring') daysLabel = `<span class="${dateCls}">⚠ ${days}d left</span>`;
  if (status === 'expired')  daysLabel = `<span class="${dateCls}">Expired ${Math.abs(days)}d ago</span>`;

  const card = document.createElement('div');
  card.className = `warranty-card ${cardCls}`;
  card.dataset.id = w.id;
  card.innerHTML = `
    <div class="card-main">
      <div class="card-top">
        <div class="card-project" title="${esc(w.projectName)}">${esc(w.projectName)}</div>
        <span class="status-pill ${pill.split(' ')[0]}">${pill.split(' ')[1]}</span>
      </div>
      <div class="card-meta">
        <span>
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/></svg>
          ${esc(w.issuedTo)}
        </span>
        ${w.projectAddress ? `<span>
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>
          ${esc(w.projectAddress)}
        </span>` : ''}
        ${w.wo ? `<span><strong>W.O.</strong> ${esc(w.wo)}</span>` : ''}
        ${w.pm ? `<span><strong>PM:</strong> ${esc(w.pm)}</span>` : ''}
      </div>
      <div class="card-dates">
        <span>📅 Issued: ${formatDate(w.dateIssued)}</span>
        <span>▶ Start: ${formatDate(w.effectiveDate)}</span>
        <span>■ End: ${formatDate(w.endDate)}</span>
        ${daysLabel}
      </div>
    </div>
    <div class="card-actions">
      <button class="card-btn btn-pdf" data-id="${w.id}" title="View Certificate">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
        PDF
      </button>
      <button class="card-btn btn-edit" data-id="${w.id}" title="Edit">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
        Edit
      </button>
      <button class="card-btn btn-del" data-id="${w.id}" title="Delete">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a1 1 0 011-1h4a1 1 0 011 1v2"/></svg>
        Del
      </button>
    </div>
  `;
  return card;
}

function esc(s) {
  if (!s) return '';
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── CARD ACTIONS ───────────────────────────────────────────
document.addEventListener('click', e => {
  const pdfBtn  = e.target.closest('.btn-pdf');
  const editBtn = e.target.closest('.btn-edit');
  const delBtn  = e.target.closest('.btn-del');

  if (pdfBtn) {
    const w = warranties.find(x => x.id === pdfBtn.dataset.id);
    if (w) openPdfModal(w);
  }
  if (editBtn) {
    const w = warranties.find(x => x.id === editBtn.dataset.id);
    if (w) openForm(w);
  }
  if (delBtn) {
    if (confirm('Delete this warranty certificate?')) {
      warranties = warranties.filter(x => x.id !== delBtn.dataset.id);
      renderAll();
      showToast('Warranty deleted.', 'warn');
    }
  }
});

// ── PDF CERTIFICATE PREVIEW ────────────────────────────────
function openPdfModal(w) {
  pdfWarranty = w;
  modalTitle.textContent = w.projectName + ' — Certificate';
  modalContent.innerHTML = buildCertHTML(w);
  pdfModal.style.display = 'flex';
}

btnCloseModal.addEventListener('click', () => { pdfModal.style.display = 'none'; });
pdfModal.addEventListener('click', e => { if (e.target === pdfModal) pdfModal.style.display = 'none'; });

btnDownloadPDF.addEventListener('click', () => {
  if (!pdfWarranty) return;
  generatePDF(pdfWarranty);
});

function buildCertHTML(w) {
  return `
  <div class="cert-page" id="certPageContent">
    <!-- TOP HEADER -->
    <div class="cert-header-top">
      <div class="cert-logo-area">
        <svg class="cert-logo-icon" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
          <circle cx="24" cy="24" r="21" stroke="#E8201A" stroke-width="3"/>
          <path d="M15 24 L24 13 L33 24 L24 35 Z" fill="#E8201A"/>
          <circle cx="24" cy="24" r="5" fill="white"/>
        </svg>
        <div class="cert-brand-block">
          <div class="cert-brand-name">Guild Electric</div>
          <div class="cert-brand-sub">Limited</div>
        </div>
      </div>
      <div class="cert-contact">
        470 Midwest Road<br/>
        Toronto, Ontario<br/>
        Canada M1P 4Y5<br/>
        Tel: (416) 288-8222 &nbsp; Fax: (416) 288-9915<br/>
        Toll Free: (800) 387-6087<br/>
        <span class="cert-site">www.guildelectric.com</span>
      </div>
    </div>

    <!-- FIELDS TABLE -->
    <table class="cert-fields-table">
      <tr>
        <td class="field-label">Project Name / Contract No.</td>
        <td class="field-value">${esc(w.projectName)}</td>
        <td class="field-label">Date Issued</td>
        <td class="field-value">${formatDate(w.dateIssued)}</td>
      </tr>
      <tr>
        <td class="field-label">Project Address</td>
        <td class="field-value">${esc(w.projectAddress) || '—'}</td>
        <td class="field-label">Effective Date</td>
        <td class="field-value">${formatDate(w.effectiveDate)}</td>
      </tr>
      <tr>
        <td class="field-label">Issued To</td>
        <td class="field-value">${esc(w.issuedTo)}</td>
        <td class="field-label">Warranty Period</td>
        <td class="field-value warranty-period-val">${esc(w.warrantyPeriod)}</td>
      </tr>
      <tr>
        <td class="field-label">Subject</td>
        <td class="field-value" colspan="3">${esc(w.subject) || '—'}</td>
      </tr>
    </table>

    <!-- TITLE -->
    <div class="cert-title">Limited Warranty Certificate</div>

    <!-- INTRO -->
    <div class="cert-intro">
      Guild Electric Limited provides this Limited Warranty for the work described in the Subject line (the <strong>"Work"</strong>), as installed by Guild Electric and documented in the issued as-built drawings at the time of handover. This warranty shall commence upon Substantial Completion / Owner Acceptance of the Work on the Effective Date noted above.
    </div>

    <!-- SCOPE -->
    <div class="cert-section-head">Scope of Warranty</div>
    <div class="cert-body-text">
      The Work shall be free from defects in <strong>workmanship</strong> under normal use and service conditions for the period indicated above from the Effective Date. Equipment and materials are covered solely by the respective <strong>manufacturer's warranties</strong> and are not included within this Limited Warranty.
    </div>

    <!-- CLAIMS -->
    <div class="cert-section-head">Warranty Claims Procedure</div>
    <div class="cert-body-text">
      Warranty claims must be submitted in <strong>writing</strong> within the warranty period, and reasonable access must be provided for inspection. Guild Electric will, in accordance with applicable codes and project requirements, restore defective workmanship to the condition existing at the time of acceptance, by repair or replacement at its sole option, within a reasonable timeframe. Such remedy shall not require upgrades or improvements beyond the original accepted installation.
    </div>

    <!-- EXCLUSIONS -->
    <div class="cert-section-head">Exclusions &amp; Limitations</div>
    <ul class="cert-bullets">
      <li>Work modified, altered, or repaired by parties other than Guild Electric Limited.</li>
      <li>Damage caused by misuse, abuse, neglect, accident, or acts of God.</li>
      <li>Normal wear and tear consistent with the expected service life of the materials.</li>
      <li>Issues arising from design deficiencies or Owner-supplied equipment / materials.</li>
      <li>Failure to follow manufacturer's operating or maintenance instructions.</li>
    </ul>

    <div class="cert-closing">
      Liability under this warranty is strictly limited to the remedies described above and expressly excludes all consequential, incidental, or special damages, including but not limited to loss of use or revenue. This warranty is provided in accordance with the applicable project specifications and subcontract agreement requirements.
    </div>

    <div style="margin-top:12px; color:#333;">Yours truly,<br/><strong style="color:#E8201A;">Guild Electric Limited</strong></div>

    <!-- SIGNATURE BLOCK -->
    <div class="cert-sig-block">
      <div class="cert-sig-left">
        <div class="cert-sig-line"></div>
        <div class="cert-sig-name">Dan Camilleri</div>
        <div class="cert-sig-title">Chief Operating Officer</div>
        <div style="margin-top:12px; font-size:11px; color:#555;">
          <strong>Attachments / Notes:</strong><br/>
          ${esc(w.notes) || '— N/A'}
        </div>
      </div>
      <div class="cert-sig-right">
        <div class="cert-wo">W.O. # ${esc(w.wo) || '___________'}</div>
        <br/>
        <div class="cert-dist">
          <div class="cert-dist-label">Distribution:</div>
          <ul class="cert-dist-list">
            <li>L. Bertrand</li>
            <li>Shah</li>
            <li>Manual</li>
            <li>Service Department</li>
            <li>${w.pm ? esc(w.pm) + ' (PM)' : 'Project Manager'}</li>
          </ul>
        </div>
      </div>
    </div>

    <!-- FOOTER -->
    <div class="cert-footer">
      <span>Electrical Communication</span>
      <span>Communications</span>
      <span>Airports, Highways &amp; Traffic</span>
      <span>Maintenance &amp; Service</span>
      <span>Signs</span>
    </div>
  </div>`;
}

// ── PDF GENERATION ─────────────────────────────────────────
async function generatePDF(w) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'letter' });

  const pageW   = 215.9;
  const pageH   = 279.4;
  const margin  = 18;
  const contentW = pageW - margin * 2;
  let y = margin;

  // Colors
  const RED   = [232, 32, 26];
  const BLACK = [26, 26, 26];
  const GRAY  = [100, 100, 100];
  const LGRAY = [200, 200, 200];
  const BGRAY = [245, 245, 245];

  // ── HELPERS ──
  const setFont = (size, style = 'normal', color = BLACK) => {
    doc.setFontSize(size);
    doc.setFont('helvetica', style);
    doc.setTextColor(...color);
  };

  const addLine = (x1, y1, x2, y2, color = LGRAY, lw = 0.3) => {
    doc.setDrawColor(...color);
    doc.setLineWidth(lw);
    doc.line(x1, y1, x2, y2);
  };

  const addRect = (x, ry, rw, rh, fillColor, strokeColor, lw = 0.3) => {
    if (fillColor)   doc.setFillColor(...fillColor);
    if (strokeColor) doc.setDrawColor(...strokeColor);
    doc.setLineWidth(lw);
    if (fillColor && strokeColor) doc.rect(x, ry, rw, rh, 'FD');
    else if (fillColor)           doc.rect(x, ry, rw, rh, 'F');
    else                          doc.rect(x, ry, rw, rh, 'D');
  };

  // Wrap text and return new Y
  const wrapText = (text, x, yPos, maxW, lineH) => {
    const lines = doc.splitTextToSize(text || '', maxW);
    lines.forEach(line => { doc.text(line, x, yPos); yPos += lineH; });
    return yPos;
  };

  // ── FIX 1: REAL GUILD LOGO ──
  // Logo image is 28mm wide × 28mm tall in the header, left-aligned
  const logoH = 18;
  const logoW = 18; // roughly square crop of the logo
  if (typeof GUILD_LOGO_B64 !== 'undefined') {
    doc.addImage(GUILD_LOGO_B64, 'JPEG', margin, y, logoW, logoH);
  }
  const logoRight = margin + logoW + 3;

  // Contact block — right side
  setFont(8, 'normal', GRAY);
  const contactLines = [
    '470 Midwest Road',
    'Toronto, Ontario  Canada M1P 4Y5',
    'Tel: (416) 288-8222   Fax: (416) 288-9915',
    'Toll Free: (800) 387-6087'
  ];
  let cy2 = y + 2;
  contactLines.forEach(l => {
    doc.text(l, pageW - margin, cy2, { align: 'right' });
    cy2 += 4;
  });
  setFont(8, 'bold', RED);
  doc.text('www.guildelectric.com', pageW - margin, cy2, { align: 'right' });

  y += logoH + 4;
  addLine(margin, y, pageW - margin, y, RED, 0.6);
  y += 6;

  // ── FIELDS TABLE ──
  // FIX 3: Taller rows (10mm) to allow text to wrap without overflow
  const col1W = 46, col2W = 68, col3W = 38, col4W = contentW - col1W - col2W - col3W;
  const rowH  = 10;
  const labelSize = 7;
  const valueSize = 8; // FIX 4: smaller value font so text fits in cell

  const colX = [margin, margin + col1W, margin + col1W + col2W, margin + col1W + col2W + col3W];

  // Helper: draw one field row
  const drawRow = (label1, val1, label2, val2, boldVal1 = false) => {
    addRect(colX[0], y, col1W,  rowH, BGRAY, LGRAY);
    addRect(colX[1], y, col2W,  rowH, null,  LGRAY);
    addRect(colX[2], y, col3W,  rowH, BGRAY, LGRAY);
    addRect(colX[3], y, col4W,  rowH, null,  LGRAY);

    // Labels
    setFont(labelSize, 'bold', GRAY);
    doc.text(label1 || '', colX[0] + 2, y + 7);
    doc.text(label2 || '', colX[2] + 2, y + 7);

    // FIX 2 & 5: Value col 1 — truncate to fit cell; bold if flagged
    setFont(valueSize, boldVal1 ? 'bold' : 'normal', BLACK);
    const v1lines = doc.splitTextToSize(val1 || '\u2014', col2W - 4);
    doc.text(v1lines[0], colX[1] + 2, y + 7); // only first line (single row)

    // Value col 2 — truncate to fit
    setFont(valueSize, 'normal', BLACK);
    const v2lines = doc.splitTextToSize(val2 || '\u2014', col4W - 4);
    doc.text(v2lines[0], colX[3] + 2, y + 7);

    y += rowH;
  };

  // FIX 5: Project name is bold (boldVal1 = true on first row)
  drawRow('PROJECT NAME / CONTRACT NO.', w.projectName,    'DATE ISSUED',     formatDate(w.dateIssued),    true);
  drawRow('PROJECT ADDRESS',             w.projectAddress, 'EFFECTIVE DATE',  formatDate(w.effectiveDate), false);
  drawRow('ISSUED TO',                   w.issuedTo,       'WARRANTY PERIOD', w.warrantyPeriod,            false);

  // Subject row — full width, may need taller cell if text is long
  const subjectText   = w.subject || '\u2014';
  const subjectLines  = doc.splitTextToSize(subjectText, contentW - col1W - 4);
  const subjectRowH   = Math.max(rowH, subjectLines.length * 4.5 + 4);
  addRect(colX[0], y, col1W,            subjectRowH, BGRAY, LGRAY);
  addRect(colX[1], y, contentW - col1W, subjectRowH, null,  LGRAY);
  setFont(labelSize, 'bold', GRAY);
  doc.text('SUBJECT', colX[0] + 2, y + 7);
  setFont(valueSize, 'normal', BLACK);
  subjectLines.slice(0, 2).forEach((line, i) => {   // max 2 lines
    doc.text(line, colX[1] + 2, y + 5 + i * 4.5);
  });
  y += subjectRowH + 5;

  // ── CERTIFICATE TITLE ──
  addLine(margin, y, pageW - margin, y, RED, 0.8);
  y += 1.5;
  setFont(17, 'bold', RED);
  doc.text('LIMITED WARRANTY CERTIFICATE', pageW / 2, y + 6, { align: 'center' });
  y += 8;
  addLine(margin, y, pageW - margin, y, RED, 0.8);
  y += 6;

  // ── BODY TEXT (FIX 3: size 8 instead of 9 to prevent footer overlap) ──
  const bodySize = 8;
  const lineH    = 4.5;
  const sectionGap = 3;

  setFont(bodySize, 'normal', BLACK);
  y = wrapText(
    'Guild Electric Limited provides this Limited Warranty for the work described in the Subject line (the \u201cWork\u201d), as installed by Guild Electric and documented in the issued as-built drawings at the time of handover. This warranty shall commence upon Substantial Completion / Owner Acceptance of the Work on the Effective Date noted above.',
    margin, y, contentW, lineH
  );
  y += sectionGap;

  // Section block helper
  const section = (title, body) => {
    setFont(9, 'bold', RED);
    doc.text(title.toUpperCase(), margin, y);
    y += 5;
    setFont(bodySize, 'normal', BLACK);
    y = wrapText(body, margin, y, contentW, lineH);
    y += sectionGap;
  };

  section('Scope of Warranty',
    'The Work shall be free from defects in workmanship under normal use and service conditions for the period indicated above from the Effective Date. Equipment and materials are covered solely by the respective manufacturer\u2019s warranties and are not included within this Limited Warranty.'
  );

  section('Warranty Claims Procedure',
    'Warranty claims must be submitted in writing within the warranty period, and reasonable access must be provided for inspection. Guild Electric will, in accordance with applicable codes and project requirements, restore defective workmanship to the condition existing at the time of acceptance, by repair or replacement at its sole option, within a reasonable timeframe. Such remedy shall not require upgrades or improvements beyond the original accepted installation.'
  );

  // Exclusions heading
  setFont(9, 'bold', RED);
  doc.text('EXCLUSIONS & LIMITATIONS', margin, y);
  y += 5;

  // FIX 2: Use a proper bullet via latin1 character (ASCII 149 = bullet in win-1252,
  // but jsPDF standard fonts support the en-dash reliably; use unicode \u2022 bullet)
  setFont(bodySize, 'normal', BLACK);
  const bullets = [
    'Work modified, altered, or repaired by parties other than Guild Electric Limited.',
    'Damage caused by misuse, abuse, neglect, accident, or acts of God.',
    'Normal wear and tear consistent with the expected service life of the materials.',
    'Issues arising from design deficiencies or Owner-supplied equipment / materials.',
    'Failure to follow manufacturer\u2019s operating or maintenance instructions.'
  ];
  const bulletIndent = 5;
  bullets.forEach(b => {
    // Draw solid filled circle as bullet — reliable across all PDF viewers
    doc.setFillColor(...BLACK);
    doc.circle(margin + 1.2, y - 1.2, 0.9, 'F');
    const wrapped = doc.splitTextToSize(b, contentW - bulletIndent);
    wrapped.forEach((line, i) => {
      doc.text(line, margin + bulletIndent, y + (i * lineH));
    });
    y += wrapped.length * lineH + 1;
  });
  y += 2;

  setFont(bodySize, 'normal', BLACK);
  y = wrapText(
    'Liability under this warranty is strictly limited to the remedies described above and expressly excludes all consequential, incidental, or special damages, including but not limited to loss of use or revenue. This warranty is provided in accordance with the applicable project specifications and subcontract agreement requirements.',
    margin, y, contentW, lineH
  );
  y += 4;

  // ── CLOSING & SIGNATURE ──
  // Reserve fixed space at bottom for sig block + footer
  // Footer occupies bottom 18mm; sig block needs ~50mm; attachments ~15mm
  // So body must not go below pageH - 18 - 50 - 15 = ~196mm
  // We clamp y here — content is sized (bodySize=8) to fit comfortably

  setFont(8, 'normal', GRAY);
  doc.text('Yours truly,', margin, y);
  y += 4.5;
  setFont(9, 'bold', RED);
  doc.text('Guild Electric Limited', margin, y);
  y += 12;

  addLine(margin, y, margin + 58, y, [26, 26, 26], 0.4);
  y += 4.5;
  setFont(9, 'bold', BLACK);
  doc.text('Dan Camilleri', margin, y);
  y += 4.5;
  setFont(8, 'normal', GRAY);
  doc.text('Chief Operating Officer', margin, y);
  y += 6;

  setFont(8, 'bold', GRAY);
  doc.text('Attachments / Notes:', margin, y);
  y += 4.5;
  setFont(8, 'normal', BLACK);
  const notesLines = doc.splitTextToSize(w.notes || '— N/A', 110);
  notesLines.forEach(l => { doc.text(l, margin, y); y += 4; });

  // ── W.O. + DISTRIBUTION (fixed at bottom-right, above footer) ──
  const distBaseY = pageH - 20 - (5 * 4.5) - 14; // anchored from bottom up
  setFont(9, 'bold', RED);
  doc.text(`W.O. # ${w.wo || '___________'}`, pageW - margin, distBaseY, { align: 'right' });

  setFont(8, 'bold', GRAY);
  doc.text('Distribution:', pageW - margin, distBaseY + 6, { align: 'right' });
  const distList = ['L. Bertrand', 'Shah', 'Manual', 'Service Department',
                    w.pm ? `${w.pm} (PM)` : 'Project Manager'];
  setFont(8, 'normal', GRAY);
  distList.forEach((d, i) => {
    doc.text(`\u2014 ${d}`, pageW - margin, distBaseY + 11 + i * 4.5, { align: 'right' });
  });

  // ── FOOTER — always pinned to bottom ──
  const fy = pageH - 10;
  addLine(margin, fy - 4, pageW - margin, fy - 4, RED, 0.8);
  setFont(7, 'bold', RED);
  const footItems = ['Electrical Communication', 'Communications', 'Airports, Highways & Traffic', 'Maintenance & Service', 'Signs'];
  const spacing = contentW / footItems.length;
  footItems.forEach((item, i) => {
    doc.text(item, margin + spacing * i + spacing / 2, fy, { align: 'center' });
  });

  // ── SAVE ──
  const filename = `Guild_Warranty_${w.projectName.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 40)}_${w.dateIssued}.pdf`;
  doc.save(filename);
  showToast('PDF downloaded.');
}

// ── INIT ───────────────────────────────────────────────────
(function init() {
  fDateIssued.value    = today();
  fEffectiveDate.value = today();
  formPanel.classList.add('hidden');
  renderAll();
})();
