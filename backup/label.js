/* ─── State ─────────────────────────────────────────── */
let parsedData = [], columns = [], fields = [], selectedId = null, pdfDoc = null;
let zoomLevel = 1, previewMode = false, previewRow = 0;
let cw = 0, ch = 0;
const MM2PX = 3.7795;
// CANVAS_DPR is computed per-size in applySize() — small labels need higher DPR
let CANVAS_DPR = 2;
let bgImage = null; // Optional background/reference image on the canvas

/* ─── Barcode / QR caches ───────────────────────────── */
// Each entry: { img, val, w, h }  — keyed by field id
const bcCache = {};   // barcode canvas cache
const qrCache = {};   // qr image cache (keyed by val|w|h)

/* ─── Step navigation ───────────────────────────────── */
function goStep(n) {
  [1, 2, 3].forEach(i => {
    document.getElementById('step' + i).style.display = i === n ? 'block' : 'none';
    const t = document.getElementById('tab' + i);
    t.classList.toggle('active', i === n);
    t.classList.toggle('done', i < n);
  });
  if (n === 2) initDesigner();
  if (n === 3) {
    const wmm = parseFloat(document.getElementById('lwmm').value) || 80;
    const hmm = parseFloat(document.getElementById('lhmm').value) || 40;
    document.getElementById('s3info').textContent =
      parsedData.length + ' labels at ' + wmm + 'W × ' + hmm + 'H mm each.';
  }
}

/* ─── Step 1 ─────────────────────────────────────────── */
const dropZone = document.getElementById('dropZone');
document.getElementById('fileInput').addEventListener('change', e => handleFile(e.target.files[0]));
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag');
  handleFile(e.dataTransfer.files[0]);
});

function handleFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (!json.length) { alert('No data found.'); return; }
    parsedData = json;
    columns = Object.keys(json[0]);
    renderDataPreview();
    document.getElementById('toStep2').style.display = 'inline-flex';
  };
  reader.readAsArrayBuffer(file);
}

function renderDataPreview() {
  document.getElementById('dataPreviewWrap').style.display = 'block';
  document.getElementById('rowCount').textContent = parsedData.length + ' rows';
  document.getElementById('colCount').textContent = columns.length + ' columns';
  let h = '<table><tr>' + columns.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
  parsedData.slice(0, 5).forEach(r => {
    h += '<tr>' + columns.map(c => '<td>' + esc(String(r[c] ?? '')) + '</td>').join('') + '</tr>';
  });
  document.getElementById('dataPreview').innerHTML = h + '</table>';
}

/* ─── Step 2 ─────────────────────────────────────────── */
function initDesigner() {
  const sel = document.getElementById('colSelect');
  sel.innerHTML = columns.map(c => '<option>' + esc(c) + '</option>').join('');
  renderColChips();
  applySize();
}

function getUsedCols() { return new Set(fields.map(f => f.col)); }

function renderColChips() {
  const used = getUsedCols();
  document.getElementById('colChips').innerHTML = columns.map(c => {
    const u = used.has(c);
    return `<span class="col-chip${u ? ' used' : ''}" onclick="quickAdd('${esc(c)}')">${u ? '<span style="color:#3B6D11;font-size:10px">✓</span> ' : ''}${esc(c)}</span>`;
  }).join('');
  const sel = document.getElementById('colSelect');
  sel.innerHTML = columns.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join('');
  if (columns.length) sel.value = columns[0];
}

function onSizePreset() {
  const v = document.getElementById('sizePreset').value;
  const map = { A4: [210, 297], A5: [148, 210], label_50x25: [50, 25], label_80x40: [80, 40], label_100x50: [100, 50] };
  if (map[v]) {
    document.getElementById('lwmm').value = map[v][0];
    document.getElementById('lhmm').value = map[v][1];
  }
  document.getElementById('customSizeFields').style.opacity = v === 'custom' ? '1' : '0.6';
  applySize();
}

function applySize() {
  const wmm = parseFloat(document.getElementById('lwmm').value) || 80;
  const hmm = parseFloat(document.getElementById('lhmm').value) || 40;
  cw = Math.round(wmm * MM2PX);
  ch = Math.round(hmm * MM2PX);

  // Higher DPR for small labels (they get zoomed in more → need more pixels)
  // Capped so large labels don't use excessive memory
  const pixelArea = cw * ch;
  const screenDPR = window.devicePixelRatio || 1;
  if (pixelArea < 60000)       CANVAS_DPR = 4;   // ≤ ~100×60mm labels
  else if (pixelArea < 250000) CANVAS_DPR = 3;   // ≤ ~130×100mm
  else if (pixelArea < 700000) CANVAS_DPR = 2;   // up to ~A4
  else                         CANVAS_DPR = Math.max(1, Math.min(2, screenDPR));

  const canvas = document.getElementById('labelCanvas');
  // Internal buffer at DPR resolution; CSS size stays at logical pixels → sharp at all zoom levels
  canvas.width = Math.round(cw * CANVAS_DPR);
  canvas.height = Math.round(ch * CANVAS_DPR);
  canvas.style.width = cw + 'px';
  canvas.style.height = ch + 'px';
  const inner = document.getElementById('canvasInner');
  inner.style.width = cw + 'px'; inner.style.height = ch + 'px';
  document.getElementById('fieldsLayer').style.width = cw + 'px';
  document.getElementById('fieldsLayer').style.height = ch + 'px';
  document.getElementById('canvasInfo').textContent = wmm + 'W × ' + hmm + 'H mm';
  // Invalidate image caches — size changed
  Object.keys(bcCache).forEach(k => delete bcCache[k]);
  Object.keys(qrCache).forEach(k => delete qrCache[k]);
  // Update PDF filename with size suffix
  const pdfNameEl = document.getElementById('pdfName');
  if (pdfNameEl) {
    const base = pdfNameEl.value.replace(/-\d+x\d+mm$/, '') || 'seznik-labels';
    pdfNameEl.value = `${base}-${wmm}x${hmm}mm`;
  }
  zoomFit();
  drawCanvas();
}

function zoom(d) {
  zoomLevel = Math.max(0.15, Math.min(3, +(zoomLevel + d).toFixed(2)));
  applyZoom();
}

function zoomFit() {
  const s = document.getElementById('canvasScroll');
  const maxZ = 3;
  zoomLevel = +Math.min((s.clientWidth - 48) / cw, (s.clientHeight - 48) / ch, maxZ).toFixed(2);
  zoomLevel = Math.max(0.1, zoomLevel);
  applyZoom();
}

function applyZoom() {
  const inner = document.getElementById('canvasInner');
  inner.style.transform = 'scale(' + zoomLevel + ')';
  inner.style.transformOrigin = 'top left';
  document.getElementById('canvasScroll').style.minHeight =
    Math.max(300, Math.round(ch * zoomLevel) + 48) + 'px';
  document.getElementById('zoomLabel').textContent = Math.round(zoomLevel * 100) + '%';
}

/* ─── Fields ─────────────────────────────────────────── */
function quickAdd(col) {
  document.getElementById('colSelect').value = col;
  addField('text');
}

function addField(type) {
  const col = document.getElementById('colSelect').value;
  if (!col) return;
  // Determine index (which record offset this field should map to).
  // E.g., if this is the 3rd time adding "Name", it gets index 2 (shows 3rd record on the page)
  let idx = 0;
  fields.forEach(f => { if (f.col === col && f.index >= idx) idx = f.index + 1; });
  
  const f = {
    id: 'f' + Date.now(), col, type,
    x: 20, y: 20 + (fields.length * 32),
    w: type === 'qr' ? 80 : type === 'barcode' ? 140 : 120,
    h: type === 'qr' ? 80 : type === 'barcode' ? 56 : 26,
    fontSize: 14, fontFamily: 'system-ui',
    bold: false, italic: false, underline: false, strikethrough: false,
    align: 'left', prefix: '', bcText: true, rotation: 0,
    index: idx
  };
  fields.push(f);
  createFieldEl(f);
  selectField(f.id);
  renderColChips();
  drawCanvas();
}

function createFieldEl(f) {
  const layer = document.getElementById('fieldsLayer');
  const el = document.createElement('div');
  el.className = 'field-el';
  el.id = 'fe_' + f.id;
  updateElTransform(el, f);
  el.innerHTML = `
    <button class="del-btn" onclick="event.stopPropagation();deleteField('${f.id}')">×</button>
    <div class="f-inner" id="fi_${f.id}"></div>
    <div class="res-handle" id="rh_${f.id}"></div>`;
  layer.appendChild(el);

  el.addEventListener('pointerdown', e => {
    if (e.target.classList.contains('res-handle') ||
      e.target.classList.contains('del-btn')) return;
    selectField(f.id);
    startDrag(e, f, el);
  });
  document.getElementById('rh_' + f.id).addEventListener('pointerdown', e => {
    e.preventDefault(); e.stopPropagation();
    startResize(e, f, el);
  });
}

function updateElTransform(el, f) {
  el.style.cssText = `left:${f.x}px;top:${f.y}px;width:${f.w}px;height:${f.h}px;position:absolute;transform:rotate(${f.rotation || 0}deg);transform-origin:center center;`;
}

// Returns the axis-aligned bounding box size of a rotated field
function getRotatedBBox(f) {
  const angle = (f.rotation || 0) * Math.PI / 180;
  const cos = Math.abs(Math.cos(angle));
  const sin = Math.abs(Math.sin(angle));
  return { bw: f.w * cos + f.h * sin, bh: f.w * sin + f.h * cos };
}

function startDrag(e0, f, el) {
  // Capture start position so delta is computed in canvas-coordinate space
  const startX = f.x, startY = f.y;
  const mx0 = e0.clientX, my0 = e0.clientY;
  function onMove(e) {
    const rawX = startX + (e.clientX - mx0) / zoomLevel;
    const rawY = startY + (e.clientY - my0) / zoomLevel;
    // Clamp using the rotated bounding box so the field can reach all edges
    const { bw, bh } = getRotatedBBox(f);
    f.x = Math.max(bw / 2 - f.w / 2, Math.min(cw - f.w / 2 - bw / 2, rawX));
    f.y = Math.max(bh / 2 - f.h / 2, Math.min(ch - f.h / 2 - bh / 2, rawY));
    el.style.left = f.x + 'px';
    el.style.top = f.y + 'px';
    drawCanvas();
  }
  function onUp() {
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
  }
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
}

function startResize(e0, f, el) {
  const rw = f.w, rh = f.h, rx = e0.clientX, ry = e0.clientY;
  function onMove(e) {
    f.w = Math.max(24, rw + (e.clientX - rx) / zoomLevel);
    f.h = Math.max(14, rh + (e.clientY - ry) / zoomLevel);
    el.style.width = f.w + 'px';
    el.style.height = f.h + 'px';
    // Invalidate barcode cache so it re-renders at new size
    delete bcCache[f.id];
    drawCanvas();
  }
  function onUp() {
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
  }
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
}



function selectField(id) {
  selectedId = id;
  document.querySelectorAll('.field-el').forEach(e => e.classList.remove('selected'));
  const el = document.getElementById('fe_' + id);
  if (el) el.classList.add('selected');
  const f = fields.find(x => x.id === id);
  if (!f) return;
  document.getElementById('noSel').style.display = 'none';
  document.getElementById('selProps').style.display = 'block';
  document.getElementById('propFieldName').textContent = f.col + ' — ' + f.type;
  ['ptText', 'ptBarcode', 'ptQR'].forEach(x => document.getElementById(x).classList.remove('active'));
  document.getElementById(f.type === 'qr' ? 'ptQR' : f.type === 'barcode' ? 'ptBarcode' : 'ptText').classList.add('active');
  document.getElementById('textProps').style.display = f.type === 'text' ? 'block' : 'none';
  document.getElementById('barcodeProps').style.display = f.type === 'barcode' ? 'block' : 'none';
  document.getElementById('propFontSize').value = f.fontSize;
  document.getElementById('propFontFamily').value = f.fontFamily || 'system-ui';
  document.getElementById('propPrefix').value = f.prefix || '';
  document.getElementById('propBcText').value = f.bcText ? '1' : '0';
  ['bold', 'italic', 'underline', 'strikethrough'].forEach(s =>
    document.getElementById('sb-' + s).classList.toggle('active', !!f[s]));
  ['left', 'center', 'right'].forEach(a =>
    document.getElementById('ab-' + a).classList.toggle('active', f.align === a));
  const rot = Math.round(f.rotation || 0);
  document.getElementById('propRotation').value = rot;
  document.getElementById('rotVal').textContent = rot + '°';
}

function updateFontFamily() {
  if (!selectedId) return;
  const f = fields.find(x => x.id === selectedId);
  if (!f) return;
  f.fontFamily = document.getElementById('propFontFamily').value;
  drawCanvas();
}

/* ─── Template Save / Import ─────────────────────────── */
function saveTemplate() {
  const wmm = parseFloat(document.getElementById('lwmm').value) || 80;
  const hmm = parseFloat(document.getElementById('lhmm').value) || 40;
  const blob = new Blob([JSON.stringify({ version: 1, wmm, hmm, fields }, null, 2)],
    { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `template-${wmm}x${hmm}mm.json`;
  a.click();
}

function importTemplate(file) {
  if (!file) return;
  const isImage = file.type.startsWith('image/');

  if (isImage) {
    // Load as a background/reference image on the canvas
    const reader = new FileReader();
    reader.onload = e => {
      const img = new Image();
      img.onload = () => {
        bgImage = img;
        // Auto-set canvas size to match image aspect ratio (keep width, adjust height)
        const wmm = parseFloat(document.getElementById('lwmm').value) || 80;
        const ratio = img.naturalHeight / img.naturalWidth;
        const hmm = Math.round(wmm * ratio * 10) / 10;
        document.getElementById('lhmm').value = hmm;
        applySize();
        document.getElementById('clearBgBtn').style.display = 'inline-flex';
        drawCanvas();
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
    return;
  }

  // JSON template
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const t = JSON.parse(e.target.result);
      if (t.wmm) document.getElementById('lwmm').value = t.wmm;
      if (t.hmm) document.getElementById('lhmm').value = t.hmm;
      applySize();
      // Clear existing fields
      fields.forEach(f => { const el = document.getElementById('fe_' + f.id); if (el) el.remove(); });
      fields = []; selectedId = null;
      document.getElementById('noSel').style.display = 'block';
      document.getElementById('selProps').style.display = 'none';
      // Restore saved fields
      if (Array.isArray(t.fields)) {
        t.fields.forEach(saved => {
          const nf = { ...saved, id: 'f' + Date.now() + Math.round(Math.random() * 1e6) };
          fields.push(nf);
          createFieldEl(nf);
        });
      }
      renderColChips();
      drawCanvas();
    } catch { alert('Invalid template file.'); }
  };
  reader.readAsText(file);
}

function clearBackground() {
  bgImage = null;
  document.getElementById('clearBgBtn').style.display = 'none';
  drawCanvas();
}

function changePropType(type) {
  if (!selectedId) return;
  const f = fields.find(x => x.id === selectedId);
  if (!f) return;
  f.type = type;
  if (type === 'qr') { f.w = Math.max(f.w, 60); f.h = Math.max(f.h, 60); }
  const el = document.getElementById('fe_' + f.id);
  if (el) { el.style.width = f.w + 'px'; el.style.height = f.h + 'px'; }
  const fi = document.getElementById('fi_' + f.id);
  if (fi) fi.innerHTML = '';
  // Invalidate barcode cache when type changes
  delete bcCache[f.id];
  selectField(f.id);
  drawCanvas();
}

function toggleStyle(s) {
  if (!selectedId) return;
  const f = fields.find(x => x.id === selectedId);
  if (!f) return;
  f[s] = !f[s];
  document.getElementById('sb-' + s).classList.toggle('active', f[s]);
  drawCanvas();
}

function setAlign(a) {
  if (!selectedId) return;
  const f = fields.find(x => x.id === selectedId);
  if (!f) return;
  f.align = a;
  ['left', 'center', 'right'].forEach(x =>
    document.getElementById('ab-' + x).classList.toggle('active', x === a));
  drawCanvas();
}

function updateProp() {
  if (!selectedId) return;
  const f = fields.find(x => x.id === selectedId);
  if (!f) return;
  f.fontSize = parseInt(document.getElementById('propFontSize').value) || 14;
  f.prefix = document.getElementById('propPrefix').value;
  const newBcText = document.getElementById('propBcText').value === '1';
  if (f.bcText !== newBcText) {
    f.bcText = newBcText;
    delete bcCache[f.id]; // Invalidate cache when displayValue changes
  } else {
    f.bcText = newBcText;
  }
  drawCanvas();
}

function updateRotation() {
  if (!selectedId) return;
  const f = fields.find(x => x.id === selectedId);
  if (!f) return;
  f.rotation = parseInt(document.getElementById('propRotation').value) || 0;
  document.getElementById('rotVal').textContent = f.rotation + '°';
  const el = document.getElementById('fe_' + f.id);
  if (el) updateElTransform(el, f);
  drawCanvas();
}

function setRotation(deg) {
  if (!selectedId) return;
  document.getElementById('propRotation').value = deg;
  updateRotation();
}

function deleteSelected() { if (selectedId) deleteField(selectedId); }

function deleteField(id) {
  fields = fields.filter(f => f.id !== id);
  delete bcCache[id];
  const el = document.getElementById('fe_' + id);
  if (el) el.remove();
  if (selectedId === id) {
    selectedId = null;
    document.getElementById('noSel').style.display = 'block';
    document.getElementById('selProps').style.display = 'none';
  }
  renderColChips();
  drawCanvas();
}

/* ─── Preview ─────────────────────────────────────────── */
function togglePreview() {
  previewMode = !previewMode;
  previewRow = 0;
  document.getElementById('previewToggle').textContent = previewMode ? '✏ Edit Mode' : '👁 Preview Data';
  document.getElementById('previewNav').style.display = previewMode ? 'flex' : 'none';
  document.getElementById('previewTableWrap').style.display = previewMode ? 'block' : 'none';
  document.getElementById('previewStatus').textContent = previewMode ? 'Previewing real data' : 'Design mode';
  if (previewMode) renderPreviewTable();
  updateRowIndicator();
  drawCanvas();
}

function renderPreviewTable() {
  let h = '<table><tr><th>#</th>' + columns.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
  parsedData.forEach((r, i) => {
    h += `<tr class="${i === previewRow ? 'active-row' : ''}" onclick="jumpToRow(${i})"><td style="color:#aaa;font-size:10px">${i + 1}</td>`;
    h += columns.map(c => '<td>' + esc(String(r[c] ?? '')) + '</td>').join('') + '</tr>';
  });
  document.getElementById('previewDataTable').innerHTML = h + '</table>';
  setTimeout(() => {
    const ar = document.querySelector('#previewDataTable tr.active-row');
    if (ar) ar.scrollIntoView({ block: 'nearest' });
  }, 50);
}

function jumpToRow(i) { previewRow = i; renderPreviewTable(); updateRowIndicator(); drawCanvas(); }
function prevRow() { if (previewRow > 0) { previewRow--; renderPreviewTable(); updateRowIndicator(); drawCanvas(); } }
function nextRow() { if (previewRow < parsedData.length - 1) { previewRow++; renderPreviewTable(); updateRowIndicator(); drawCanvas(); } }
function updateRowIndicator() {
  document.getElementById('rowIndicator').textContent =
    previewMode ? 'Row ' + (previewRow + 1) + ' / ' + parsedData.length : '';
}

/* ─── Draw canvas ─────────────────────────────────────── */
// Map CSS font family name to jsPDF built-in font
function jsPdfFont(fontFamily) {
  const f = (fontFamily || '').toLowerCase();
  if (f.includes('georgia') || f.includes('times')) return 'times';
  if (f.includes('courier')) return 'courier';
  return 'helvetica';
}

function buildFont(f) {
  const family = f.fontFamily || 'system-ui';
  return (f.italic ? 'italic ' : '') + (f.bold ? 'bold ' : '') + `${f.fontSize || 14}px ${family},sans-serif`;
}

function drawCanvas() {
  const canvas = document.getElementById('labelCanvas');
  const ctx = canvas.getContext('2d');
  // Clear the full physical buffer
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  // Scale all drawing to DPR so coordinates stay in logical pixels
  ctx.save();
  ctx.scale(CANVAS_DPR, CANVAS_DPR);
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cw, ch);
  // Draw background image (reference template) if loaded
  if (bgImage) {
    ctx.globalAlpha = 0.55; // Semi-transparent so fields are readable on top
    ctx.drawImage(bgImage, 0, 0, cw, ch);
    ctx.globalAlpha = 1;
  }
  fields.forEach(f => {
    // Map to the correct record based on the field's index slot
    let val = '[' + f.col + (f.index > 0 ? ` #${f.index+1}` : '') + ']';
    if (previewMode && parsedData.length) {
      const targetRowIndex = previewRow + (f.index || 0);
      const rowData = parsedData[targetRowIndex];
      val = rowData ? String(rowData[f.col] ?? '') : ''; // blank if we run out of records
    }
    const cx = f.x + f.w / 2, cy = f.y + f.h / 2;
    const rot = (f.rotation || 0) * Math.PI / 180;
    ctx.save();
    ctx.translate(cx, cy);
    ctx.rotate(rot);
    ctx.translate(-f.w / 2, -f.h / 2);
    if (f.type === 'text') {
      drawText(ctx, f, val);
    } else if (f.type === 'barcode') {
      drawBarcode(ctx, f, val);
    } else if (f.type === 'qr') {
      drawQR(ctx, f, val);
    }
    ctx.restore();
  });
  ctx.restore();
}

function drawText(ctx, f, val) {
  const text = (f.prefix || '') + val;
  ctx.fillStyle = '#1a1a1a';
  ctx.font = buildFont(f);
  ctx.textBaseline = 'top';
  ctx.textAlign = f.align || 'left';
  const lh = (f.fontSize || 14) * 1.35;
  const lines = wrapTextLines(ctx, text, f.w);
  const ax = f.align === 'center' ? f.w / 2 : f.align === 'right' ? f.w : 0;
  lines.forEach((line, li) => {
    const ty = li * lh;
    ctx.fillText(line, ax, ty);
    if (f.underline) {
      ctx.strokeStyle = '#1a1a1a'; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(0, ty + f.fontSize + 1); ctx.lineTo(f.w, ty + f.fontSize + 1); ctx.stroke();
    }
    if (f.strikethrough) {
      ctx.strokeStyle = '#1a1a1a'; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(0, ty + f.fontSize / 2); ctx.lineTo(f.w, ty + f.fontSize / 2); ctx.stroke();
    }
  });
  ctx.textAlign = 'left';
}

/* ── FIX: barcode uses a cache so drawImage runs synchronously
         inside the correct ctx.save/restore block. ──────────────
   When the cache has a valid entry for this field+val+size, we
   draw it immediately.  Otherwise we kick off an async render,
   store the result, and call drawCanvas() again when ready.     */
function drawBarcode(ctx, f, val) {
  const entry = bcCache[f.id];
  // Cache hit: same value and dimensions → draw synchronously
  if (entry && entry.val === val && entry.w === f.w && entry.h === f.h) {
    ctx.drawImage(entry.img, 0, 0, f.w, f.h);
    return;
  }
  // Show placeholder while loading
  ctx.fillStyle = '#e0f0e0';
  ctx.fillRect(0, 0, f.w, f.h);
  ctx.fillStyle = '#3B6D11';
  ctx.font = '10px monospace';
  ctx.textBaseline = 'middle';
  ctx.fillText('▦ ' + (val || 'barcode'), 4, f.h / 2);
  ctx.textBaseline = 'top';

  // Async render into off-screen canvas → cache → redraw
  const svgEl = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  try {
    JsBarcode(svgEl, val || '0', {
      format: 'CODE128',
      width: 2,
      height: Math.max(10, Math.round(f.h - 20)),
      displayValue: f.bcText,
      fontSize: 12,
      textMargin: 4,
      margin: 4,
      background: '#ffffff',
      lineColor: '#000000'
    });
    const svgStr = new XMLSerializer().serializeToString(svgEl);
    const img = new Image();
    img.onload = () => {
      // Store in cache with the dimensions used
      bcCache[f.id] = { img, val, w: f.w, h: f.h };
      drawCanvas(); // redraw — this time the cache hit path runs synchronously ✓
    };
    img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgStr)));
  } catch (e) {
    ctx.fillStyle = '#888';
    ctx.font = '10px monospace';
    ctx.fillText('[BC:' + val + ']', 0, 12);
  }
}

function drawQR(ctx, f, val) {
  const k = val + '|' + Math.round(f.w) + '|' + Math.round(f.h);
  if (qrCache[k]) { ctx.drawImage(qrCache[k], 0, 0, f.w, f.h); return; }
  const div = document.createElement('div');
  div.style.cssText = 'position:fixed;left:-9999px;top:-9999px';
  document.body.appendChild(div);
  try {
    new QRCode(div, { text: val || ' ', width: Math.round(f.w), height: Math.round(f.h), correctLevel: QRCode.CorrectLevel.M });
    setTimeout(() => {
      const el = div.querySelector('canvas') || div.querySelector('img');
      if (el) {
        const i = new Image();
        i.onload = () => { qrCache[k] = i; drawCanvas(); };
        i.src = el.tagName === 'CANVAS' ? el.toDataURL() : el.src;
      }
      document.body.removeChild(div);
    }, 80);
  } catch (e) { document.body.removeChild(div); }
}

function wrapTextLines(ctx, text, maxW) {
  const words = text.split(' ');
  const lines = [];
  let line = '';
  for (let i = 0; i < words.length; i++) {
    const t = line + (line ? ' ' : '') + words[i];
    if (ctx.measureText(t).width > maxW && line) { lines.push(line); line = words[i]; }
    else line = t;
  }
  if (line) lines.push(line);
  return lines.length ? lines : [''];
}

/* ─── PDF Generation (WYSIWYG High-DPI Canvas) ───────── */
async function drawPageToPdf(doc, startRowIdx, pgW, pgH) {
  return new Promise(async resolve => {
    // Render at 4x resolution for crystal clear print quality
    const c = document.createElement('canvas');
    c.width = cw * 4;
    c.height = ch * 4;
    const ctx = c.getContext('2d');
    ctx.scale(4, 4);
    
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, cw, ch);
    
    // Fill background exactly like visual canvas
    if (bgImage) {
      ctx.drawImage(bgImage, 0, 0, cw, ch);
    }

    for (const f of fields) {
      const targetRowIndex = startRowIdx + (f.index || 0);
      if (targetRowIndex >= parsedData.length) continue;
      const val = String(parsedData[targetRowIndex][f.col] ?? '');
      if (!val) continue;

      const cx = f.x + f.w / 2, cy = f.y + f.h / 2;
      const rot = (f.rotation || 0) * Math.PI / 180;
      
      ctx.save();
      ctx.translate(cx, cy);
      ctx.rotate(rot);
      ctx.translate(-f.w / 2, -f.h / 2);

      if (f.type === 'text') {
        drawText(ctx, f, val);
      } else if (f.type === 'barcode') {
        const png = await barcodeToPng(val, Math.round(f.w * 4), Math.round(f.h * 4), f.bcText);
        const img = new Image(); img.src = png;
        await new Promise(r => img.onload = r);
        ctx.drawImage(img, 0, 0, f.w, f.h);
      } else if (f.type === 'qr') {
        const png = await qrToPng(val, Math.round(f.w * 2), Math.round(f.h * 2));
        const img = new Image(); img.src = png;
        await new Promise(r => img.onload = r);
        ctx.drawImage(img, 0, 0, f.w, f.h);
      }
      ctx.restore();
    }
    
    // Add the flawlessly rendered page to the PDF
    const imgData = c.toDataURL('image/jpeg', 0.95);
    doc.addImage(imgData, 'JPEG', 0, 0, pgW, pgH);
    resolve();
  });
}

async function generatePDF() {
  if (!parsedData.length || !fields.length) { alert('Please add data and fields.'); return; }
  const btn = document.getElementById('genBtn');
  btn.disabled = true; btn.textContent = 'Generating...';
  document.getElementById('progressBar').style.display = 'block';
  document.getElementById('progressFill').style.width = '0%';

  const pgW  = parseFloat(document.getElementById('lwmm').value) || 80;
  const pgH  = parseFloat(document.getElementById('lhmm').value) || 40;
  const { jsPDF } = window.jspdf;

  const doc = new jsPDF({ unit: 'mm', format: [pgW, pgH], orientation: pgW > pgH ? 'landscape' : 'portrait' });

  // Calculate how many records each physical sheet consumes
  let maxIndex = 0;
  fields.forEach(f => { if (f.index > maxIndex) maxIndex = f.index; });
  const recordsPerPage = maxIndex + 1;

  let pageNum = 0;
  for (let i = 0; i < parsedData.length; i += recordsPerPage) {
    if (pageNum > 0) doc.addPage([pgW, pgH]);
    await drawPageToPdf(doc, i, pgW, pgH);
    pageNum++;
    document.getElementById('progressFill').style.width =
      Math.round((i / parsedData.length) * 100) + '%';
    await tick();
  }

  pdfDoc = doc;
  btn.disabled = false; btn.textContent = '⚡ Generate PDF';
  document.getElementById('progressBar').style.display = 'none';
  document.getElementById('dlWrap').style.display = 'flex';
  document.getElementById('previewBtn').style.display = 'inline-flex';
}

/* ─── PDF Preview ─────────────────────────────── */
function previewPDF() {
  if (!pdfDoc) { alert('Generate PDF first.'); return; }
  const blob = pdfDoc.output('blob');
  const url = URL.createObjectURL(blob);
  document.getElementById('pdfPreviewFrame').src = url;
  document.getElementById('pdfPreviewModal').style.display = 'flex';
}

function closePdfPreview() {
  const frame = document.getElementById('pdfPreviewFrame');
  document.getElementById('pdfPreviewModal').style.display = 'none';
  if (frame.src.startsWith('blob:')) URL.revokeObjectURL(frame.src);
  frame.src = '';
}

function barcodeToPng(val, w, h, showText) {
  return new Promise((res, rej) => {
    const svgEl = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    try {
      JsBarcode(svgEl, String(val), {
        format: 'CODE128',
        width: 3,
        height: Math.max(40, h - 40),
        displayValue: showText,
        fontSize: 24,
        textMargin: 6,
        margin: 6,
        background: '#ffffff',
        lineColor: '#000000',
        valid: () => true
      });
      const svgStr = new XMLSerializer().serializeToString(svgEl);
      const img = new Image();
      img.onload = () => {
        const c = document.createElement('canvas');
        c.width = w; c.height = h;
        const ctx = c.getContext('2d');
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, w, h);
        // Fill the field at exact w×h so the PDF placement doesn't stretch it further.
        // This keeps bar positions consistent with the canvas designer preview.
        ctx.drawImage(img, 0, 0, w, h);
        res(c.toDataURL('image/png'));
      };
      img.onerror = rej;
      img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgStr)));
    } catch (e) { rej(e); }
  });
}

function qrToPng(val, w, h) {
  return new Promise((res, rej) => {
    const div = document.createElement('div');
    div.style.cssText = 'position:fixed;left:-9999px;top:-9999px';
    document.body.appendChild(div);
    try {
      new QRCode(div, { text: val, width: w, height: h, correctLevel: QRCode.CorrectLevel.M });
      setTimeout(() => {
        const el = div.querySelector('canvas') || div.querySelector('img');
        if (!el) { document.body.removeChild(div); rej(); return; }
        const src = el.tagName === 'CANVAS' ? el.toDataURL() : el.src;
        document.body.removeChild(div);
        res(src);
      }, 130);
    } catch (e) { document.body.removeChild(div); rej(e); }
  });
}

function downloadPDF() {
  if (!pdfDoc) return;
  const name = (document.getElementById('pdfName').value.trim() || 'seznik-labels').replace(/\.pdf$/i, '');
  pdfDoc.save(name + '.pdf');
}

function tick() { return new Promise(r => setTimeout(r, 0)); }
function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
