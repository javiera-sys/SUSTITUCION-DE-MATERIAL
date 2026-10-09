/**
 * Sustituidor de Materiales Excel
 * Procesamiento 100% local con ExcelJS + JSZip
 * 
 * Muestra N° y DESCRIPTION de la fila donde se encuentra cada coincidencia.
 * Resalta en amarillo los cambios para replicar la vista de Excel.
 */

// ============================================================
// ESTADO GLOBAL
// ============================================================
const state = {
  files: [],
  rules: [],
  previewResults: null,
  processedResults: null,
  fileIdCounter: 0,
  ruleIdCounter: 0
};

// ============================================================
// REFERENCIAS DOM
// ============================================================
const $ = (sel) => document.querySelector(sel);
const fileInput = $('#fileInput');
const dropZone = $('#dropZone');
const fileList = $('#fileList');
const searchInput = $('#searchInput');
const replaceInput = $('#replaceInput');
const addRuleBtn = $('#addRuleBtn');
const caseSensitiveCheck = $('#caseSensitive');
const exactMatchCheck = $('#exactMatch');
const rulesBody = $('#rulesBody');
const analyzeBtn = $('#analyzeBtn');
const applyBtn = $('#applyBtn');
const progressBar = $('#progressBar');
const progressFill = $('#progressFill');
const progressText = $('#progressText');
const sectionPreview = $('#section-preview');
const previewContent = $('#previewContent');
const sectionResults = $('#section-results');
const resultsSummary = $('#resultsSummary');
const resultsDetails = $('#resultsDetails');
const downloadZipBtn = $('#downloadZipBtn');
const downloadReportBtn = $('#downloadReportBtn');

// ============================================================
// GESTIÓN DE ARCHIVOS
// ============================================================
function addFiles(files) {
  const validFiles = Array.from(files).filter(f =>
    f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
  );

  validFiles.forEach(file => {
    if (file.name.endsWith('.xls')) {
      alert(`"${file.name}" es un archivo .xls.\n\nExcelJS no puede procesar .xls sin pérdida de formato. Por favor:\n1. Abre el archivo en Excel\n2. Guárdalo como .xlsx\n3. Vuelve a cargarlo.`);
      return;
    }
    state.files.push({ file, id: `file-${++state.fileIdCounter}` });
  });

  renderFileList();
  updateButtonStates();
}

function removeFile(id) {
  state.files = state.files.filter(f => f.id !== id);
  renderFileList();
  updateButtonStates();
}

function renderFileList() {
  if (state.files.length === 0) { fileList.innerHTML = ''; return; }
  fileList.innerHTML = state.files.map(f => `
    <li>
      <span>${escapeHtml(f.file.name)} <small style="color:#64748b">(${formatSize(f.file.size)})</small></span>
      <button class="remove-file" onclick="removeFile('${f.id}')" title="Eliminar">×</button>
    </li>
  `).join('');
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// ============================================================
// GESTIÓN DE REGLAS
// ============================================================
function addRule() {
  const search = searchInput.value.trim();
  const replace = replaceInput.value.trim();

  if (!search) { alert('El campo "Buscar" no puede estar vacío.'); return; }

  const exists = state.rules.find(r =>
    r.search.toLowerCase() === search.toLowerCase()
  );
  if (exists) { alert(`Ya existe una regla para "${search}".`); return; }

  state.rules.push({ id: `rule-${++state.ruleIdCounter}`, search, replace });
  searchInput.value = '';
  replaceInput.value = '';
  searchInput.focus();
  renderRules();
  updateButtonStates();
}

function removeRule(id) {
  state.rules = state.rules.filter(r => r.id !== id);
  renderRules();
  updateButtonStates();
}

function renderRules() {
  if (state.rules.length === 0) {
    rulesBody.innerHTML = `<tr><td colspan="3" class="empty-rules">No hay reglas definidas. Agrega al menos una.</td></tr>`;
    return;
  }
  rulesBody.innerHTML = state.rules.map(r => `
    <tr>
      <td>${escapeHtml(r.search)}</td>
      <td>${escapeHtml(r.replace)}</td>
      <td>
        <div class="rule-actions">
          <button onclick="removeRule('${r.id}')" title="Eliminar">🗑️</button>
        </div>
      </td>
    </tr>
  `).join('');
}

function updateButtonStates() {
  const hasFiles = state.files.length > 0;
  const hasRules = state.rules.length > 0;
  analyzeBtn.disabled = !(hasFiles && hasRules);
  applyBtn.disabled = !state.previewResults;
}

// ============================================================
// LECTURA DE ARCHIVOS
// ============================================================
async function readWorkbook(file) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  return workbook;
}

// ============================================================
// MATCHING
// ============================================================
function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function findMatchWithReplacement(text, search, replacement, options) {
  const { caseSensitive, exactMatch } = options;
  let haystack = text;
  let needle = search;

  if (!caseSensitive) {
    haystack = haystack.toLowerCase();
    needle = needle.toLowerCase();
  }

  if (exactMatch) {
    if (haystack === needle) return { matched: true, result: replacement, count: 1 };
    return { matched: false, result: text, count: 0 };
  }

  if (haystack.includes(needle)) {
    const regex = new RegExp(escapeRegex(search), caseSensitive ? 'g' : 'gi');
    const count = (text.match(regex) || []).length;
    return { matched: true, result: text.replace(regex, replacement), count };
  }

  return { matched: false, result: text, count: 0 };
}

// ============================================================
// RESALTAR PARTES DEL TEXTO
// ============================================================
/**
 * Genera HTML con las coincidencias resaltadas.
 * @param {string} text - Texto completo
 * @param {string} search - Texto a resaltar
 * @param {boolean} caseSensitive - Si distingue mayúsculas
 * @param {string} cssClass - Clase CSS a aplicar al match
 */
function highlightText(text, search, caseSensitive, cssClass) {
  if (!search) return escapeHtml(text);
  const regex = new RegExp(escapeRegex(search), caseSensitive ? 'g' : 'gi');
  // Dividir el texto en partes: antes, match, después
  const parts = [];
  let lastIndex = 0;
  let match;
  const regexGlobal = new RegExp(escapeRegex(search), caseSensitive ? 'g' : 'gi');
  while ((match = regexGlobal.exec(text)) !== null) {
    if (match.index > lastIndex) {
      parts.push(escapeHtml(text.substring(lastIndex, match.index)));
    }
    parts.push(`<span class="${cssClass}">${escapeHtml(match[0])}</span>`);
    lastIndex = match.index + match[0].length;
    if (match[0].length === 0) regexGlobal.lastIndex++; // evitar loop infinito
  }
  if (lastIndex < text.length) {
    parts.push(escapeHtml(text.substring(lastIndex)));
  }
  return parts.join('');
}

// ============================================================
// ANÁLISIS
// ============================================================
async function analyzeFiles() {
  if (state.files.length === 0 || state.rules.length === 0) return;

  showProgress('Analizando archivos...', 0);
  const results = [];
  const total = state.files.length;

  for (let i = 0; i < total; i++) {
    const fileEntry = state.files[i];
    updateProgress(`Analizando: ${fileEntry.file.name}`, (i / total) * 100);

    try {
      const workbook = await readWorkbook(fileEntry.file);
      const changes = analyzeWorkbook(workbook, state.rules, {
        caseSensitive: caseSensitiveCheck.checked,
        exactMatch: exactMatchCheck.checked
      });

      results.push({
        fileName: fileEntry.file.name,
        fileId: fileEntry.id,
        changes,
        totalChanges: changes.reduce((sum, c) => sum + c.count, 0),
        error: null
      });
    } catch (err) {
      results.push({
        fileName: fileEntry.file.name,
        fileId: fileEntry.id,
        changes: [],
        totalChanges: 0,
        error: err.message || 'Error al leer el archivo'
      });
    }
  }

  updateProgress('Análisis completado', 100);
  hideProgress();
  state.previewResults = results;
  renderPreview(results);
  updateButtonStates();
}

function analyzeWorkbook(workbook, rules, options) {
  const changes = [];

  workbook.eachSheet((worksheet) => {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const rowNum = row.number;
      const cellA = worksheet.getCell(`A${rowNum}`);
      const cellB = worksheet.getCell(`B${rowNum}`);
      const numValue = cellA.value != null ? String(cellA.value).trim() : '';
      const descValue = cellB.value != null ? String(cellB.value).trim() : '';

      row.eachCell({ includeEmpty: false }, (cell) => {
        const cellValue = cell.value;
        if (typeof cellValue !== 'string') return;
        if (cell.col === 1 || cell.col === 2) return;

        for (const rule of rules) {
          const m = findMatchWithReplacement(
            cellValue, rule.search, rule.replace, options
          );

          if (m.matched && m.result !== cellValue) {
            changes.push({
              sheet: worksheet.name,
              cell: cell.address,
              rowNum: rowNum,
              itemNum: numValue,
              itemDesc: descValue,
              original: cellValue,
              replaced: m.result,
              ruleSearch: rule.search,
              ruleReplace: rule.replace,
              count: m.count
            });
            break;
          }
        }
      });
    });
  });

  return changes;
}

// ============================================================
// VISTA PREVIA
// ============================================================
function renderPreview(results) {
  sectionPreview.classList.remove('hidden');

  const totalChanges = results.reduce((sum, r) => sum + r.totalChanges, 0);
  const filesWithChanges = results.filter(r => r.totalChanges > 0);
  const caseSensitive = caseSensitiveCheck.checked;

  let html = `<p style="margin-bottom:1rem;font-size:0.85rem;color:#64748b">
    Se encontraron <strong>${totalChanges}</strong> coincidencia(s) en
    <strong>${filesWithChanges.length}</strong> de <strong>${results.length}</strong> archivo(s).
  </p>`;

  if (totalChanges === 0) {
    html += `<div class="result-card warning">No se encontraron coincidencias con las reglas definidas.</div>`;
  }

  results.forEach(r => {
    if (r.error) {
      html += `<div class="result-card error">
        <strong>${escapeHtml(r.fileName)}</strong> — Error: ${escapeHtml(r.error)}
      </div>`;
      return;
    }

    if (r.totalChanges === 0) {
      html += `<div class="result-card warning">
        <strong>${escapeHtml(r.fileName)}</strong> — Sin coincidencias
      </div>`;
      return;
    }

    html += `<div class="preview-file">
      <div class="preview-file-header">
        ${escapeHtml(r.fileName)} — ${r.totalChanges} cambio(s)
      </div>`;

    r.changes.forEach(c => {
      // Texto original con la coincidencia resaltada en amarillo
      const originalHtml = highlightText(
        c.original, c.ruleSearch, caseSensitive, 'hl-yellow'
      );
      // Texto nuevo con la sustitución resaltada en amarillo
      const replacedHtml = highlightText(
        c.replaced, c.ruleReplace, caseSensitive, 'hl-yellow'
      );

      html += `<div class="preview-change-block preview-highlighted">
        <div class="preview-row-info">
          <span class="badge-item">N° ${escapeHtml(c.itemNum || '—')}</span>
          <span class="item-desc">${escapeHtml(c.itemDesc || '(sin descripción)')}</span>
        </div>
        <div class="preview-cell-ref">
          <span class="badge-sheet">${escapeHtml(c.sheet)}</span>
          <span class="badge-cell">${c.cell}</span>
          <span class="badge-rule">Regla: "${escapeHtml(c.ruleSearch)}" → "${escapeHtml(c.ruleReplace)}"</span>
        </div>
        <div class="preview-texts">
          <div class="preview-text-row">
            <span class="preview-label label-old">Contenido actual de la celda:</span>
            <div class="preview-value value-old">${originalHtml}</div>
          </div>
          <div class="preview-text-row">
            <span class="preview-label label-new">Contenido después de la sustitución:</span>
            <div class="preview-value value-new">${replacedHtml}</div>
          </div>
        </div>
      </div>`;
    });

    html += `</div>`;
  });

  html += `<div style="margin-top:1rem">
    <button id="cancelPreviewBtn" class="btn btn-secondary">Cancelar</button>
  </div>`;

  previewContent.innerHTML = html;

  const cancelBtn = document.getElementById('cancelPreviewBtn');
  if (cancelBtn) {
    cancelBtn.addEventListener('click', () => {
      sectionPreview.classList.add('hidden');
      state.previewResults = null;
      updateButtonStates();
    });
  }
}

// ============================================================
// APLICAR SUSTITUCIONES
// ============================================================
async function applyReplacements() {
  if (!state.previewResults) return;

  showProgress('Aplicando sustituciones...', 0);
  const results = [];
  const total = state.files.length;

  for (let i = 0; i < total; i++) {
    const fileEntry = state.files[i];
    updateProgress(`Procesando: ${fileEntry.file.name}`, (i / total) * 100);

    try {
      const workbook = await readWorkbook(fileEntry.file);
      const applyResult = applyToWorkbook(workbook, state.rules, {
        caseSensitive: caseSensitiveCheck.checked,
        exactMatch: exactMatchCheck.checked
      });

      if (applyResult.totalChanges === 0) {
        results.push({
          fileName: fileEntry.file.name,
          status: 'no_changes',
          changes: 0,
          outputName: null,
          blob: null,
          error: null,
          details: []
        });
        continue;
      }

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      const outputName = fileEntry.file.name.replace(/\.xlsx$/i, '_MODIFICADO.xlsx');

      results.push({
        fileName: fileEntry.file.name,
        status: 'modified',
        changes: applyResult.totalChanges,
        outputName,
        blob,
        error: null,
        details: applyResult.details
      });
    } catch (err) {
      results.push({
        fileName: fileEntry.file.name,
        status: 'error',
        changes: 0,
        outputName: null,
        blob: null,
        error: err.message || 'Error al procesar',
        details: []
      });
    }
  }

  updateProgress('Procesamiento completado', 100);
  hideProgress();
  state.processedResults = results;
  renderResults(results);
  sectionPreview.classList.add('hidden');
}

function applyToWorkbook(workbook, rules, options) {
  let totalChanges = 0;
  const details = [];

  workbook.eachSheet((worksheet) => {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const rowNum = row.number;
      const cellA = worksheet.getCell(`A${rowNum}`);
      const cellB = worksheet.getCell(`B${rowNum}`);
      const numValue = cellA.value != null ? String(cellA.value).trim() : '';
      const descValue = cellB.value != null ? String(cellB.value).trim() : '';

      row.eachCell({ includeEmpty: false }, (cell) => {
        const cellValue = cell.value;
        if (typeof cellValue !== 'string') return;
        if (cell.col === 1 || cell.col === 2) return;

        for (const rule of rules) {
          const m = findMatchWithReplacement(
            cellValue, rule.search, rule.replace, options
          );

          if (m.matched && m.result !== cellValue) {
            cell.value = m.result;
            totalChanges += m.count;
            details.push({
              sheet: worksheet.name,
              cell: cell.address,
              rowNum: rowNum,
              itemNum: numValue,
              itemDesc: descValue,
              original: cellValue,
              replaced: m.result,
              rule: rule.search
            });
            break;
          }
        }
      });
    });
  });

  return { totalChanges, details };
}

// ============================================================
// RESULTADOS
// ============================================================
function renderResults(results) {
  sectionResults.classList.remove('hidden');

  const modified = results.filter(r => r.status === 'modified');
  const noChanges = results.filter(r => r.status === 'no_changes');
  const errors = results.filter(r => r.status === 'error');
  const totalSubs = results.reduce((sum, r) => sum + (r.changes || 0), 0);

  resultsSummary.innerHTML = `
    <div class="result-stats">
      <div class="stat-box"><div class="stat-value">${results.length}</div><div class="stat-label">Procesados</div></div>
      <div class="stat-box"><div class="stat-value" style="color:#059669">${modified.length}</div><div class="stat-label">Modificados</div></div>
      <div class="stat-box"><div class="stat-value" style="color:#f59e0b">${noChanges.length}</div><div class="stat-label">Sin cambios</div></div>
      <div class="stat-box"><div class="stat-value" style="color:#dc2626">${errors.length}</div><div class="stat-label">Errores</div></div>
      <div class="stat-box"><div class="stat-value">${totalSubs}</div><div class="stat-label">Sustituciones</div></div>
    </div>
  `;

  let detailsHtml = '';
  results.forEach(r => {
    if (r.status === 'modified') {
      detailsHtml += `<div class="result-card success">
        <strong>${escapeHtml(r.fileName)}</strong> → <code>${escapeHtml(r.outputName)}</code>
        <br><small>${r.changes} sustitución(es) realizada(s)</small>
      </div>`;
    } else if (r.status === 'no_changes') {
      detailsHtml += `<div class="result-card warning">
        <strong>${escapeHtml(r.fileName)}</strong> — Sin coincidencias, no se generó archivo.
      </div>`;
    } else {
      detailsHtml += `<div class="result-card error">
        <strong>${escapeHtml(r.fileName)}</strong> — Error: ${escapeHtml(r.error)}
      </div>`;
    }
  });

  resultsDetails.innerHTML = detailsHtml;
  downloadZipBtn.disabled = modified.length === 0;
  downloadReportBtn.disabled = results.length === 0;
  applyBtn.disabled = true;
  state.previewResults = null;
}

// ============================================================
// DESCARGAS
// ============================================================
async function downloadZip() {
  if (!state.processedResults) return;
  const modified = state.processedResults.filter(r => r.status === 'modified');
  if (modified.length === 0) return;

  const zip = new JSZip();
  for (const r of modified) zip.file(r.outputName, r.blob);

  const zipBlob = await zip.generateAsync({ type: 'blob' });
  downloadBlob(zipBlob, 'materiales_modificados.zip');
}

function downloadReport() {
  if (!state.processedResults) return;

  const results = state.processedResults;
  const totalSubs = results.reduce((sum, r) => sum + (r.changes || 0), 0);

  let report = '========================================\n';
  report += 'REPORTE DE SUSTITUCIONES\n';
  report += `Fecha: ${new Date().toLocaleString('es-ES')}\n`;
  report += '========================================\n\n';
  report += `Archivos procesados: ${results.length}\n`;
  report += `Archivos modificados: ${results.filter(r => r.status === 'modified').length}\n`;
  report += `Archivos sin cambios: ${results.filter(r => r.status === 'no_changes').length}\n`;
  report += `Archivos con errores: ${results.filter(r => r.status === 'error').length}\n`;
  report += `Total sustituciones: ${totalSubs}\n\n`;
  report += '--- REGLAS APLICADAS ---\n';
  state.rules.forEach((r, i) => {
    report += `${i + 1}. "${r.search}" → "${r.replace}"\n`;
  });
  report += '\n--- DETALLE POR ARCHIVO ---\n\n';

  results.forEach(r => {
    report += `Archivo: ${r.fileName}\n`;
    report += `  Estado: ${r.status === 'modified' ? 'MODIFICADO' : r.status === 'no_changes' ? 'SIN CAMBIOS' : 'ERROR'}\n`;
    if (r.error) report += `  Error: ${r.error}\n`;
    if (r.status === 'modified') {
      report += `  Salida: ${r.outputName}\n  Cambios: ${r.changes}\n`;
      (r.details || []).forEach(d => {
        report += `    [${d.sheet}!${d.cell}] N° ${d.itemNum || '—'} - ${d.itemDesc || '(sin desc)'}\n`;
        report += `        "${d.original}" → "${d.replaced}"\n`;
      });
    }
    report += '\n';
  });

  const blob = new Blob([report], { type: 'text/plain;charset=utf-8' });
  downloadBlob(blob, 'reporte_sustituciones.txt');
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ============================================================
// PROGRESO
// ============================================================
function showProgress(text, percent) {
  progressBar.classList.remove('hidden');
  progressText.textContent = text;
  progressFill.style.width = `${Math.min(percent, 100)}%`;
}
function updateProgress(text, percent) {
  progressText.textContent = text;
  progressFill.style.width = `${Math.min(percent, 100)}%`;
}
function hideProgress() {
  progressBar.classList.add('hidden');
  progressFill.style.width = '0%';
}

// ============================================================
// UTILIDADES
// ============================================================
function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = String(str);
  return div.innerHTML;
}

// ============================================================
// EVENT LISTENERS
// ============================================================
fileInput.addEventListener('change', (e) => {
  addFiles(e.target.files);
  fileInput.value = '';
});

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('dragover');
});
dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  addFiles(e.dataTransfer.files);
});

addRuleBtn.addEventListener('click', addRule);
[searchInput, replaceInput].forEach(input => {
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') addRule();
  });
});

analyzeBtn.addEventListener('click', analyzeFiles);
applyBtn.addEventListener('click', applyReplacements);
downloadZipBtn.addEventListener('click', downloadZip);
downloadReportBtn.addEventListener('click', downloadReport);

window.removeFile = removeFile;
window.removeRule = removeRule;

renderRules();
updateButtonStates();
