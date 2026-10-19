/**
 * Herramientas Excel — navegación + módulo Multi-Copias
 * Se ejecuta aislado (IIFE) para no chocar con las variables de script.js
 * (Sustituidor de Materiales).
 */
(function () {
  'use strict';

  const $ = (s) => document.querySelector(s);
  const $$ = (s) => document.querySelectorAll(s);

  // ============================================================
  // NAVEGACIÓN ENTRE HERRAMIENTAS
  // ============================================================
  const VIEWS = ['sustituidor', 'copias'];

  function showView(name) {
    if (!VIEWS.includes(name)) name = 'sustituidor';
    VIEWS.forEach((v) => $('#view-' + v).classList.toggle('hidden', v !== name));
    $$('.tab-nav-btn').forEach((b) => {
      const active = b.dataset.view === name;
      b.classList.toggle('active', active);
      b.setAttribute('aria-selected', active);
    });
    try { history.replaceState(null, '', '#' + name); } catch (e) { /* ignorar */ }
  }

  $$('.tab-nav-btn').forEach((b) => b.addEventListener('click', () => showView(b.dataset.view)));
  showView(location.hash.replace('#', '') || 'sustituidor');

  // ============================================================
  // MULTI-COPIAS
  // ============================================================
  let baseFile = null;
  let names = [];
  const supportsDirPicker = 'showDirectoryPicker' in window;

  const dropBase = $('#mc-dropBase');
  const baseInput = $('#mc-baseInput');
  const baseDisplay = $('#mc-baseDisplay');
  const textarea = $('#mc-textarea');
  const addNamesBtn = $('#mc-addNamesBtn');
  const namesFileInput = $('#mc-namesFileInput');
  const listUI = $('#mc-list');
  const emptyUI = $('#mc-empty');
  const totalUI = $('#mc-total');
  const clearBtn = $('#mc-clearBtn');
  const dupWarning = $('#mc-dupWarning');
  const removeDupBtn = $('#mc-removeDupBtn');
  const generateBtn = $('#mc-generateBtn');
  const generateHint = $('#mc-generateHint');
  const toasts = $('#mc-toasts');
  const apiStatus = $('#mc-api-status');

  const BASE_PLACEHOLDER = baseDisplay.innerHTML;
  const GENERATE_LABEL = generateBtn.innerHTML;

  // --- Compatibilidad ---
  if (supportsDirPicker) {
    apiStatus.textContent = '📁 Guardado directo en carpeta activo';
  } else {
    apiStatus.textContent = '🗜️ Modo ZIP activo (respaldo)';
    apiStatus.classList.add('warn');
    generateBtn.innerHTML = '🗜️ GENERAR ARCHIVO ZIP';
    generateHint.textContent = 'Tu navegador no soporta el guardado directo en carpetas. Se descargará un archivo .zip';
  }

  // --- Toasts ---
  function toast(message, type = 'success', duration = 3000) {
    const el = document.createElement('div');
    el.className = 'mc-toast ' + type;
    el.textContent = message;
    toasts.appendChild(el);
    requestAnimationFrame(() => el.classList.add('show'));
    setTimeout(() => {
      el.classList.remove('show');
      setTimeout(() => el.remove(), 300);
    }, duration);
  }

  // --- Pestañas de ingreso de nombres ---
  $$('.mc-tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      $$('.mc-tab').forEach((b) => b.classList.remove('active'));
      $$('.mc-tab-content').forEach((c) => c.classList.add('hidden'));
      btn.classList.add('active');
      $('#' + btn.dataset.mcTarget).classList.remove('hidden');
    });
  });

  // --- Archivo base ---
  function resetBaseDisplay() {
    baseFile = null;
    baseInput.value = '';
    baseDisplay.innerHTML = BASE_PLACEHOLDER;
  }

  function handleBaseFile(file) {
    if (!file) return;
    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (!['.xlsx', '.xls', '.xlsm', '.pdf'].includes(ext)) {
      toast('Formato no válido. Usa .xlsx, .xls, .xlsm o .pdf', 'error');
      return;
    }
    baseFile = file;

    baseDisplay.innerHTML = '';
    const icon = document.createElement('p');
    icon.className = 'mc-icon';
    icon.textContent = ext === '.pdf' ? '📕' : '📗';
    const nameEl = document.createElement('p');
    nameEl.className = 'file-name';
    nameEl.textContent = file.name;
    const meta = document.createElement('p');
    meta.className = 'file-meta';
    meta.textContent = (file.size / 1024).toFixed(1) + ' KB';
    const rm = document.createElement('button');
    rm.className = 'mc-link-danger';
    rm.textContent = 'Quitar archivo';
    rm.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); resetBaseDisplay(); });
    baseDisplay.append(icon, nameEl, meta, rm);

    toast('Archivo base cargado correctamente');
  }

  baseInput.addEventListener('change', (e) => { handleBaseFile(e.target.files[0]); });
  dropBase.addEventListener('dragover', (e) => { e.preventDefault(); dropBase.classList.add('dragover'); });
  dropBase.addEventListener('dragleave', (e) => { e.preventDefault(); dropBase.classList.remove('dragover'); });
  dropBase.addEventListener('drop', (e) => {
    e.preventDefault();
    dropBase.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleBaseFile(e.dataTransfer.files[0]);
  });

  // --- Lista de nombres ---
  function sanitizeFilename(name) {
    return String(name).replace(/[\\/:"*?<>|]/g, '').trim();
  }

  function renderNames() {
    listUI.innerHTML = '';
    totalUI.textContent = names.length;

    if (names.length === 0) {
      emptyUI.classList.remove('hidden');
      listUI.classList.add('hidden');
      dupWarning.classList.add('hidden');
      return;
    }
    emptyUI.classList.add('hidden');
    listUI.classList.remove('hidden');
    dupWarning.classList.toggle('hidden', new Set(names).size === names.length);

    const frag = document.createDocumentFragment();
    names.forEach((name, i) => {
      const li = document.createElement('li');
      const left = document.createElement('div');
      const idx = document.createElement('span');
      idx.className = 'mc-idx';
      idx.textContent = (i + 1) + '.';
      const label = document.createElement('span');
      label.textContent = name;
      label.title = name;
      if (names.indexOf(name) !== names.lastIndexOf(name)) label.className = 'mc-dupe';
      left.append(idx, label);

      const del = document.createElement('button');
      del.className = 'mc-del';
      del.title = 'Eliminar';
      del.textContent = '✕';
      del.addEventListener('click', () => { names.splice(i, 1); renderNames(); });

      li.append(left, del);
      frag.appendChild(li);
    });
    listUI.appendChild(frag);
  }

  function addNames(rawArray) {
    let added = 0;
    rawArray.forEach((line) => {
      const clean = sanitizeFilename(line);
      if (clean.length > 0) { names.push(clean); added++; }
    });
    if (added > 0) { renderNames(); toast('Se agregaron ' + added + ' nombres'); }
    else toast('No se encontraron nombres válidos', 'warning');
  }

  addNamesBtn.addEventListener('click', () => {
    addNames(textarea.value.split(/\r?\n/));
    textarea.value = '';
  });

  namesFileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = ev.target.result;
        let raw = [];
        if (ext === '.txt' || ext === '.csv') {
          raw = new TextDecoder('utf-8').decode(data).split(/\r?\n/);
        } else {
          const wb = XLSX.read(data, { type: 'array' });
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
          raw = rows.map((r) => r[0]).filter((v) => v !== undefined && v !== null).map(String);
        }
        addNames(raw);
      } catch (err) {
        console.error(err);
        toast('Error al leer el archivo.', 'error');
      }
      namesFileInput.value = '';
    };
    reader.readAsArrayBuffer(file);
  });

  clearBtn.addEventListener('click', () => {
    if (!names.length) return;
    names = [];
    renderNames();
    toast('Lista limpiada');
  });

  removeDupBtn.addEventListener('click', () => {
    const unique = [...new Set(names)];
    const removed = names.length - unique.length;
    names = unique;
    renderNames();
    toast('Se eliminaron ' + removed + ' duplicados');
  });

  // --- Generación ---
  function uniqueName(name, ext, used) {
    const full = name.toLowerCase().endsWith(ext.toLowerCase()) ? name : name + ext;
    let candidate = full;
    let n = 1;
    while (used.has(candidate)) {
      candidate = full.substring(0, full.length - ext.length) + '_(' + n + ')' + ext;
      n++;
    }
    used.add(candidate);
    return candidate;
  }

  generateBtn.addEventListener('click', async () => {
    if (!baseFile) {
      toast('Primero debes importar un archivo base', 'error');
      dropBase.style.borderColor = 'var(--danger)';
      setTimeout(() => { dropBase.style.borderColor = ''; }, 1000);
      return;
    }
    if (names.length === 0) {
      toast('La lista de nombres está vacía', 'error');
      return;
    }

    const originalHtml = generateBtn.innerHTML;
    generateBtn.disabled = true;
    generateBtn.textContent = '⏳ PROCESANDO...';

    try {
      const baseExt = baseFile.name.substring(baseFile.name.lastIndexOf('.'));
      const used = new Set();

      if (supportsDirPicker) {
        let dirHandle;
        try {
          dirHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
        } catch (err) {
          if (err.name === 'AbortError') { toast('Selección de carpeta cancelada', 'warning'); return; }
          throw err;
        }
        let ok = 0;
        const failed = [];
        for (const name of names) {
          const finalName = uniqueName(name, baseExt, used);
          try {
            const fh = await dirHandle.getFileHandle(finalName, { create: true });
            const w = await fh.createWritable();
            await w.write(baseFile);
            await w.close();
            ok++;
          } catch (fileErr) {
            console.error('Error escribiendo ' + finalName, fileErr);
            failed.push(finalName);
          }
        }
        if (failed.length) toast('Se guardaron ' + ok + ' archivos; fallaron ' + failed.length, 'warning', 5000);
        else toast('¡Éxito! Se guardaron ' + ok + ' archivos en la carpeta.');
      } else {
        generateBtn.textContent = '🗜️ EMPAQUETANDO ZIP...';
        const zip = new JSZip();
        names.forEach((name) => zip.file(uniqueName(name, baseExt, used), baseFile));
        const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
        const baseName = baseFile.name.substring(0, baseFile.name.lastIndexOf('.'));
        const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'Copias_' + baseName + '_' + dateStr + '.zip';
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(a.href); }, 100);
        toast('¡Éxito! Se generó el ZIP con ' + names.length + ' archivos.');
      }
    } catch (err) {
      console.error(err);
      toast('Ocurrió un error inesperado al procesar.', 'error');
    } finally {
      generateBtn.disabled = false;
      generateBtn.innerHTML = originalHtml;
    }
  });

  renderNames();
})();
