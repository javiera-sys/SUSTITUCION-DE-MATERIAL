/**
 * Guardado de archivos (Excel / PDF / ZIP / TXT) con confirmación de reemplazo.
 *
 * Estrategia por niveles (se elige según lo que soporte el navegador):
 *   1) CARPETA  — showDirectoryPicker (Chrome/Edge de PC, y Chrome Android reciente).
 *                 Se comprueba qué archivos ya existen y se pregunta ANTES de escribir.
 *                 Nunca se sobrescribe nada sin autorización del usuario.
 *   2) ARCHIVO  — showSaveFilePicker. El propio sistema/navegador pregunta si se
 *                 reemplaza el archivo existente (no se puede simular desde la página).
 *   3) DESCARGA — <a download>. Es el único método universal (Firefox, Safari, iOS),
 *                 pero la página NO puede saber si el archivo existe ni reemplazarlo;
 *                 el navegador decide (normalmente renombra a "nombre (1).ext").
 *                 Se avisa de esto con claridad y no se simula ninguna confirmación.
 *
 * API pública:  window.Guardado.saveFiles(items, opts)
 *   items: [{ name: 'archivo.xlsx', blob: Blob }]
 *   opts : { prefer: 'single' (guardar archivo por archivo con showSaveFilePicker),
 *            onProgress: (hechos, total) => void }
 *   Devuelve: { mode, saved, replaced, renamed, skipped, errors[], cancelled }
 */
(function (global) {
  'use strict';

  const hasDirPicker = typeof window.showDirectoryPicker === 'function';
  const hasSavePicker = typeof window.showSaveFilePicker === 'function';
  const PICKER_ID_DIR = 'habilitados-salida';
  const PICKER_ID_FILE = 'habilitados-guardar';
  const CONCURRENCY = /Android/i.test(navigator.userAgent) ? 1 : 4;
  let downloadInfoShown = false;

  const MIME = {
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.xlsm': 'application/vnd.ms-excel.sheet.macroEnabled.12',
    '.xls': 'application/vnd.ms-excel',
    '.pdf': 'application/pdf',
    '.zip': 'application/zip',
    '.txt': 'text/plain'
  };

  // ------------------------------------------------------------
  // Utilidades de nombre
  // ------------------------------------------------------------
  function extOf(name) { const i = name.lastIndexOf('.'); return i > 0 ? name.slice(i) : ''; }
  function stem(name) { const i = name.lastIndexOf('.'); return i > 0 ? name.slice(0, i) : name; }
  function cleanName(n) { return String(n == null ? '' : n).replace(/[\\/:"*?<>|]/g, '').trim(); }

  function normalizeName(value, ext) {
    let c = cleanName(value);
    if (!c) return '';
    if (ext && !c.toLowerCase().endsWith(ext.toLowerCase())) c += ext;
    return c;
  }

  // Evita choques de nombre dentro del mismo lote (sin distinguir mayúsculas, como Windows)
  function dedupe(name, used) {
    let candidate = name;
    let n = 1;
    while (used.has(candidate.toLowerCase())) {
      candidate = stem(name) + ' (' + n + ')' + extOf(name);
      n++;
    }
    used.add(candidate.toLowerCase());
    return candidate;
  }

  // ------------------------------------------------------------
  // Diálogo modal (promesa)
  // ------------------------------------------------------------
  function el(tag, cls, text) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }

  function openDialog(cfg) {
    return new Promise((resolve) => {
      const prevFocus = document.activeElement;
      const overlay = el('div', 'gd-overlay');
      const box = el('div', 'gd-modal');
      box.setAttribute('role', 'dialog');
      box.setAttribute('aria-modal', 'true');
      box.appendChild(el('h3', 'gd-title', cfg.title));
      (cfg.paragraphs || []).forEach((t) => box.appendChild(el('p', 'gd-text', t)));

      let inp = null, errEl = null, cb = null;
      if (cfg.input !== undefined) {
        inp = el('input', 'gd-input');
        inp.type = 'text';
        inp.value = cfg.input;
        errEl = el('p', 'gd-error');
        box.append(inp, errEl);
      }
      if (cfg.checkbox) {
        const lab = el('label', 'checkbox-label gd-check');
        cb = document.createElement('input');
        cb.type = 'checkbox';
        lab.append(cb, document.createTextNode(' ' + cfg.checkbox));
        box.appendChild(lab);
      }

      const cancelBtn = cfg.buttons.filter((b) => b.isCancel)[0];
      const cancelValue = cancelBtn ? cancelBtn.value : undefined;

      function finish(value) {
        document.removeEventListener('keydown', onKey, true);
        overlay.remove();
        if (prevFocus && prevFocus.focus) { try { prevFocus.focus(); } catch (e) { /* ignorar */ } }
        resolve({ value: value, checked: !!(cb && cb.checked), input: inp ? inp.value : null });
      }
      function onKey(e) {
        if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); finish(cancelValue); }
      }

      const row = el('div', 'gd-actions');
      let busy = false;
      let primaryBtn = null;
      let focusBtn = null;
      cfg.buttons.forEach((b) => {
        const btn = el('button', 'btn ' + (b.cls || 'btn-secondary'), b.label);
        btn.type = 'button';
        btn.addEventListener('click', async () => {
          if (busy) return;
          if (b.validate && inp) {
            busy = true;
            const msg = await b.validate(inp.value);
            busy = false;
            if (msg) { errEl.textContent = msg; inp.focus(); return; }
          }
          finish(b.value);
        });
        if (b.primary) primaryBtn = btn;
        if (b.focus) focusBtn = btn;
        row.appendChild(btn);
      });
      box.appendChild(row);
      overlay.appendChild(box);
      document.body.appendChild(overlay);
      document.addEventListener('keydown', onKey, true);

      if (inp) {
        inp.addEventListener('keydown', (e) => {
          if (e.key === 'Enter' && primaryBtn) { e.preventDefault(); primaryBtn.click(); }
        });
        inp.addEventListener('input', () => { errEl.textContent = ''; });
        inp.focus();
        inp.select();
      } else {
        (focusBtn || row.firstChild).focus();
      }
    });
  }

  function askConflict(name, remaining) {
    return openDialog({
      title: '⚠️ El archivo ya existe',
      paragraphs: ['Ya existe «' + name + '» en la carpeta seleccionada.', '¿Deseas sustituirlo?'],
      checkbox: remaining > 1 ? 'Aplicar esta decisión a los demás archivos que ya existan (' + (remaining - 1) + ' más)' : null,
      buttons: [
        { label: 'Sí, sustituir archivo', value: 'replace', cls: 'btn-primary' },
        { label: 'No, conservar archivo existente', value: 'keep', cls: 'btn-secondary', focus: true },
        { label: 'Cancelar descarga', value: 'cancel', cls: 'btn-secondary', isCancel: true }
      ]
    });
  }

  function askRename(suggested, validate) {
    return openDialog({
      title: 'Nuevo nombre para el archivo',
      paragraphs: ['El archivo existente se conserva intacto. Escribe el nombre con el que se guardará el nuevo.'],
      input: suggested,
      buttons: [
        { label: 'Guardar con este nombre', value: 'ok', cls: 'btn-primary', primary: true, validate: validate },
        { label: 'Omitir este archivo', value: 'skip', cls: 'btn-secondary' },
        { label: 'Cancelar descarga', value: 'cancel', cls: 'btn-secondary', isCancel: true }
      ]
    });
  }

  function askBlockedFolder(err, canSingle) {
    const buttons = [{ label: 'Elegir otra carpeta', value: 'retry', cls: 'btn-primary', focus: true }];
    if (canSingle) buttons.push({ label: 'Guardar archivo por archivo', value: 'single', cls: 'btn-secondary' });
    buttons.push({ label: 'Cancelar', value: 'cancel', cls: 'btn-secondary', isCancel: true });
    return openDialog({
      title: 'No se pudo usar esa carpeta',
      paragraphs: [
        'El navegador no permitió usar la carpeta elegida (' + ((err && (err.message || err.name)) || 'error desconocido') + ').',
        'Algunos navegadores bloquean carpetas del sistema como «Descargas» completa; prueba con una subcarpeta.'
      ],
      buttons: buttons
    });
  }

  function askNextSingle(name, i, total) {
    return openDialog({
      title: 'Archivo ' + i + ' de ' + total,
      paragraphs: [
        'Siguiente: «' + name + '»',
        'El navegador abrirá su ventana de guardado. Si ya existe un archivo con ese nombre, te preguntará si deseas reemplazarlo.'
      ],
      buttons: [
        { label: 'Guardar…', value: 'go', cls: 'btn-primary', focus: true },
        { label: 'Omitir este archivo', value: 'skip', cls: 'btn-secondary' },
        { label: 'Cancelar el resto', value: 'cancel', cls: 'btn-secondary', isCancel: true }
      ]
    });
  }

  function askDownloadInfo(count) {
    return openDialog({
      title: 'Descarga estándar del navegador',
      paragraphs: [
        'Este navegador no permite elegir carpeta ni confirmar el reemplazo desde la página.',
        'Se descargará(n) ' + count + ' archivo(s) en tu carpeta de descargas. Si ya existe uno con el mismo nombre, el navegador decidirá qué hacer (normalmente lo renombra como «nombre (1)»). Para poder sustituir archivos usa Chrome o Edge.'
      ],
      buttons: [
        { label: 'Descargar', value: 'go', cls: 'btn-primary', focus: true },
        { label: 'Cancelar', value: 'cancel', cls: 'btn-secondary', isCancel: true }
      ]
    });
  }

  // ------------------------------------------------------------
  // Operaciones sobre el sistema de archivos
  // ------------------------------------------------------------
  async function exists(dir, name) {
    try {
      await dir.getFileHandle(name);
      return true;
    } catch (e) {
      if (e && e.name === 'NotFoundError') return false;
      if (e && e.name === 'TypeMismatchError') return true; // hay una carpeta con ese nombre
      throw e;
    }
  }

  // createWritable escribe en un archivo temporal y reemplaza al cerrar:
  // si algo falla a medias, el archivo existente queda intacto.
  async function writeBlob(fileHandle, blob) {
    const w = await fileHandle.createWritable();
    try {
      await w.write(blob);
      await w.close();
    } catch (e) {
      try { await w.abort(); } catch (_) { /* ignorar */ }
      throw e;
    }
  }

  async function suggestFree(dir, name, taken) {
    let n = 1;
    let candidate;
    do {
      candidate = stem(name) + ' (' + n + ')' + extOf(name);
      n++;
    } while (taken.has(candidate.toLowerCase()) || await exists(dir, candidate));
    return candidate;
  }

  function newResult(mode) {
    return { mode: mode, saved: 0, replaced: 0, renamed: 0, skipped: 0, errors: [], cancelled: false };
  }

  // ------------------------------------------------------------
  // Nivel 1: carpeta
  // ------------------------------------------------------------
  async function saveToDirectory(dir, items, opts) {
    const res = newResult('folder');
    const used = new Set();
    const entries = items.map((it) => ({
      name: dedupe(cleanName(it.name) || 'archivo', used),
      blob: it.blob
    }));

    // Comprobación de existencia en paralelo (rápida) ANTES de escribir nada
    const flags = await Promise.all(entries.map((e) => exists(dir, e.name)));
    const taken = new Set(entries.map((e) => e.name.toLowerCase()));
    let pending = flags.filter(Boolean).length;
    let auto = null; // decisión aplicada a todos: 'replace' | 'keep'
    const plan = [];

    for (let i = 0; i < entries.length; i++) {
      const e = entries[i];
      if (!flags[i]) { plan.push(e); continue; }

      let decision = auto;
      if (!decision) {
        const r = await askConflict(e.name, pending);
        if (r.value === 'cancel') { res.cancelled = true; return res; } // no se escribió nada aún
        decision = r.value;
        if (r.checked) auto = decision;
      }
      pending--;

      if (decision === 'replace') {
        res.replaced++;
        plan.push(e);
      } else {
        const ext = extOf(e.name);
        const suggestion = await suggestFree(dir, e.name, taken);
        let finalName = suggestion;
        if (auto !== 'keep') {
          const r2 = await askRename(suggestion, async (val) => {
            const nn = normalizeName(val, ext);
            if (!nn) return 'Escribe un nombre.';
            if (taken.has(nn.toLowerCase()) || await exists(dir, nn)) return 'Ya existe un archivo con ese nombre. Elige otro.';
            return null;
          });
          if (r2.value === 'cancel') { res.cancelled = true; return res; }
          if (r2.value === 'skip') { res.skipped++; continue; }
          finalName = normalizeName(r2.input, ext);
        }
        taken.add(finalName.toLowerCase());
        res.renamed++;
        plan.push({ name: finalName, blob: e.blob });
      }
    }

    // Escritura (en paralelo limitado)
    let idx = 0;
    let done = 0;
    async function worker() {
      while (idx < plan.length) {
        const e = plan[idx++];
        try {
          const fh = await dir.getFileHandle(e.name, { create: true });
          await writeBlob(fh, e.blob);
          res.saved++;
        } catch (err) {
          res.errors.push(e.name + ': ' + ((err && (err.message || err.name)) || 'error'));
        }
        done++;
        if (opts.onProgress) opts.onProgress(done, plan.length);
      }
    }
    await Promise.all(Array.from({ length: Math.min(CONCURRENCY, plan.length) }, worker));
    return res;
  }

  // ------------------------------------------------------------
  // Nivel 2: showSaveFilePicker (confirmación de reemplazo nativa)
  // ------------------------------------------------------------
  async function saveWithPickers(items, opts) {
    const res = newResult('single');
    for (let i = 0; i < items.length; i++) {
      const name = cleanName(items[i].name) || 'archivo';
      // Cada ventana nativa exige un clic nuevo del usuario: de ahí este paso intermedio
      if (i > 0) {
        const r = await askNextSingle(name, i + 1, items.length);
        if (r.value === 'cancel') { res.cancelled = true; break; }
        if (r.value === 'skip') { res.skipped++; continue; }
      }
      try {
        const ext = extOf(name).toLowerCase();
        const options = { suggestedName: name, id: PICKER_ID_FILE };
        if (MIME[ext]) {
          const accept = {};
          accept[MIME[ext]] = [ext];
          options.types = [{ description: 'Archivo ' + ext.slice(1).toUpperCase(), accept: accept }];
        }
        const fh = await window.showSaveFilePicker(options);
        await writeBlob(fh, items[i].blob);
        res.saved++;
      } catch (e) {
        if (e && e.name === 'AbortError') {
          res.skipped++;
          if (items.length === 1) res.cancelled = true;
          continue;
        }
        res.errors.push(name + ': ' + ((e && (e.message || e.name)) || 'error'));
      }
      if (opts.onProgress) opts.onProgress(i + 1, items.length);
    }
    return res;
  }

  // ------------------------------------------------------------
  // Nivel 3: descarga estándar
  // ------------------------------------------------------------
  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  async function saveWithDownloads(items, opts) {
    const res = newResult('download');
    if (!downloadInfoShown) {
      const r = await askDownloadInfo(items.length);
      if (r.value !== 'go') { res.cancelled = true; return res; }
      downloadInfoShown = true;
    }
    const used = new Set();
    for (let i = 0; i < items.length; i++) {
      const name = dedupe(cleanName(items[i].name) || 'archivo', used);
      downloadBlob(items[i].blob, name);
      res.saved++;
      if (opts.onProgress) opts.onProgress(i + 1, items.length);
      if (i < items.length - 1) await new Promise((r) => setTimeout(r, 300));
    }
    return res;
  }

  // ------------------------------------------------------------
  // Punto de entrada
  // ------------------------------------------------------------
  async function saveFiles(items, opts) {
    opts = opts || {};
    items = (items || []).filter((i) => i && i.blob);
    if (!items.length) return newResult('none');

    const preferSingle = opts.prefer === 'single';

    if (!preferSingle && hasDirPicker) {
      let dir = null;
      for (;;) {
        try {
          dir = await window.showDirectoryPicker({ id: PICKER_ID_DIR, mode: 'readwrite', startIn: 'downloads' });
          break;
        } catch (e) {
          if (e && e.name === 'AbortError') { const r = newResult('folder'); r.cancelled = true; return r; }
          const choice = await askBlockedFolder(e, hasSavePicker);
          if (choice.value === 'retry') continue;
          if (choice.value === 'single') return saveWithPickers(items, opts);
          const r = newResult('folder'); r.cancelled = true; return r;
        }
      }
      try {
        return await saveToDirectory(dir, items, opts);
      } catch (e) {
        const r = newResult('folder');
        r.errors.push((e && (e.message || e.name)) || 'error');
        return r;
      }
    }

    if (hasSavePicker) return saveWithPickers(items, opts);
    return saveWithDownloads(items, opts);
  }

  global.Guardado = {
    saveFiles: saveFiles,
    cleanName: cleanName,
    supportsFolder: hasDirPicker,
    supportsSavePicker: hasSavePicker
  };
})(window);
