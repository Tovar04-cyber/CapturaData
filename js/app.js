'use strict';

/* =====================================================================
   CapturaDatos – aplicación de captura de datos (PWA)
- Almacenamiento local: IndexedDB (persistente, por usuario).
- Campos base + campos dinámicos (texto, numérico, imagen).
- Exportación: Excel (.xlsx real), CSV y PDF (impresión).
- Respaldo: Google Drive (inicio de sesión con Google) y archivo .json.
- Instalable en celular (manifest.json + sw.js).
   ===================================================================== */

const GOOGLE_CLIENT_ID = window.CAPTURADATOS_GOOGLE_CLIENT_ID;
const GOOGLE_SCOPES = 'openid email profile https://www.googleapis.com/auth/drive.file';
const BACKUP_FILE_NAME = 'capturadatos-respaldo.json';

/* ---------- Conexión con el servidor ---------- */

const API_URL = window.CAPTURADATOS_API_URL || 'http://localhost:3000';
const TOKEN_KEY = 'capturadatos-token';
const USER_KEY = 'capturadatos-user';

let authToken = sessionStorage.getItem(TOKEN_KEY);
let authUser = null;

try {
  authUser = JSON.parse(sessionStorage.getItem(USER_KEY) || 'null');
} catch (_) {
  sessionStorage.removeItem(USER_KEY);
}

let isOnline = false;
let offlineMode = false;

/* ---------- Comunicación con la API ---------- */

async function apiRequest(path, options = {}) {
  const headers = {
    ...(options.headers || {})
  };

  if (authToken) {
    headers.Authorization = `Bearer ${authToken}`;
  }

  if (options.body && !headers['Content-Type']) {
    headers['Content-Type'] = 'application/json';
  }

  let response;

  try {
    response = await fetch(`${API_URL}${path}`, {
      ...options,
      headers
    });
  } catch (error) {
    isOnline = false;
    throw new Error('No se pudo conectar con el servidor.');
  }

  isOnline = true;

  if (response.status === 401 && authToken) {
    authToken = null;
    sessionStorage.removeItem(TOKEN_KEY);
    throw new Error('La sesión expiró o no está autorizada. Inicia sesión nuevamente.');
  }

  const data = await response.json().catch(() => ({}));

  if (!response.ok) {
    throw new Error(data.error || data.message || 'Error al comunicarse con el servidor.');
  }

  return data;
}

/* ---------- Autenticación ---------- */

async function loginWithGoogle(idToken) {
  logout(); // Limpiar estado previo antes de procesar la nueva cuenta

  const data = await apiRequest('/api/auth/google', {
    method: 'POST',
    body: JSON.stringify({ idToken })
  });

  if (!data.token || !data.user) {
    throw new Error('El servidor no devolvió una sesión válida.');
  }

  authToken = data.token;
  authUser = data.user;

  sessionStorage.setItem(TOKEN_KEY, authToken);
  sessionStorage.setItem(USER_KEY, JSON.stringify(authUser));
  localStorage.setItem('capturadatos-offline-user', JSON.stringify(authUser));

  return authUser;
}

function logout() {
  offlineMode = false;
  authToken = null;
  authUser = null;
  google.token = null;
  google.expires = 0;
  google.profile = null;

  sessionStorage.removeItem(TOKEN_KEY);
  sessionStorage.removeItem(USER_KEY);
  try { localStorage.removeItem('cd_google_profile'); } catch (_) {}
}

function showLogin() {
  $('#view-login').hidden = false;

  for (const v of ['form', 'list', 'fields', 'backup']) {
    $('#view-' + v).hidden = true;
  }

  $('#tabsNav').hidden = true;
  $('#apiLogoutBtn').hidden = true;
}

function showApp() {
  $('#view-login').hidden = true;
  $('#tabsNav').hidden = false;
  $('#apiLogoutBtn').hidden = false;

  showView('list');
}

async function startOfflineSession() {
  logout(); // Limpiar tokens previos

  try {
    db = await openDB('capturadatos-offline-db');

    let fields = await dbGetAll('fields');

    if (fields.length === 0) {
      await dbWrite('fields', (s) => {
        for (const field of BASE_FIELDS) {
          s.put(field);
        }
      });

      fields = await dbGetAll('fields');
    }

    const records = await dbGetAll('records');

    offlineMode = true;

    state.fields = fields.sort((a, b) => a.order - b.order);
    state.records = records;
    state.query = '';

    renderForm();
    renderFieldList();
    renderListView();

    showApp();

  } catch (error) {
    offlineMode = false;
    $('#loginError').textContent = error.message;
  }
}

async function handleLogin(e) {
  e.preventDefault();

  const passwordInput = $('#loginPassword');
  const errorBox = $('#loginError');
  const button = $('#loginBtn');

  errorBox.textContent = '';
  button.disabled = true;
  button.textContent = 'Conectando...';

  try {
    await login(passwordInput.value);

    await loadFields();
    await loadRecords();

    renderForm();
    renderFieldList();
    renderListView();

    passwordInput.value = '';
    showApp();

  } catch (err) {
    errorBox.textContent = err.message;
    logout();

  } finally {
    button.disabled = false;
    button.textContent = 'Iniciar sesión';
  }
}

const DB_NAME = 'capturadatos-db';
const DB_VERSION = 1;

/** Campos iniciales solicitados por el profesor. */
const BASE_FIELDS = [
  { id: 'matricula',       label: 'Matrícula',        type: 'text',   required: true,  base: true, order: 0 },
  { id: 'nombre',          label: 'Nombre',           type: 'text',   required: true,  base: true, order: 1 },
  { id: 'apellidoPaterno', label: 'Apellido paterno', type: 'text',   required: true,  base: true, order: 2 },
  { id: 'apellidoMaterno', label: 'Apellido materno', type: 'text',   required: false, base: true, order: 3 },
  { id: 'edad',            label: 'Edad',             type: 'number', required: true,  base: true, order: 4, integer: true, min: 0, max: 120 }
];

const TYPE_LABELS = { text: 'Texto', number: 'Numérico', image: 'Imagen' };

const state = {
  fields: [],      // definición de campos
  records: [],     // registros guardados
  editingId: null, // id del registro en edición
  images: {},      // imágenes pendientes del formulario {fieldId: dataURL|null}
  query: ''
};

/* ---------------------------- Utilidades ---------------------------- */

const $ = (sel) => document.querySelector(sel);

/** Normaliza texto: sin acentos, minúsculas, sin espacios extremos. */
const norm = (s) => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();

/** Crea elementos DOM de forma segura (sin innerHTML). */
function h(tag, attrs, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'dataset') Object.assign(el.dataset, v);
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
    else if (v === true) el.setAttribute(k, '');
    else el.setAttribute(k, v);
  }
  for (const kid of kids.flat()) {
    if (kid == null || kid === false) continue;
    el.append(kid);
  }
  return el;
}

function fmtDate(ts) {
  try {
    return new Intl.DateTimeFormat('es-MX', { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(ts));
  } catch (_) {
    return new Date(ts).toLocaleString();
  }
}

function todayStamp() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

let toastTimer = null;
function toast(msg, isError = false) {
  const t = $('#toast');
  t.textContent = msg;
  t.classList.toggle('error', isError);
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3600);
}

function confirmDialog(title, msg, okLabel = 'Eliminar') {
  return new Promise((resolve) => {
    const d = $('#confirm');
    $('#confirmTitle').textContent = title;
    $('#confirmMsg').textContent = msg;
    $('#confirmOk').textContent = okLabel;
    d.returnValue = '';
    d.addEventListener('close', () => resolve(d.returnValue === 'ok'), { once: true });
    d.showModal();
  });
}

function openLightbox(src, caption) {
  $('#lightboxImg').src = src;
  $('#lightboxImg').alt = caption;
  $('#lightboxCaption').textContent = caption;
  $('#lightbox').showModal();
}

function downloadBlob(blob, filename) {
  const a = h('a', { href: URL.createObjectURL(blob), download: filename });
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

/** Bloquea un botón mientras corre una tarea asíncrona. */
async function busy(btn, fn) {
  btn.disabled = true;
  try { return await fn(); } finally { btn.disabled = false; }
}

/* ------------------------------ IndexedDB ------------------------------ */

let db = null;

function openDB(name = DB_NAME) {  
  return new Promise((resolve, reject) => {
    if (!('indexedDB' in window)) {
      reject(new Error('Este navegador no permite guardar datos locales (IndexedDB).'));
      return;
    }
    const req = indexedDB.open(name, DB_VERSION);
    req.onupgradeneeded = () => {
      const d = req.result;
      if (!d.objectStoreNames.contains('fields')) d.createObjectStore('fields', { keyPath: 'id' });
      if (!d.objectStoreNames.contains('records')) d.createObjectStore('records', { keyPath: 'id', autoIncrement: true });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('No se pudo abrir la base de datos.'));
    req.onblocked = () => reject(new Error('La base de datos está bloqueada por otra pestaña.'));
  });
}

function getAccountDBName(userId) {
  if (typeof userId !== 'string' || !userId.trim()) {
    throw new Error('No se identificó la cuenta para el almacenamiento local.');
  }

  return `${DB_NAME}-${encodeURIComponent(userId)}`;
}

function dbGetAll(storeName) {
  return new Promise((resolve, reject) => {
    const r = db.transaction(storeName, 'readonly').objectStore(storeName).getAll();
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error);
  });
}

/** Ejecuta una escritura y espera a que la transacción termine realmente. */
function dbWrite(storeName, fn) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readwrite');
    let result;

    const req = fn(tx.objectStore(storeName));

    if (req) {
      req.onsuccess = () => {
        result = req.result;
      };
    }

    tx.oncomplete = () => resolve(result);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(
      tx.error || new Error('Operación cancelada (¿almacenamiento lleno?).')
    );
  });
}

/** Elimina un campo y sus valores en todos los registros (una sola transacción). */
function dbDeleteField(fieldId) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(['fields', 'records'], 'readwrite');
    tx.objectStore('fields').delete(fieldId);
    const cursorReq = tx.objectStore('records').openCursor();
    cursorReq.onsuccess = () => {
      const cursor = cursorReq.result;
      if (!cursor) return;
      const rec = cursor.value;
      if (rec.values && fieldId in rec.values) {
        delete rec.values[fieldId];
        cursor.update(rec);
      }
      cursor.continue();
    };
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error('Operación cancelada.'));
  });
}

/** Reemplaza TODOS los campos y registros (restauración de respaldo). */
function dbReplaceAll(fields, records) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(['fields', 'records'], 'readwrite');
    const fs = tx.objectStore('fields');
    const rs = tx.objectStore('records');
    fs.clear();
    rs.clear();
    for (const f of fields) fs.put(f);
    for (const r of records) rs.put(r);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error('Operación cancelada (¿almacenamiento lleno?).'));
  });
}

async function loadFields() {
  let fields;

  try {
    fields = await apiRequest('/api/fields');

    await dbWrite('fields', (s) => {
      s.clear();
      for (const field of fields) {
        s.put(field);
      }
    });

  } catch (err) {
    if (isOnline) throw err;

    fields = await dbGetAll('fields');

    if (fields.length === 0) {
      fields = BASE_FIELDS.slice();
    }
  }

  state.fields = fields.sort((a, b) => a.order - b.order);
}

async function loadRecords() {
  try {
    const records = await apiRequest('/api/records');

    await dbWrite('records', (s) => {
      s.clear();
      for (const record of records) {
        s.put(record);
      }
    });

    state.records = records;

  } catch (err) {
    if (isOnline) throw err;

    state.records = await dbGetAll('records');
  }
}

/* ------------------------------ Imágenes ------------------------------ */

/** Lee una imagen, la reduce (máx. 1024 px) y la devuelve como JPEG dataURL. */
function fileToDataURL(file, maxSide = 1024, quality = 0.85) {
  return new Promise((resolve, reject) => {
    if (!file.type.startsWith('image/')) {
      reject(new Error('El archivo seleccionado no es una imagen.'));
      return;
    }
    if (file.size > 15 * 1024 * 1024) {
      reject(new Error('La imagen es demasiado grande (máximo 15 MB).'));
      return;
    }
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      const scale = Math.min(1, maxSide / Math.max(img.naturalWidth, img.naturalHeight));
      const w = Math.max(1, Math.round(img.naturalWidth * scale));
      const hh = Math.max(1, Math.round(img.naturalHeight * scale));
      const canvas = document.createElement('canvas');
      canvas.width = w;
      canvas.height = hh;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, w, hh);
      ctx.drawImage(img, 0, 0, w, hh);
      URL.revokeObjectURL(url);
      resolve(canvas.toDataURL('image/jpeg', quality));
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error('No se pudo leer la imagen. Prueba con un archivo JPG o PNG.'));
    };
    img.src = url;
  });
}

/* ------------------------------ Formulario ------------------------------ */

function setError(fieldId, msg) {
  const err = document.getElementById('err_' + fieldId);
  const input = document.getElementById('in_' + fieldId);
  if (err) err.textContent = msg || '';
  if (input) {
    if (msg) input.setAttribute('aria-invalid', 'true');
    else input.removeAttribute('aria-invalid');
  }
}

function imageControl(f) {
  const preview = h('img', { class: 'thumb-preview', alt: 'Vista previa de ' + f.label });
  const box = h('div', { class: 'preview-box' }, preview);
  const input = h('input', { id: 'in_' + f.id, type: 'file', accept: 'image/*', class: 'file-input', 'aria-describedby': 'err_' + f.id });
  const clearBtn = h('button', { type: 'button', class: 'btn btn-ghost btn-sm' }, 'Quitar imagen');

  const refresh = () => {
    const v = state.images[f.id];
    box.hidden = !v;
    clearBtn.hidden = !v;
    if (v) preview.src = v;
  };

  input.addEventListener('change', async () => {
    const file = input.files && input.files[0];
    if (!file) return;
    try {
      state.images[f.id] = await fileToDataURL(file);
      setError(f.id, '');
    } catch (e) {
      setError(f.id, e.message);
    }
    input.value = ''; // permite volver a elegir el mismo archivo
    refresh();
  });
  clearBtn.addEventListener('click', () => {
    state.images[f.id] = null;
    input.value = '';
    refresh();
  });

  refresh();
  return h('div', { class: 'image-control' }, input, box, clearBtn);
}

/** Dibuja el formulario a partir de la definición de campos. */
function renderForm(values = {}) {
  const wrap = $('#fields');
  wrap.replaceChildren();
  state.images = {};

  for (const f of state.fields) {
    const id = 'in_' + f.id;
    const label = h('label', { for: id }, f.label, f.required ? h('span', { class: 'req', 'aria-hidden': 'true' }, ' *') : null);
    let control;

    if (f.type === 'image') {
      state.images[f.id] = values[f.id] || null;
      control = imageControl(f);
    } else {
      control = h('input', {
        id,
        name: f.id,
        type: 'text',
        inputmode: f.type === 'number' ? (f.integer ? 'numeric' : 'decimal') : 'text',
        maxlength: f.type === 'number' ? 20 : 120,
        autocomplete: 'off',
        'aria-describedby': 'err_' + f.id,
        'aria-required': f.required ? 'true' : null,
        placeholder: f.type === 'number' ? (f.integer ? 'Ej. 20' : 'Ej. 0001 o 12.5') : null,
        value: values[f.id] ?? ''
      });
    }

    wrap.append(
      h('div', { class: 'field' + (f.type === 'image' ? ' field-image' : '') },
        label, control, h('p', { class: 'error', id: 'err_' + f.id, role: 'alert' }))
    );
  }
}

/** Toma lo que el usuario ya escribió para no perderlo al redibujar. */
function snapshotForm() {
  const snap = {};
  for (const f of state.fields) {
    if (f.type === 'image') snap[f.id] = state.images[f.id] || null;
    else {
      const el = document.getElementById('in_' + f.id);
      if (el) snap[f.id] = el.value;
    }
  }
  return snap;
}

function validateField(f, raw) {
  if (f.type === 'image') {
    return f.required && !state.images[f.id] ? 'Selecciona una imagen.' : '';
  }
  const v = String(raw ?? '').trim();
  if (!v) return f.required ? 'Este campo es obligatorio.' : '';

  if (f.type === 'number') {
    if (f.integer) {
      if (!/^\d+$/.test(v)) return 'Escribe solo números enteros.';
      const n = Number(v);
      if ((f.min != null && n < f.min) || (f.max != null && n > f.max)) {
        return `Debe estar entre ${f.min} y ${f.max}.`;
      }
    } else if (!/^-?\d+([.,]\d+)?$/.test(v)) {
      return 'Escribe solo números (ej. 12, 0001 o 3.5).';
    }
  }
  return '';
}

async function onSubmit(e) {
  e.preventDefault();
  const values = {};
  let firstBad = null;

  for (const f of state.fields) {
    const el = document.getElementById('in_' + f.id);
    let msg = validateField(f, f.type === 'image' ? null : el.value);

    if (!msg && f.id === 'matricula' && el.value.trim()) {
      const dup = state.records.find((r) => r.id !== state.editingId && norm(r.values.matricula) === norm(el.value));
      if (dup) msg = 'Ya existe un registro con esta matrícula.';
    }

    setError(f.id, msg);
    if (msg) {
      firstBad = firstBad || el;
      continue;
    }
    if (f.type === 'image') values[f.id] = state.images[f.id] || null;
    else {
      let v = el.value.trim();
      if (f.type === 'number') v = v.replace(',', '.'); // conserva ceros a la izquierda (0001)
      values[f.id] = v;
    }
  }

  if (firstBad) {
    firstBad.focus();
    toast('Revisa los campos marcados en rojo.', true);
    return;
  }

  const existing = state.editingId != null
    ? state.records.find((r) => r.id === state.editingId)
    : null;

  const wasEditing = !!existing;

  try {
    if (offlineMode) {
      const now = new Date().toISOString();

      if (existing) {
        const updated = {
          ...existing,
          matricula: values.matricula,
          values,
          updatedAt: now
        };

        await dbWrite('records', (s) => s.put(updated));
      } else {
        const newRecord = {
          matricula: values.matricula,
          values,
          createdAt: now,
          updatedAt: now
        };

        const newId = await dbWrite(
          'records',
          (s) => s.add(newRecord)
        );

        newRecord.id = newId;
      }

      state.records = await dbGetAll('records');

    } else {
      const payload = {
        matricula: values.matricula,
        values
      };

      const saved = existing
        ? await apiRequest(`/api/records/${state.editingId}`, {
            method: 'PUT',
            body: JSON.stringify(payload)
          })
        : await apiRequest('/api/records', {
            method: 'POST',
            body: JSON.stringify(payload)
          });

      await dbWrite('records', (s) => s.put(saved));
      await loadRecords();
    }

  } catch (err) {
    toast(
      'No se pudo guardar: ' +
      (err && err.message ? err.message : 'error desconocido'),
      true
    );
    return;
  }

  resetForm();
  renderListView();
  toast(wasEditing ? 'Registro actualizado.' : 'Registro guardado correctamente.');
  if (wasEditing) showView('list');
  else $('#in_' + state.fields[0].id)?.focus();
}

function resetForm() {
  state.editingId = null;
  $('#formTitle').textContent = 'Nuevo registro';
  $('#saveBtn').textContent = 'Guardar registro';
  $('#editBanner').hidden = true;
  renderForm();
}

function startEdit(id) {
  const rec = state.records.find((r) => r.id === id);
  if (!rec) return;
  state.editingId = id;
  $('#formTitle').textContent = 'Editar registro';
  $('#saveBtn').textContent = 'Guardar cambios';
  $('#editBanner').hidden = false;
  renderForm(rec.values);
  showView('form');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/* ------------------------------ Listado ------------------------------ */

function matchesQuery(rec, q) {
  if (!q) return true;
  return state.fields.some((f) => f.type !== 'image' && norm(rec.values[f.id]).includes(q));
}

function renderListView() {
  const total = state.records.length;
  $('#countBadge').textContent = String(total);

  const q = norm(state.query);
  const list = state.records
    .filter((r) => matchesQuery(r, q))
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  $('#listMeta').textContent = q
    ? `${list.length} de ${total} registro${total === 1 ? '' : 's'}`
    : `${total} registro${total === 1 ? '' : 's'}`;

  $('#printTitle').textContent = `Registros – CapturaDatos (${new Date().toLocaleDateString('es-MX')})`;

  const wrap = $('#tableWrap');
  const empty = $('#emptyState');
  wrap.replaceChildren();
  empty.replaceChildren();

  if (list.length === 0) {
    wrap.hidden = true;
    empty.hidden = false;

    if (total === 0) {
      empty.append(
        h('strong', null, 'Aún no hay registros'),
        h('span', null, 'Captura el primero para verlo aquí.'),
        h('div', null,
          h('button', {
            type: 'button',
            class: 'btn btn-primary',
            onclick: () => showView('form')
          }, 'Crear registro')
        )
      );
    } else {
      empty.append(
        h('strong', null, 'Sin resultados'),
        h('span', null, 'Ningún registro coincide con tu búsqueda.')
      );
    }
    return;
  }

  wrap.hidden = false;
  empty.hidden = true;

  const head = h('tr', null,
    state.fields.map((f) => h('th', { scope: 'col' }, f.label)),
    h('th', { scope: 'col' }, 'Registrado'),
    h('th', { scope: 'col', class: 'cell-actions' }, 'Acciones')
  );

  const rows = list.map((r) => {
    const cells = state.fields.map((f) => {
      const v = r.values[f.id];
      let content;

      if (f.type === 'image') {
        content = v
          ? h('button', {
              type: 'button',
              class: 'thumb',
              title: 'Ver imagen completa',
              'aria-label': `Ver ${f.label}`,
              onclick: () => openLightbox(v, `${f.label} – ${r.values.matricula ?? ''}`)
            }, h('img', { src: v, alt: f.label, loading: 'lazy' }))
          : h('span', { class: 'dash' }, '—');
      } else {
        content = v ? String(v) : h('span', { class: 'dash' }, '—');
      }

      return h('td', { dataset: { label: f.label } }, content);
    });

    cells.push(
      h('td', { dataset: { label: 'Registrado' } }, fmtDate(r.createdAt))
    );

    cells.push(
      h('td', { class: 'cell-actions' },
        h('button', {
          type: 'button',
          class: 'btn btn-outline btn-sm',
          onclick: () => startEdit(r.id)
        }, 'Editar'),
        h('button', {
          type: 'button',
          class: 'btn btn-danger-outline btn-sm',
          onclick: () => deleteRecord(r.id)
        }, 'Eliminar')
      )
    );

    return h('tr', null, cells);
  });

  wrap.append(h('table', null, h('thead', null, head), h('tbody', null, rows)));
}

async function deleteRecord(id) {
  const rec = state.records.find((r) => r.id === id);
  if (!rec) return;

  const who = [rec.values.nombre, rec.values.apellidoPaterno]
    .filter(Boolean).join(' ') || 'este registro';

  const ok = await confirmDialog(
    '¿Eliminar registro?',
    `Se eliminará "${who}" (matrícula ${rec.values.matricula ?? 's/n'}). Esta acción no se puede deshacer.`
  );

  if (!ok) return;

  try {
    if (offlineMode) {
      await dbWrite('records', (s) => s.delete(id));
      state.records = await dbGetAll('records');

    } else {
      await apiRequest(`/api/records/${id}`, {
        method: 'DELETE'
      });

      await dbWrite('records', (s) => s.delete(id));
      await loadRecords();
    }

    if (state.editingId === id) resetForm();

    renderListView();
    toast('Registro eliminado.');

  } catch (err) {
    toast('No se pudo eliminar: ' + err.message, true);
  }
}

/* ------------------------- Exportar: Excel, CSV, PDF ------------------------- */

function sortedRecords() {
  return state.records.slice().sort((a, b) => a.createdAt - b.createdAt);
}

/** Excel real (.xlsx). Los números con ceros a la izquierda (0001) se guardan como texto. */
function exportXLSX() {
  if (state.records.length === 0) {
    toast('No hay registros para descargar.', true);
    return;
  }
  const columns = state.fields.map((f) => ({
    header: f.label,
    width: f.type === 'image' ? 16 : Math.max(14, Math.min(34, f.label.length + 6))
  }));
  columns.push({ header: 'Fecha de registro', width: 22 });

  const rows = sortedRecords().map((r) => {
    const cells = state.fields.map((f) => {
      const v = r.values[f.id];
      if (f.type === 'image') return { v: v ? 'Imagen adjunta' : '', t: 's' };
      if (f.type === 'number') {
        const s = String(v ?? '');
        const keepAsText = /^0\d/.test(s) || /^-?0\d/.test(s) || s.length > 15; // 0001 → texto
        return { v: s, t: keepAsText || !/^-?\d+(\.\d+)?$/.test(s) ? 's' : 'n' };
      }
      return { v: v ?? '', t: 's' };
    });
    cells.push({ v: fmtDate(r.createdAt), t: 's' });
    return cells;
  });

  const bytes = window.buildXlsx({ sheetName: 'Registros', columns, rows });
  const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  downloadBlob(blob, `registros_${todayStamp()}.xlsx`);
  toast('Archivo Excel generado.');
}

function csvCell(value, type) {
  let s = String(value ?? '');
  if (type === 'number' && /^0\d+/.test(s)) return `"=""${s}"""`; // Excel conserva 0001
  if (type !== 'number' && /^[=+\-@]/.test(s)) s = "'" + s;       // evita fórmulas en Excel
  return '"' + s.replace(/"/g, '""') + '"';
}

function exportCSV() {
  if (state.records.length === 0) {
    toast('No hay registros para descargar.', true);
    return;
  }
  const header = state.fields.map((f) => csvCell(f.label, 'text')).concat(csvCell('Fecha de registro', 'text'));
  const lines = [header.join(',')];
  for (const r of sortedRecords()) {
    const cells = state.fields.map((f) =>
      f.type === 'image' ? csvCell(r.values[f.id] ? 'Imagen adjunta' : '', 'text') : csvCell(r.values[f.id], f.type)
    );
    cells.push(csvCell(fmtDate(r.createdAt), 'text'));
    lines.push(cells.join(','));
  }
  const blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  downloadBlob(blob, `registros_${todayStamp()}.csv`);
  toast('Archivo CSV generado.');
}

/** PDF: abre el diálogo de impresión sobre la lista; se elige "Guardar como PDF". */
function exportPDF() {
  if (state.records.length === 0) {
    toast('No hay registros para imprimir.', true);
    return;
  }
  showView('list');
  toast('En el diálogo elige “Guardar como PDF”.');
  setTimeout(() => window.print(), 350);
}

/* ------------------------- Respaldo en archivo (.json) ------------------------- */

function buildBackup() {
  return {
    app: 'capturadatos',
    version: 1,
    exportedAt: new Date().toISOString(),
    fields: state.fields,
    records: state.records
  };
}

/** Valida y limpia un respaldo; lanza Error con mensaje claro si es inválido. */
function parseBackup(data) {
  if (!data || data.app !== 'capturadatos' || !Array.isArray(data.fields) || !Array.isArray(data.records)) {
    throw new Error('El archivo no es un respaldo válido de CapturaDatos.');
  }
  const types = ['text', 'number', 'image'];
  const fields = data.fields.map((f, i) => {
    if (!f || typeof f.id !== 'string' || typeof f.label !== 'string' || !types.includes(f.type)) {
      throw new Error('El respaldo tiene campos dañados.');
    }
    return { ...f, order: Number.isFinite(f.order) ? f.order : i };
  });
  const records = data.records.map((r) => {
    if (!r || typeof r.values !== 'object' || r.values === null) throw new Error('El respaldo tiene registros dañados.');
    return r;
  });
  return { fields, records };
}

async function applyBackup(data, sourceLabel) {
  const { fields, records } = parseBackup(data);

  const destination = offlineMode
    ? 'el dispositivo'
    : 'el servidor';

  const ok = await confirmDialog(
    '¿Restaurar respaldo?',
    `Se reemplazarán los datos actuales de ${destination} por ${records.length} registro${records.length === 1 ? '' : 's'} de ${sourceLabel}. Esta acción no se puede deshacer.`,
    'Restaurar'
  );

  if (!ok) return false;

  if (offlineMode) {
    await dbWrite('fields', (s) => {
      s.clear();
      for (const field of fields) s.put(field);
    });

    await dbWrite('records', (s) => {
      s.clear();
      for (const record of records) s.put(record);
    });

    state.fields = await dbGetAll('fields');
    state.records = await dbGetAll('records');

  } else {
    await apiRequest('/api/backup/restore', {
      method: 'POST',
      body: JSON.stringify({ fields, records })
    });

    await loadFields();
    await loadRecords();
  }

  resetForm();
  renderFieldList();
  renderListView();

  return true;
}

function jsonBackup() {
  const blob = new Blob([JSON.stringify(buildBackup())], { type: 'application/json' });
  downloadBlob(blob, `capturadatos-respaldo_${todayStamp()}.json`);
  toast('Respaldo descargado.');
}

async function onJsonRestore(e) {
  const file = e.target.files && e.target.files[0];
  e.target.value = '';
  if (!file) return;
  try {
    const data = JSON.parse(await file.text());
    if (await applyBackup(data, 'el archivo')) toast('Respaldo restaurado.');
  } catch (err) {
    toast(err instanceof SyntaxError ? 'El archivo no es un JSON válido.' : err.message, true);
  }
}

/* ---------------------------------------------------------------------
   Google: sesión y Drive
   --------------------------------------------------------------------- */

const google = {
  token: null,
  expires: 0,
  client: null,
  profile: null // { name, email, picture }
};

try { google.profile = JSON.parse(localStorage.getItem('cd_google_profile') || 'null'); } catch (_) { /* ignorar */ }

function loadGoogleScript() {
  return new Promise((resolve, reject) => {
    if (window.google?.accounts) {
      resolve();
      return;
    }

    const scriptUrl = 'https://accounts.google.com/gsi/client';
    const existingScript = document.querySelector(
      `script[src="${scriptUrl}"]`
    );

    if (existingScript) {
      if (existingScript.dataset.loaded === 'true') {
        resolve();
        return;
      }

      existingScript.addEventListener('load', () => {
        existingScript.dataset.loaded = 'true';
        resolve();
      }, { once: true });

      existingScript.addEventListener('error', () => {
        reject(new Error('No se pudo cargar el servicio de Google.'));
      }, { once: true });

      return;
    }

    const script = document.createElement('script');
    script.src = scriptUrl;
    script.async = true;
    script.defer = true;

    script.onload = () => {
      script.dataset.loaded = 'true';
      resolve();
    };

    script.onerror = () => {
      reject(new Error('No se pudo cargar el servicio de Google.'));
    };

    document.head.appendChild(script);
  });
}

async function initializeGoogleSignIn() {
  await loadGoogleScript();

  if (!window.google.accounts.id) {
    throw new Error('No se pudo cargar el inicio de sesión de Google.');
  }

  window.google.accounts.id.initialize({
    client_id: GOOGLE_CLIENT_ID,

    callback: async (response) => {
      const errorBox = $('#loginError');
      errorBox.textContent = '';

      try {
        if (!response.credential) {
          throw new Error('Google no devolvió las credenciales.');
        }

        await loginWithGoogle(response.credential);

        db = await openDB(getAccountDBName(authUser.id));

        await loadFields();
        await loadRecords();

        renderForm();
        renderFieldList();
        renderListView();

        showApp();
      } catch (error) {
        errorBox.textContent = error.message;
      }
    }
  });

  window.google.accounts.id.renderButton(
    $('#googleSignInButton'),
    {
      type: 'standard',
      theme: 'outline',
      size: 'large',
      text: 'signin_with',
      shape: 'rectangular',
      width: 280
    }
  );
}

/** Pide un token de acceso (abre la ventana de Google si hace falta). */
async function requestToken(prompt, loginHint) {
  if (!GOOGLE_CLIENT_ID) {
    throw new Error('Falta configurar el Client ID de Google.');
  }

  await loadGoogleScript();

  return new Promise((resolve, reject) => {
    google.client = window.google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope: GOOGLE_SCOPES,
      prompt,
      login_hint: loginHint || undefined,

      callback: (resp) => {
        if (resp.error) {
          reject(new Error(
            'Google rechazó el acceso: ' +
            (resp.error_description || resp.error)
          ));
          return;
        }

        google.token = resp.access_token;
        google.expires = Date.now() +
          (Number(resp.expires_in) || 3600) * 1000;

        resolve(google.token);
      },

      error_callback: (err) => {
        reject(new Error(
          err && err.type === 'popup_closed'
            ? 'Cerraste la ventana de Google.'
            : 'No se pudo abrir el inicio de sesión de Google.'
        ));
      }
    });

    google.client.requestAccessToken();
  });
}

async function getToken() {
  if (google.token && Date.now() < google.expires - 30000) {
    return google.token;
  }

  const email = authUser?.email || google.profile?.email || '';

  return requestToken(
    email ? '' : 'select_account',
    email
  );
}

async function googleFetch(url, opts = {}) {
  const token = await getToken();
  const res = await fetch(url, { ...opts, headers: { ...(opts.headers || {}), Authorization: 'Bearer ' + token } });
  if (res.status === 401) {
    google.token = null;
    throw new Error('La sesión de Google expiró. Inténtalo de nuevo.');
  }
  if (!res.ok) {
    const detail = await res.text();
    console.error('Error detallado de Google:', detail);
    throw new Error(`Google respondió con error ${res.status}: ${detail}`);
  }
  return res;
}

async function googleLogin(email = '') {
  const accountEmail = email || authUser?.email || '';

  await requestToken(
    accountEmail ? '' : 'select_account',
    accountEmail
  );

  const res = await googleFetch(
    'https://www.googleapis.com/oauth2/v3/userinfo'
  );

  const info = await res.json();

  google.profile = {
    name: info.name || '',
    email: info.email || '',
    picture: info.picture || ''
  };

  try {
    localStorage.setItem(
      'cd_google_profile',
      JSON.stringify(google.profile)
    );
  } catch (_) {}

  renderAccount();
  toast('Cuenta de Google conectada.');
}

function googleLogout() {
  const token = google.token;
  if (token && window.google && window.google.accounts) {
    try { window.google.accounts.oauth2.revoke(token, () => {}); } catch (_) { /* ignorar */ }
  }
  google.token = null;
  google.expires = 0;
  google.profile = null;
  try { localStorage.removeItem('cd_google_profile'); } catch (_) { /* ignorar */ }
  renderAccount();
  toast('Sesión cerrada.');
}

function renderAccount() {
  const p = google.profile;
  $('#cfgWarn').hidden = !!GOOGLE_CLIENT_ID;
  $('#googleLoginBtn').hidden = !!p;
  $('#googleLogoutBtn').hidden = !p;
  $('#accName').textContent = p ? (p.name || 'Cuenta de Google') : 'Sin sesión iniciada';
  $('#accEmail').textContent = p ? p.email : 'Conecta tu cuenta para usar Google Drive.';
  const av = $('#accAvatar');
  av.replaceChildren();
  if (p && p.picture) av.append(h('img', { src: p.picture, alt: '', referrerpolicy: 'no-referrer' }));
  else av.textContent = p ? (p.name || p.email || '?').charAt(0).toUpperCase() : '?';
}

async function findBackupFile() {
  const q = encodeURIComponent(`name='${BACKUP_FILE_NAME}' and trashed=false`);
  const res = await googleFetch(`https://www.googleapis.com/drive/v3/files?q=${q}&spaces=drive&fields=files(id,modifiedTime)&orderBy=modifiedTime desc`);
  const data = await res.json();
  return (data.files && data.files[0]) || null;
}

async function driveBackup() {
  const json = JSON.stringify(buildBackup());
  const existing = await findBackupFile();
  const boundary = 'cd' + Date.now().toString(36);
  const meta = existing ? {} : { name: BACKUP_FILE_NAME, mimeType: 'application/json' };
  const body =
    `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n` +
    `--${boundary}\r\nContent-Type: application/json\r\n\r\n${json}\r\n--${boundary}--`;
  const url = existing
    ? `https://www.googleapis.com/upload/drive/v3/files/${existing.id}?uploadType=multipart`
    : 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart';
  await googleFetch(url, {
    method: existing ? 'PATCH' : 'POST',
    headers: { 'Content-Type': `multipart/related; boundary=${boundary}` },
    body
  });
  const when = fmtDate(Date.now());
  $('#driveStatus').textContent = `Último respaldo en Drive: ${when} (${state.records.length} registro${state.records.length === 1 ? '' : 's'}).`;
  toast('Respaldo guardado en Google Drive.');
}

async function driveRestore() {
  const file = await findBackupFile();
  if (!file) { toast('No se encontró ningún respaldo en tu Drive.', true); return; }
  const res = await googleFetch(`https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`);
  const data = await res.json();
  if (await applyBackup(data, `Google Drive (${fmtDate(file.modifiedTime)})`)) toast('Respaldo restaurado desde Drive.');
}

/** Envuelve una acción de Google mostrando errores como aviso. */
function guarded(btn, fn) {
  return () => busy(btn, async () => {
    try {
      if (!GOOGLE_CLIENT_ID) { renderAccount(); toast('Falta configurar el Client ID de Google (ver README).', true); return; }
      await fn();
    } catch (err) {
      toast(err.message || 'Ocurrió un error con Google.', true);
    }
  });
}

/* ------------------------- Instalación (PWA) ------------------------- */

let deferredInstall = null;
const isStandalone = () => window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
const isIOS = () => /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

function updateInstallUI() {
  const installed = isStandalone();
  $('#installBtnTop').hidden = installed || !deferredInstall;
  $('#installBtn').hidden = installed;
  $('#installStatus').textContent = installed
    ? 'CapturaDatos ya está instalada en este dispositivo.'
    : 'Puedes instalar CapturaDatos como una app y abrirla desde la pantalla de inicio.';
}

async function installApp() {
  if (deferredInstall) {
    deferredInstall.prompt();
    try { await deferredInstall.userChoice; } catch (_) { /* ignorar */ }
    deferredInstall = null;
    updateInstallUI();
  } else if (isIOS()) {
    toast('En Safari: botón Compartir → “Agregar a pantalla de inicio”.');
  } else {
    toast('Abre el menú del navegador (⋮) y elige “Instalar aplicación”.');
  }
}

window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault();
  deferredInstall = e;
  updateInstallUI();
});
window.addEventListener('appinstalled', () => {
  deferredInstall = null;
  updateInstallUI();
  toast('Aplicación instalada.');
});

if ('serviceWorker' in navigator && (location.protocol === 'https:' || ['localhost', '127.0.0.1'].includes(location.hostname))) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('sw.js').catch(() => { /* sin soporte: la app funciona igual */ });
  });
}

/* ------------------------- Gestión de campos ------------------------- */

function renderFieldList() {
  const ul = $('#fieldList');
  ul.replaceChildren();
  for (const f of state.fields) {
    ul.append(h('li', null,
      h('div', null,
        h('div', { class: 'field-name' }, f.label),
        h('div', { class: 'badges' },
          h('span', { class: 'badge badge-' + f.type }, TYPE_LABELS[f.type]),
          f.required ? h('span', { class: 'badge badge-req' }, 'Obligatorio') : null,
          f.base ? h('span', { class: 'badge' }, 'Campo base') : null)),
      f.base ? null : h('button', { type: 'button', class: 'btn btn-danger-outline btn-sm', onclick: () => removeField(f.id) }, 'Eliminar campo')
    ));
  }
}

async function onAddField(e) {
  e.preventDefault();

  const label = $('#newLabel').value.trim().replace(/\s+/g, ' ');
  const type = $('#newType').value;
  const required = $('#newRequired').checked;
  const err = $('#newLabelErr');

  err.textContent = '';
  $('#newLabel').removeAttribute('aria-invalid');

  let msg = '';

  if (label.length < 2) {
    msg = 'Escribe un nombre de al menos 2 caracteres.';
  } else if (state.fields.some((f) => norm(f.label) === norm(label))) {
    msg = 'Ya existe un campo con ese nombre.';
  }

  if (msg) {
    err.textContent = msg;
    $('#newLabel').setAttribute('aria-invalid', 'true');
    $('#newLabel').focus();
    return;
  }

  const field = {
    id: offlineMode ? crypto.randomUUID() : undefined,
    label,
    type,
    required,
    base: false,
    active: true,
    order: Math.max(...state.fields.map((f) => f.order)) + 1
  };

  if (type === 'number') field.integer = false;

  $('#fieldForm').reset();

  try {
    const snap = snapshotForm();

    if (offlineMode) {
      await dbWrite('fields', (s) => s.put(field));
      state.fields = await dbGetAll('fields');
      state.fields.sort((a, b) => a.order - b.order);
    } else {
      delete field.id;
      await apiRequest('/api/fields', {
        method: 'POST',
        body: JSON.stringify(field)
      });

      await loadFields();
    }

    renderForm(snap);
    renderFieldList();
    renderListView();

    toast(`Campo "${label}" agregado al formulario.`);
  } catch (error) {
    toast('No se pudo agregar el campo: ' + error.message, true);
  }
}

async function removeField(id) {
  const f = state.fields.find((x) => x.id === id);
  if (!f || f.base) return;

  const used = state.records.filter(
    (r) => r.values[id] != null && r.values[id] !== ''
  ).length;

  const detail = used > 0
    ? `El campo dejará de aparecer en el formulario, pero sus datos se conservarán en ${used} registro${used === 1 ? '' : 's'}.`
    : 'El campo dejará de aparecer en el formulario.';

  const ok = await confirmDialog(
    `¿Eliminar el campo "${f.label}"?`,
    detail
  );

  if (!ok) return;

  try {
    const snap = snapshotForm();
    delete snap[id];

    if (offlineMode) {
      await dbWrite('fields', (s) => s.delete(id));
      state.fields = await dbGetAll('fields');
      state.fields.sort((a, b) => a.order - b.order);
    } else {
      await apiRequest(`/api/fields/${encodeURIComponent(id)}`, {
        method: 'DELETE'
      });

      await loadFields();
    }

    renderForm(snap);
    renderFieldList();
    renderListView();

    toast('Campo eliminado. Los datos históricos se conservaron.');
  } catch (err) {
    toast('No se pudo eliminar el campo: ' + err.message, true);
  }
}

/* ------------------------------ Navegación ------------------------------ */

function showView(name) {
  for (const v of ['form', 'list', 'fields', 'backup']) {
    $('#view-' + v).hidden = v !== name;
  }
  document.querySelectorAll('.tab').forEach((t) => {
    if (t.dataset.view === name) t.setAttribute('aria-current', 'page');
    else t.removeAttribute('aria-current');
  });
}

/* ------------------------------ Arranque ------------------------------ */

async function init() {
  document.querySelectorAll('.tab').forEach((t) => {
    t.addEventListener('click', () => showView(t.dataset.view));
  });

  $('#recordForm').addEventListener('submit', onSubmit);
  $('#resetBtn').addEventListener('click', () => {
    resetForm();
    toast('Formulario limpio.');
  });
  $('#cancelEditBtn').addEventListener('click', () => {
    resetForm();
    showView('list');
  });
  $('#fieldForm').addEventListener('submit', onAddField);

  $('#offlineLoginBtn').addEventListener('click', startOfflineSession);

  $('#apiLogoutBtn').addEventListener('click', () => {
    logout();
    resetForm();
    showLogin();
  });

  $('#searchInput').addEventListener('input', (e) => {
    state.query = e.target.value;
    renderListView();
  });

  $('#exportXlsxBtn').addEventListener('click', exportXLSX);
  $('#exportXlsxBtn2').addEventListener('click', exportXLSX);
  $('#exportCsvBtn').addEventListener('click', exportCSV);
  $('#printBtn').addEventListener('click', exportPDF);
  $('#printBtn2').addEventListener('click', exportPDF);

  $('#jsonBackupBtn').addEventListener('click', jsonBackup);
  $('#jsonRestoreInput').addEventListener('change', onJsonRestore);

  const googleLoginBtn = $('#googleLoginBtn');
  const googleLogoutBtn = $('#googleLogoutBtn');
  const backupBtn = $('#driveBackupBtn');
  const restoreBtn = $('#driveRestoreBtn');

  if (googleLoginBtn) {
    googleLoginBtn.addEventListener(
      'click',
      guarded(googleLoginBtn, googleLogin)
    );
  }

  if (googleLogoutBtn) {
    googleLogoutBtn.addEventListener('click', googleLogout);
  }

  if (backupBtn) {
    backupBtn.addEventListener('click', guarded(backupBtn, async () => {
      if (!google.profile) await googleLogin();
      await driveBackup();
    }));
  }

  if (restoreBtn) {
    restoreBtn.addEventListener('click', guarded(restoreBtn, async () => {
      if (!google.profile) await googleLogin();
      await driveRestore();
    }));
  }

  renderAccount();

  try {
    await initializeGoogleSignIn();
  } catch (error) {
    console.error('Google Sign-In:', error);
  }

  $('#installBtn').addEventListener('click', installApp);
  $('#installBtnTop').addEventListener('click', installApp);
  updateInstallUI();

  try {
    if (navigator.storage && navigator.storage.persist) {
      navigator.storage.persist().catch(() => {});
    }

    if (authToken && authUser) {
      db = await openDB(getAccountDBName(authUser.id));

      try {
        await loadFields();
        await loadRecords();

        renderForm();
        renderFieldList();
        renderListView();

        showApp();
        return;

      } catch (err) {
        console.error('Error al recuperar la sesión:', err);
        logout();
      }
    }

    showLogin();

  } catch (err) {
    console.error('Error al iniciar CapturaDatos:', err);
    showLogin();
  }
}

init();
