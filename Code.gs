/**
 * Chemas JC Sport — puente de sincronización con Google Sheets
 * Pega este código en la hoja: Extensiones > Apps Script.
 * 1) Cambia CLAVE por una contraseña tuya.
 * 2) Implementar > Nueva implementación > Aplicación web
 *    Ejecutar como: Yo   |   Quién tiene acceso: Cualquier persona
 * 3) Copia la URL que termina en /exec y pégala en la app.
 *
 * La hoja es la base de datos. NO edites las filas a mano: la fuente real es la
 * columna "json"; las demás columnas son solo para que puedas leer y filtrar.
 */
const SHEET_ID = '1jbGphucehy4Nf_JnDQvWThTPoFnLI07J8NddrPJbuFE';
const CLAVE = 'CAMBIA-ESTA-CLAVE';

const STORES = {
  products: 'Productos', customers: 'Clientes', sales: 'Ventas', payments: 'Abonos',
  suppliers: 'Proveedores', supplierPayments: 'PagosProveedores',
  supplierLots: 'Consignaciones', orders: 'Pedidos'
};
const BASE = ['id', 'json', 'actualizado_ms', 'eliminado', 'servidor_ms']; // columnas A-E
const NB = BASE.length;

function doGet() { return ContentService.createTextOutput('Chemas JC Sport: sincronización activa'); }

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const req = JSON.parse(e.postData.contents);
    if (req.key !== CLAVE) return out_({ ok: false, error: 'Clave incorrecta' });
    const since = Number(req.since) || 0;
    const props = PropertiesService.getScriptProperties();
    const changes = Array.isArray(req.changes) ? req.changes : [];
    const ss = SpreadsheetApp.openById(SHEET_ID);
    if (req.action === 'ping') { Object.keys(STORES).forEach(s => sheet_(ss, s)); return out_({ ok: true }); }

    const now = Date.now();
    if (changes.length) {
      const by = {};
      changes.forEach(c => { if (STORES[c.store]) (by[c.store] = by[c.store] || []).push(c); });
      Object.keys(by).forEach(store => write_(ss, store, by[store], now));
      props.setProperty('lastWrite', String(now));
    }
    const lastWrite = Number(props.getProperty('lastWrite')) || 0;
    if (since >= lastWrite) return out_({ ok: true, changes: [], cursor: Math.max(since, lastWrite) });

    const result = [];
    let cursor = since;
    Object.keys(STORES).forEach(store => {
      const sh = ss.getSheetByName(STORES[store]);
      if (!sh || sh.getLastRow() < 2) return;
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, NB).getValues();
      rows.forEach(r => {
        const ts = Number(r[4]) || 0;
        if (ts > cursor) cursor = ts;
        if (ts > since && r[0] !== '') {
          const del = Number(r[3]) === 1;
          let data = null;
          if (!del) { try { data = JSON.parse(r[1]); } catch (err) { return; } }
          result.push({ store: store, id: String(r[0]), deleted: del, updatedAt: Number(r[2]) || 0, data: data });
        }
      });
    });
    return out_({ ok: true, changes: result, cursor: cursor });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function sheet_(ss, store) {
  let sh = ss.getSheetByName(STORES[store]);
  if (!sh) {
    sh = ss.insertSheet(STORES[store]);
    sh.getRange(1, 1, 1, NB).setValues([BASE]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, NB).setFontWeight('bold').setBackground('#101827').setFontColor('#ffffff');
  }
  return sh;
}

function isDateKey_(k, v) { return typeof v === 'number' && v > 946684800000 && /(At|Date|date)$|^date$|^fecha/.test(k); }

function flat_(data) {
  const o = {};
  Object.keys(data || {}).forEach(k => {
    if (k === 'photo' || k === 'updatedAt' || k === 'id') return;
    const v = data[k];
    if (v === null || v === undefined) o[k] = '';
    else if (isDateKey_(k, v)) o[k] = new Date(v);
    else if (typeof v === 'object') o[k] = JSON.stringify(v).slice(0, 500);
    else o[k] = (typeof v === 'string' && v.length > 2000) ? v.slice(0, 2000) : v;
  });
  return o;
}

function write_(ss, store, list, now) {
  const sh = sheet_(ss, store);
  const last = sh.getLastRow();
  const existing = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 3).getValues().forEach((r, i) => { existing[String(r[0])] = { row: i + 2, upd: Number(r[2]) || 0 }; });

  // encabezados extra (legibles)
  const lastCol = Math.max(sh.getLastColumn(), NB);
  let extra = lastCol > NB ? sh.getRange(1, NB + 1, 1, lastCol - NB).getValues()[0].map(String) : [];
  const prepared = [];
  list.forEach(c => {
    const prev = existing[String(c.id)];
    if (prev && prev.upd > Number(c.updatedAt || 0)) return; // gana el cambio más reciente
    const f = c.deleted ? {} : flat_(c.data);
    Object.keys(f).forEach(k => { if (extra.indexOf(k) < 0) extra.push(k); });
    const json = c.deleted ? '' : JSON.stringify(Object.assign({}, c.data, { photo: undefined }));
    if (json.length > 49000) throw new Error('Registro demasiado grande: ' + store + ' ' + c.id);
    prepared.push({ c: c, prev: prev, f: f, json: json });
  });
  if (!prepared.length) return;
  if (extra.length) sh.getRange(1, NB + 1, 1, extra.length).setValues([extra]).setFontWeight('bold').setBackground('#101827').setFontColor('#ffffff');
  const width = NB + extra.length;
  const mk = p => {
    const row = [String(p.c.id), p.json, Number(p.c.updatedAt) || now, p.c.deleted ? 1 : 0, now];
    extra.forEach(k => row.push(p.f[k] === undefined ? '' : p.f[k]));
    return row;
  };
  const fresh = [];
  const seen = {};
  prepared.forEach(p => {
    const id = String(p.c.id);
    if (p.prev) sh.getRange(p.prev.row, 1, 1, width).setValues([mk(p)]);
    else if (seen[id] !== undefined) fresh[seen[id]] = mk(p);
    else { seen[id] = fresh.length; fresh.push(mk(p)); }
  });
  if (fresh.length) {
    const start = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(start, 1, fresh.length, width).setValues(fresh);
  }
}
