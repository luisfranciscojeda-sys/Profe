/**
 * Detectives del Pensamiento — Modo en vivo
 * Base de datos en Google Sheets para salas en tiempo real.
 *
 * INSTALACIÓN (una sola vez):
 * 1. Crea una hoja nueva en Google Sheets (con tu cuenta personal de Gmail).
 * 2. Menú Extensiones → Apps Script. Borra lo que aparece y pega TODO este código.
 * 3. Cambia la clave de abajo (CLAVE_DOCENTE) por una tuya.
 * 4. Guarda (ícono de disquete).
 * 5. Implementar → Nueva implementación → tipo «Aplicación web».
 *      Ejecutar como: Yo
 *      Quién tiene acceso: Cualquier persona
 *    → Implementar → Autorizar acceso → copia la «URL de la aplicación web».
 * 6. Pega esa URL y tu clave en la app: Panel docente → 🔴 En vivo.
 */

const CLAVE_DOCENTE = 'cambia-esta-clave';

function doGet(e) {
  let out;
  try { out = route_(e.parameter || {}); }
  catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function route_(p) {
  switch (p.a) {
    case 'ping':       return { ok: true, docente: p.clave === CLAVE_DOCENTE };
    case 'crear':      docente_(p); return crear_(p.mision || 'm1');
    case 'iniciar':    docente_(p); return cambiarEstado_(p.sala, 'jugando');
    case 'terminar':   docente_(p); return cambiarEstado_(p.sala, 'terminado');
    case 'resultados': docente_(p); return resultados_(p.sala);
    case 'unirse':     return unirse_(p.sala, p.nombre, p.integrantes);
    case 'estado':     return estado_(p.sala);
    case 'responder':  return responder_(p);
    default: throw new Error('Acción desconocida');
  }
}

function docente_(p) { if (p.clave !== CLAVE_DOCENTE) throw new Error('Clave docente incorrecta'); }

function hoja_(nombre, encabezados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(nombre);
  if (!s) { s = ss.insertSheet(nombre); s.appendRow(encabezados); s.setFrozenRows(1); }
  return s;
}
const SALAS = () => hoja_('Salas', ['sala', 'estado', 'mision', 'creada', 'iniciada']);
const EQUIPOS = () => hoja_('Equipos', ['sala', 'equipoId', 'equipo', 'integrantes', 'unido']);
const RESP = () => hoja_('Respuestas', ['fecha', 'sala', 'equipoId', 'equipo', 'item', 'caso', 'eleccion', 'correcta', 'pista', 'segundos']);

function conLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function salaInfo_(sala) {
  const c = CacheService.getScriptCache();
  const hit = c.get('sala_' + sala);
  if (hit) return JSON.parse(hit);
  const rows = SALAS().getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(sala)) {
      const info = { fila: i + 1, estado: rows[i][1], mision: rows[i][2] };
      c.put('sala_' + sala, JSON.stringify(info), 21600);
      return info;
    }
  }
  return null;
}

function crear_(mision) {
  return conLock_(() => {
    const s = SALAS();
    const usadas = new Set(s.getDataRange().getValues().slice(1).map(r => String(r[0])));
    let sala;
    do { sala = String(Math.floor(1000 + Math.random() * 9000)); } while (usadas.has(sala));
    s.appendRow([sala, 'espera', mision, new Date(), '']);
    CacheService.getScriptCache().remove('sala_' + sala);
    return { ok: true, sala: sala, mision: mision };
  });
}

function cambiarEstado_(sala, estado) {
  return conLock_(() => {
    const info = salaInfo_(sala);
    if (!info) throw new Error('La sala no existe');
    SALAS().getRange(info.fila, 2).setValue(estado);
    if (estado === 'jugando') SALAS().getRange(info.fila, 5).setValue(new Date());
    info.estado = estado;
    CacheService.getScriptCache().put('sala_' + sala, JSON.stringify(info), 21600);
    return { ok: true, estado: estado };
  });
}

function equiposDe_(sala) {
  const c = CacheService.getScriptCache();
  const hit = c.get('eq_' + sala);
  if (hit) return JSON.parse(hit);
  const lista = EQUIPOS().getDataRange().getValues().slice(1)
    .filter(r => String(r[0]) === String(sala))
    .map(r => ({ id: String(r[1]), nombre: String(r[2]) }));
  c.put('eq_' + sala, JSON.stringify(lista), 21600);
  return lista;
}

function unirse_(sala, nombre, integrantes) {
  nombre = String(nombre || '').trim().slice(0, 30);
  if (!nombre) throw new Error('Falta el nombre del equipo');
  return conLock_(() => {
    const info = salaInfo_(sala);
    if (!info) throw new Error('La sala no existe');
    if (info.estado === 'terminado') throw new Error('La sala ya terminó');
    const id = Utilities.getUuid().slice(0, 8);
    EQUIPOS().appendRow([sala, id, nombre, String(integrantes || '').slice(0, 60), new Date()]);
    CacheService.getScriptCache().remove('eq_' + sala);
    return { ok: true, equipoId: id, mision: info.mision, estado: info.estado };
  });
}

function estado_(sala) {
  const info = salaInfo_(sala);
  if (!info) throw new Error('La sala no existe');
  return { ok: true, estado: info.estado, mision: info.mision, equipos: equiposDe_(sala).map(e => e.nombre) };
}

function responder_(p) {
  const info = salaInfo_(p.sala);
  if (!info) throw new Error('La sala no existe');
  return conLock_(() => {
    RESP().appendRow([new Date(), p.sala, p.equipo, String(p.nombre || '').slice(0, 30), Number(p.item),
      String(p.caso || ''), Number(p.eleccion), Number(p.correcta) ? 1 : 0, Number(p.pista) ? 1 : 0, Number(p.seg) || 0]);
    return { ok: true };
  });
}

function resultados_(sala) {
  const info = salaInfo_(sala);
  if (!info) throw new Error('La sala no existe');
  const equipos = equiposDe_(sala);
  const porEquipo = {};
  equipos.forEach(e => porEquipo[e.id] = { nombre: e.nombre, n: 0, ok: 0, pistas: 0, items: {} });
  const porItem = {};
  RESP().getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[1]) !== String(sala)) return;
    const eq = porEquipo[String(r[2])];
    const item = Number(r[4]);
    if (eq) {
      if (eq.items[item] !== undefined) return; // solo cuenta la primera respuesta por caso
      eq.items[item] = Number(r[7]);
      eq.n++; eq.ok += Number(r[7]); eq.pistas += Number(r[8]);
    }
    const it = porItem[item] || (porItem[item] = { n: 0, ok: 0, elecciones: {} });
    it.n++; it.ok += Number(r[7]);
    it.elecciones[r[6]] = (it.elecciones[r[6]] || 0) + 1;
  });
  return { ok: true, estado: info.estado, mision: info.mision,
    equipos: Object.values(porEquipo).map(e => ({ nombre: e.nombre, n: e.n, ok: e.ok, pistas: e.pistas })),
    items: porItem };
}
