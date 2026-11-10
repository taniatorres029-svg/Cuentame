/**
 * Cuéntame cómo estás · base privada de respuestas
 * Dra. Tania Torres · Oncología médica
 *
 * Va dentro de una Hoja de cálculo de Google (Extensiones → Apps Script).
 * - doPost: la app de pacientes manda aquí cada cuestionario y se agrega un renglón.
 * - doGet: el panel de la doctora lee las respuestas, solo si trae la CLAVE correcta.
 */

// Cambia esta clave por una tuya (la usarás para abrir tu panel). No la compartas.
const CLAVE = 'cambia-esta-clave';

const HOJA = 'Respuestas';
const COLS = ['Recibido', 'Fecha', 'Nombre', 'Celular', 'Última quimio', 'Día post-quimio',
  'Cansancio', 'Náusea', 'Vómito', 'Diarrea', 'Estreñimiento', 'Dolor', 'Boca', 'Apetito',
  'Fiebre', 'Nota', 'Alarma'];
const SINTOMAS = COLS.slice(6, 14);

function hoja_() {
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(HOJA);
  if (!h) {
    h = ss.insertSheet(HOJA);
    h.appendRow(COLS);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, COLS.length).setFontWeight('bold');
    h.getRange('B:B').setNumberFormat('@');
    h.getRange('E:E').setNumberFormat('@');
  }
  return h;
}

// Evita que un texto escrito por el paciente se interprete como fórmula.
function texto_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max || 200);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
function grado_(v) {
  const n = Number(v);
  return (n === 0 || n === 1 || n === 2 || n === 3) ? n : '';
}
function fechaISO_(v) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? String(v) : '';
}

// Hoja con datos que pone la doctora por paciente (por ahora, el tratamiento).
const HOJA_PAC = 'Pacientes';
function hojaPac_() {
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(HOJA_PAC);
  if (!h) {
    h = ss.insertSheet(HOJA_PAC);
    h.appendRow(['Paciente', 'Nombre', 'Tratamiento', 'Actualizado']);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, 4).setFontWeight('bold');
    h.getRange('A:A').setNumberFormat('@');
  }
  return h;
}
function guardarTratamiento_(d) {
  if (d.clave !== CLAVE || CLAVE === 'cambia-esta-clave') return salida_({ ok: false, error: 'clave' });
  const id = texto_(d.id, 80);
  if (!id) return salida_({ ok: false });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const h = hojaPac_();
    const ids = h.getRange(1, 1, h.getLastRow(), 1).getValues().map(r => String(r[0]));
    const fila = [id, texto_(d.nombre, 80), texto_(d.tratamiento, 300), new Date()];
    const i = ids.indexOf(id);
    if (i > 0) h.getRange(i + 1, 1, 1, 4).setValues([fila]); else h.appendRow(fila);
  } finally { lock.releaseLock(); }
  return salida_({ ok: true });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return salida_({ ok: false }); }
  if (d.tipo === 'tratamiento') return guardarTratamiento_(d);
  const s = d.s || {};
  const fila = [
    new Date(),
    fechaISO_(d.fecha),
    texto_(d.nombre, 80),
    texto_(String(d.celular || '').replace(/\D/g, ''), 15),
    fechaISO_(d.ultimaQuimio),
    (d.dia === '' || d.dia == null || isNaN(Number(d.dia))) ? '' : Number(d.dia)
  ].concat(SINTOMAS.map(k => grado_(s[k])), [
    texto_(d.fiebre, 10),
    texto_(d.nota, 1000),
    d.alarma ? 'SÍ' : ''
  ]);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { hoja_().appendRow(fila); } finally { lock.releaseLock(); }
  return salida_({ ok: true });
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.clave !== CLAVE || CLAVE === 'cambia-esta-clave') {
    return salida_({ ok: false, error: 'clave' });
  }
  const tz = Session.getScriptTimeZone();
  const valores = hoja_().getDataRange().getValues();
  const filas = valores.slice(1).map(r => r.map((v, i) => {
    if (v instanceof Date) return Utilities.formatDate(v, tz, i === 0 ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
    return v;
  }));
  const tratamientos = {};
  hojaPac_().getDataRange().getValues().slice(1).forEach(r => { if (r[0]) tratamientos[String(r[0])] = String(r[2] || ''); });
  return salida_({ ok: true, cols: valores[0], filas: filas, tratamientos: tratamientos });
}

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
