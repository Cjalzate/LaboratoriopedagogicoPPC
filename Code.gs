/**
 * EXPEDIENTE LADERA — servidor en Google Apps Script
 * Guarda el estado del juego, la bitácora y el ranking en la hoja de cálculo de la docente.
 *
 * Instalación resumida (ver la guía completa):
 *  1. Cree una hoja de cálculo en el Drive institucional y abra Extensiones > Apps Script.
 *  2. Pegue este archivo como Code.gs y cree un archivo HTML llamado «Index» con el contenido de index.html.
 *  3. Ejecute la función configurarHoja (autorice los permisos).
 *  4. Implementar > Nueva implementación > Aplicación web: «Ejecutar como: yo»;
 *     «Quién tiene acceso: cualquier persona» (o «cualquier persona de su organización»).
 */

var SH = { CONFIG: 'Config', SALAS: 'Salas', EST: 'Estudiantes', ESTADO: 'Estado', EV: 'Eventos', RANK: 'Ranking' };
var CHUNK = 40000;

/* ---------------- menú y configuración ---------------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Hemiciclo Vivo')
    .addItem('1. Configurar hoja', 'configurarHoja')
    .addItem('2. Generar códigos personales', 'generarCodigos')
    .addItem('3. Crear sala nueva', 'crearSalaMenu')
    .addToUi();
}

function configurarHoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  var defs = [
    [SH.CONFIG, ['clave', 'valor', 'nota']],
    [SH.SALAS, ['sala', 'pinMesa', 'creada', 'nota']],
    [SH.EST, ['codigo', 'id', 'nombre', 'grupo', 'nota']],
    [SH.ESTADO, ['sala', 'clave', 'json', 'actualizado']],
    [SH.EV, ['fecha', 't', 'sala', 'sid', 'mesa', 'rol', 'nodo', 'fase', 'accion', 'elemento', 'detalle', 'origen']],
    [SH.RANK, ['guardado', 'sala', 'posicion', 'codigo', 'nombre', 'seudonimo', 'mesa', 'rol', 'nodo', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'indice', 'nivel', 'revision_docente']]
  ];
  defs.forEach(function (d) {
    var sh = ss.getSheetByName(d[0]) || ss.insertSheet(d[0]);
    if (sh.getLastRow() === 0) {
      sh.appendRow(d[1]); sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, d[1].length).setFontWeight('bold');
    }
  });
  var cfg = ss.getSheetByName(SH.CONFIG);
  if (cfg.getLastRow() < 2) cfg.appendRow(['PIN_DOCENTE', String(Math.floor(100000 + Math.random() * 900000)), 'Cámbielo por un PIN que solo usted conozca.']);
  try { ss.getSheetByName(SH.ESTADO).hideSheet(); } catch (e) { }
  alert_('Hoja configurada. Su PIN docente está en la pestaña Config. Escriba los nombres de sus estudiantes en la pestaña Estudiantes y luego use el menú Hemiciclo Vivo > Generar códigos personales.');
}

function generarCodigos() {
  var sh = sheet_(SH.EST); var last = sh.getLastRow();
  if (last < 2) { alert_('Escriba primero los nombres en la pestaña Estudiantes (columna «nombre»).'); return; }
  var rng = sh.getRange(2, 1, last - 1, 3); var vals = rng.getValues(); var used = {}; var n = 0;
  vals.forEach(function (r) { if (r[0]) used[String(r[0]).toUpperCase()] = 1; });
  vals.forEach(function (r) {
    if (!r[0] && !r[2]) return;
    if (!r[0]) { var c; do { c = randCode_(6); } while (used[c]); used[c] = 1; r[0] = c; n++; }
    if (!r[1]) r[1] = 'E' + randCode_(8);
  });
  rng.setValues(vals);
  CacheService.getScriptCache().remove('roster');
  alert_('Códigos nuevos: ' + n + '. Entregue a cada estudiante su código personal. No comparta la columna «nombre».');
}

function crearSalaMenu() { var r = crearSala_(); alert_('Sala creada: ' + r.sala + '\nPIN de la Mesa Técnica: ' + r.pinMesa); }

/* ---------------- web ---------------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.q) return json_(handle_(JSON.parse(e.parameter.q)));
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Hemiciclo Vivo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doPost(e) {
  var req = {};
  try { req = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'Solicitud inválida.' }); }
  return json_(handle_(req));
}
/* llamada directa desde Index.html servido por Apps Script */
function api(req) { return handle_(req); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle_(req) {
  try {
    var a = req.a; var sala = String(req.sala || '').toUpperCase().trim();
    if (a === 'join') return join_(sala, req.code);
    if (a === 'pin') return pin_(req.role, sala, req.pin);
    var auth = auth_(req, sala); if (!auth.ok) return auth;
    if (a === 'poll') return poll_(sala, req, auth);
    if (a === 'ops') return ops_(sala, req, auth);
    if (a === 'events') { if (auth.role !== 'doc') return { ok: false, error: 'Solo la docente puede ver la bitácora.' }; return events_(sala); }
    if (a === 'guardarRanking') { if (auth.role !== 'doc') return { ok: false, error: 'Sin permiso.' }; return saveRank_(sala, req.rows || []); }
    return { ok: false, error: 'Acción desconocida.' };
  } catch (err) {
    return { ok: false, error: 'Error del servidor: ' + (err && err.message ? err.message : err) };
  }
}

/* ---------------- acceso ---------------- */
function join_(sala, code) {
  if (!salaInfo_(sala)) return { ok: false, error: 'No existe la sala ' + sala + '. Revisa el código que proyecta tu docente.' };
  var r = roster_().byCode[String(code || '').toUpperCase().trim()];
  if (!r) return { ok: false, error: 'No encontramos ese código personal. Revísalo con tu docente.' };
  return { ok: true, sid: r.id, now: Date.now() };
}
function pin_(role, sala, pin) {
  if (role === 'doc') {
    if (String(pin) !== docPin_()) return { ok: false, error: 'PIN docente incorrecto.' };
    if (!sala) { var c = crearSala_(); return { ok: true, sala: c.sala, pinMesa: c.pinMesa, created: true }; }
    var info = salaInfo_(sala); if (!info) return { ok: false, error: 'No existe la sala ' + sala + '.' };
    return { ok: true, sala: sala, pinMesa: info.pinMesa };
  }
  if (role === 'mesa') {
    var i2 = salaInfo_(sala); if (!i2) return { ok: false, error: 'No existe la sala ' + sala + '.' };
    if (String(pin) !== String(i2.pinMesa)) return { ok: false, error: 'PIN de la Mesa Técnica incorrecto.' };
    return { ok: true, sala: sala };
  }
  return { ok: false, error: 'Rol desconocido.' };
}
function auth_(req, sala) {
  var info = salaInfo_(sala); if (!info) return { ok: false, error: 'No existe la sala ' + sala + '.' };
  var role = req.role;
  if (role === 'est') {
    var r = roster_().byCode[String(req.code || '').toUpperCase().trim()];
    if (!r) return { ok: false, error: 'Código personal no válido.' };
    return { ok: true, role: 'est', sid: r.id, info: info };
  }
  if (role === 'doc') { if (String(req.pin) !== docPin_()) return { ok: false, error: 'PIN docente incorrecto.' }; return { ok: true, role: 'doc', info: info }; }
  if (role === 'mesa') { if (String(req.pin) !== String(info.pinMesa)) return { ok: false, error: 'PIN de la Mesa Técnica incorrecto.' }; return { ok: true, role: 'mesa', info: info }; }
  if (role === 'proy') return { ok: true, role: 'proy', info: info };
  return { ok: false, error: 'Rol desconocido.' };
}

/* ---------------- estado ---------------- */
function poll_(sala, req, auth) {
  var st = loadSala_(sala); var now = Date.now();
  var cache = CacheService.getScriptCache(); var pk = 'p:' + sala;
  var pres = {}; try { pres = JSON.parse(cache.get(pk) || '{}'); } catch (e) { pres = {}; }
  if (auth.role === 'est') { pres[auth.sid] = now; cache.put(pk, JSON.stringify(pres), 21600); }
  var out = { ok: true, v: st.v, pres: pres, now: now };
  if (auth.role === 'doc' || auth.role === 'proy') out.url = ScriptApp.getService().getUrl();
  if (auth.role === 'doc') out.pinMesa = auth.info.pinMesa;
  if (req.v && Number(req.v) === Number(st.v)) { out.same = true; return out; }
  out.kv = st.kv; return out;
}

function permitido_(op, auth, kv) {
  if (auth.role === 'doc') return true;
  if (auth.role === 'proy') return false;
  var p = op.p || [];
  if (auth.role === 'mesa') return op.k === 'aud' && ((p[0] === 'hilo' && op.m === 'push' && p.length === 1) || p[0] === 'marcas');
  if (auth.role === 'est') {
    if (op.k === 's:' + auth.sid) return true;
    var prof = kv['s:' + auth.sid] || {};
    if (prof.team && op.k === 't:' + prof.team) return true;
    if (op.k === 'aud' && op.m === 'push' && p.length === 1 && ['cola', 'objeciones', 'objIndice'].indexOf(p[0]) >= 0) return true;
  }
  return false;
}

function ops_(sala, req, auth) {
  var ops = req.ops || []; var evs = req.ev || [];
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var st = loadSala_(sala);
    var changed = {};
    for (var i = 0; i < ops.length; i++) {
      var op = ops[i];
      if (!permitido_(op, auth, st.kv)) return { ok: false, error: 'No tienes permiso para esa acción.' };
      if (auth.role === 'est' && op.k === 'aud' && op.v && typeof op.v === 'object') op.v.sid = auth.sid;
      applyOp_(st.kv, op); changed[op.k] = true;
    }
    var vals = {};
    for (var k in changed) {
      var js = st.kv[k] === undefined ? 'null' : JSON.stringify(st.kv[k]);
      if (js.length > 45000) return { ok: false, error: 'Este espacio tiene demasiado contenido. Pide a tu docente que revise la mesa.' };
      vals[k] = st.kv[k];
    }
    if (Object.keys(vals).length) writeKeys_(sala, st, vals);
    if (evs.length) appendEvents_(sala, auth, st.kv, evs);
    return { ok: true, v: st.v, now: Date.now() };
  } finally { lock.releaseLock(); }
}

function applyOp_(kv, op) {
  var root = kv[op.k];
  if (!op.p || !op.p.length) {
    if (op.m === 'del') { delete kv[op.k]; return; }
    if (op.m === 'push') { if (!Array.isArray(root)) root = []; root.push(op.v); kv[op.k] = root; return; }
    kv[op.k] = op.v; return;
  }
  if (root === undefined || root === null || typeof root !== 'object') { root = {}; kv[op.k] = root; }
  var o = root;
  for (var i = 0; i < op.p.length - 1; i++) {
    var key = op.p[i];
    if (o[key] === undefined || o[key] === null || typeof o[key] !== 'object') o[key] = {};
    o = o[key];
  }
  var last = op.p[op.p.length - 1];
  if (op.m === 'del') delete o[last];
  else if (op.m === 'push') { if (!Array.isArray(o[last])) o[last] = []; o[last].push(op.v); }
  else o[last] = op.v;
}

function loadSala_(sala) {
  var cached = cacheGetObj_('st:' + sala);
  if (cached) return cached;
  var sh = sheet_(SH.ESTADO); var vals = sh.getDataRange().getValues();
  var kv = {}, rows = {};
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] !== sala) continue;
    rows[vals[i][1]] = i + 1;
    if (vals[i][2] && vals[i][2] !== 'null') { try { kv[vals[i][1]] = JSON.parse(vals[i][2]); } catch (e) { } }
  }
  var v = Number(PropertiesService.getScriptProperties().getProperty('v:' + sala) || 1);
  var st = { kv: kv, rows: rows, v: v };
  cachePutObj_('st:' + sala, st);
  return st;
}

function writeKeys_(sala, st, vals) {
  var sh = sheet_(SH.ESTADO); var now = new Date();
  for (var k in vals) {
    var js = vals[k] === undefined ? 'null' : JSON.stringify(vals[k]);
    if (st.rows[k]) sh.getRange(st.rows[k], 3, 1, 2).setValues([[js, now]]);
    else { sh.appendRow([sala, k, js, now]); st.rows[k] = sh.getLastRow(); }
  }
  st.v = Date.now();
  PropertiesService.getScriptProperties().setProperty('v:' + sala, String(st.v));
  cachePutObj_('st:' + sala, st);
}

/* ---------------- bitácora ---------------- */
var BLOQUEADOS_EST = ['docente_marca', 'mesa_bloque', 'mesa_respuesta', 'mesa_pregunta', 'meta_cumplida', 'dictamen_validado'];
function appendEvents_(sala, auth, kv, evs) {
  var aud = kv.aud || {}; var fase = (aud.etapa || 'espera') === 'espera' ? 'estructuracion' : 'transferencia';
  var prof = auth.role === 'est' ? (kv['s:' + auth.sid] || {}) : {};
  var t = Date.now(); var rows = [];
  evs.forEach(function (e) {
    if (auth.role === 'est' && BLOQUEADOS_EST.indexOf(e.tipo) >= 0) return;
    var origen = auth.role === 'est' ? (e.tipo === 'coev' ? 'pares' : 'auto') : (auth.role === 'mesa' ? 'mesa' : 'docente');
    var det = JSON.stringify(e.det || {}); if (det.length > 2000) det = det.slice(0, 2000);
    rows.push([new Date(t), t, sala, auth.role === 'est' ? auth.sid : '', prof.team || '', prof.rol || '', prof.nodo || '', fase, String(e.tipo || '').slice(0, 40), String(e.el || '').slice(0, 80), det, origen]);
  });
  if (!rows.length) return;
  var sh = sheet_(SH.EV);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}
function events_(sala) {
  var vals = sheet_(SH.EV).getDataRange().getValues(); var out = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i]; if (r[2] !== sala) continue;
    var det = {}; try { det = JSON.parse(r[10] || '{}'); } catch (e) { }
    out.push({ t: Number(r[1]), sid: r[3], team: r[4], rol: r[5], nodo: r[6], fase: r[7], tipo: r[8], el: r[9], det: det, origen: r[11] });
  }
  return { ok: true, rows: out };
}
function saveRank_(sala, rows) {
  var byId = roster_().byId; var now = new Date();
  var data = rows.map(function (r) {
    var ros = byId[r.sid] || {};
    return [now, sala, r.posicion, ros.codigo || '', ros.nombre || '', r.seudonimo, r.mesa, r.rol, r.nodo, r.C1, r.C2, r.C3, r.C4, r.C5, r.C6, r.indice, r.nivel, r.revision_docente];
  });
  if (!data.length) return { ok: false, error: 'No hay filas para guardar.' };
  var sh = sheet_(SH.RANK);
  sh.getRange(sh.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
  return { ok: true, n: data.length };
}

/* ---------------- utilidades ---------------- */
function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var a = SpreadsheetApp.getActiveSpreadsheet();
  if (a) return a;
  throw new Error('Ejecute configurarHoja desde la hoja de cálculo.');
}
function sheet_(name) { var sh = ss_().getSheetByName(name); if (!sh) throw new Error('Falta la pestaña ' + name + '. Ejecute configurarHoja.'); return sh; }
function alert_(m) { try { SpreadsheetApp.getUi().alert(m); } catch (e) { Logger.log(m); } }
function randCode_(n) { var a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; var s = ''; for (var i = 0; i < n; i++) s += a.charAt(Math.floor(Math.random() * a.length)); return s; }
function docPin_() {
  var c = CacheService.getScriptCache(); var v = c.get('docpin'); if (v) return v;
  var vals = sheet_(SH.CONFIG).getDataRange().getValues(); var pin = '';
  for (var i = 1; i < vals.length; i++) if (vals[i][0] === 'PIN_DOCENTE') pin = String(vals[i][1]).trim();
  if (!pin) throw new Error('Falta el PIN_DOCENTE en la pestaña Config.');
  c.put('docpin', pin, 300); return pin;
}
function salaInfo_(sala) {
  if (!sala) return null;
  var c = CacheService.getScriptCache(); var k = 'sala:' + sala; var v = c.get(k);
  if (v) return JSON.parse(v);
  var vals = sheet_(SH.SALAS).getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) if (String(vals[i][0]).toUpperCase() === sala) { var info = { pinMesa: String(vals[i][1]) }; c.put(k, JSON.stringify(info), 600); return info; }
  return null;
}
function roster_() {
  var c = CacheService.getScriptCache(); var v = c.get('roster');
  if (v) return JSON.parse(v);
  var vals = sheet_(SH.EST).getDataRange().getValues(); var byCode = {}, byId = {};
  for (var i = 1; i < vals.length; i++) {
    var code = String(vals[i][0] || '').toUpperCase().trim(); var id = String(vals[i][1] || '').trim();
    if (!code || !id) continue;
    byCode[code] = { id: id }; byId[id] = { codigo: code, nombre: String(vals[i][2] || '') };
  }
  var r = { byCode: byCode, byId: byId };
  c.put('roster', JSON.stringify(r), 300);
  return r;
}
function crearSala_() {
  var sh = sheet_(SH.SALAS); var vals = sh.getDataRange().getValues(); var used = {};
  vals.forEach(function (r) { used[String(r[0])] = 1; });
  var sala; do { sala = randCode_(6); } while (used[sala]);
  var pinMesa = String(Math.floor(1000 + Math.random() * 9000));
  sh.appendRow([sala, pinMesa, new Date(), '']);
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try { var st = loadSala_(sala); writeKeys_(sala, st, { cfg: defaultCfg_(), aud: defaultAud_() }); } finally { lock.releaseLock(); }
  return { sala: sala, pinMesa: pinMesa };
}
function defaultCfg_() {
  return { tAlegato: 180, tPreguntas: 120, tReplicas: 60, tGeneral: 90, evid: 8, andamiaje: 'alto', mesas: 7, reloj: true, linea: true, mercyGen: 'f', rangos: { S: '', A: '', B: '' }, coev: true };
}
function defaultAud_() {
  return { etapa: 'espera', turno: { team: '', sub: '', inicio: 0, dur: 0, pausa: 0 }, generalInicio: 0, hecho: [], cola: [], hilo: [], marcas: {}, objeciones: [], objIndice: [], palabra: null, reconoc: [] };
}
function cachePutObj_(key, obj) {
  var s = JSON.stringify(obj); var c = CacheService.getScriptCache();
  var n = Math.max(1, Math.ceil(s.length / CHUNK)); var m = {}; m[key + ':n'] = String(n);
  for (var i = 0; i < n; i++) m[key + ':' + i] = s.substr(i * CHUNK, CHUNK);
  try { c.putAll(m, 21600); } catch (e) { c.remove(key + ':n'); }
}
function cacheGetObj_(key) {
  var c = CacheService.getScriptCache(); var n = c.get(key + ':n'); if (!n) return null;
  n = Number(n); var keys = []; for (var i = 0; i < n; i++) keys.push(key + ':' + i);
  var got = c.getAll(keys); var s = '';
  for (var j = 0; j < n; j++) { var part = got[key + ':' + j]; if (part === undefined || part === null) return null; s += part; }
  try { return JSON.parse(s); } catch (e) { return null; }
}
