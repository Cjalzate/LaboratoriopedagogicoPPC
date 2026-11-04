/**
 * HEMICICLO VIVO — servidor en Google Apps Script (versión 4)
 * Una sala aislada por grupo (11°1 a 11°4). Guarda el estado, la bitácora y el seguimiento en la hoja de la docente.
 * Los PIN se guardan cifrados (SHA-256 con sal) en las propiedades del script: no aparecen en la hoja.
 *
 * Instalación resumida (ver la guía completa):
 *  1. Cree una hoja de cálculo en el Drive institucional y abra Extensiones > Apps Script.
 *  2. Pegue este archivo como Code.gs y cree un archivo HTML llamado «Index» con el contenido de index.html.
 *  3. Ejecute la función configurarHoja (autorice los permisos). Anote el PIN docente y los PIN de Mesa que aparecen UNA sola vez.
 *  3b. Pegue el listado en la pestaña Estudiantes (columnas nombre, nombre_visible, grupo) y use el menú «Generar códigos personales».
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
    .addItem('3. Restablecer PIN docente (olvido)', 'restablecerPinDocente')
    .addToUi();
}

var GRUPOS_SRV = [
  { sala: '1101', n: '11°1', codigo: '110100' },
  { sala: '1102', n: '11°2', codigo: '110200' },
  { sala: '1103', n: '11°3', codigo: '110300' },
  { sala: '1104', n: '11°4', codigo: '110400' }
];
function salaDeGrupo_(g) {
  g = String(g || '').replace(/\s/g, '').toUpperCase().replace('º', '°');
  for (var i = 0; i < GRUPOS_SRV.length; i++) { var x = GRUPOS_SRV[i]; if (g === x.sala || g === x.codigo || g === x.n.toUpperCase() || g === x.n.replace('°', '-').toUpperCase() || g === x.n.replace('°', '').toUpperCase()) return x.sala; }
  return '';
}
function grupoN_(sala) { for (var i = 0; i < GRUPOS_SRV.length; i++) if (GRUPOS_SRV[i].sala === sala) return GRUPOS_SRV[i].n; return sala; }

function configurarHoja() {
  var r = setup_();
  alert_('Hoja configurada.\n\nPIN DOCENTE: ' + r.pin + '\nPIN de la Mesa Técnica: ' + r.mesa.map(function (m) { return m.n + ' → ' + m.pin; }).join(' · ') +
    '\n\nAnótelos ahora: se guardan cifrados y no se vuelven a mostrar. Puede cambiarlos en el panel docente (pestaña «Contraseñas y códigos»).\n\nLuego pegue el listado en la pestaña Estudiantes y use el menú Hemiciclo Vivo > Generar códigos personales.');
}
function setup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pr = PropertiesService.getScriptProperties();
  pr.setProperty('SHEET_ID', ss.getId());
  var defs = [
    [SH.CONFIG, ['clave', 'valor', 'nota']],
    [SH.SALAS, ['sala', 'grupo', 'creada', 'nota']],
    [SH.EST, ['codigo', 'id', 'nombre', 'nombre_visible', 'grupo', 'nota']],
    [SH.ESTADO, ['sala', 'clave', 'json', 'actualizado']],
    [SH.EV, ['fecha', 't', 'sala', 'sid', 'mesa', 'rol', 'nodo', 'fase', 'accion', 'elemento', 'detalle', 'origen']],
    [SH.RANK, ['guardado', 'grupo', 'posicion', 'codigo', 'nombre', 'nombre_visible', 'mesa', 'rol', 'nodo', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'indice', 'nivel', 'revision_docente']]
  ];
  defs.forEach(function (d) {
    var sh = ss.getSheetByName(d[0]) || ss.insertSheet(d[0]);
    if (sh.getLastRow() === 0) {
      sh.appendRow(d[1]); sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, d[1].length).setFontWeight('bold');
    }
  });
  var out = { pin: '', mesa: [] };
  var cfg = ss.getSheetByName(SH.CONFIG); var cv = cfg.getDataRange().getValues(); var viejo = '', fila = -1;
  for (var k = 1; k < cv.length; k++) if (cv[k][0] === 'PIN_DOCENTE') { fila = k + 1; viejo = String(cv[k][1]).trim(); }
  if (!pr.getProperty('PIN_DOC_HASH')) {
    if (/^\d{4,8}$/.test(viejo)) { pr.setProperty('PIN_DOC_HASH', hashPin_('doc', viejo)); out.pin = '(se mantiene su PIN anterior, ahora cifrado)'; }
    else { out.pin = String(Math.floor(100000 + Math.random() * 900000)); pr.setProperty('PIN_DOC_HASH', hashPin_('doc', out.pin)); }
  } else out.pin = '(ya configurado: no cambia)';
  var nota = 'El PIN se guarda cifrado en las propiedades del script. Cámbielo desde el panel docente.';
  if (fila > 0) cfg.getRange(fila, 2, 1, 2).setValues([['cifrado', nota]]); else cfg.appendRow(['PIN_DOCENTE', 'cifrado', nota]);
  var salas = ss.getSheetByName(SH.SALAS); var vals = salas.getDataRange().getValues(); var has = {};
  for (var i = 1; i < vals.length; i++) has[String(vals[i][0])] = 1;
  GRUPOS_SRV.forEach(function (g) {
    if (!has[g.sala]) { salas.appendRow([g.sala, g.n, new Date(), 'Sala del grupo ' + g.n]); }
    if (!pr.getProperty('PIN_MESA_' + g.sala)) { var mp = String(Math.floor(1000 + Math.random() * 9000)); pr.setProperty('PIN_MESA_' + g.sala, hashPin_('mesa', mp)); out.mesa.push({ n: g.n, pin: mp }); }
    var lock = LockService.getScriptLock(); lock.waitLock(25000);
    try { var st = loadSala_(g.sala); if (!st.kv.cfg) writeKeys_(g.sala, st, { cfg: defaultCfg_(), aud: defaultAud_() }); } finally { lock.releaseLock(); }
  });
  try { ss.getSheetByName(SH.ESTADO).hideSheet(); } catch (e) { }
  return out;
}
function restablecerPinDocente() {
  var pin = String(Math.floor(100000 + Math.random() * 900000));
  PropertiesService.getScriptProperties().setProperty('PIN_DOC_HASH', hashPin_('doc', pin));
  alert_('Nuevo PIN docente: ' + pin + '\nAnótelo: no se vuelve a mostrar. Solo quien edita esta hoja puede usar esta opción.');
}

function generarCodigos() {
  var sh = sheet_(SH.EST); var last = sh.getLastRow();
  if (last < 2) { alert_('Pegue primero el listado en la pestaña Estudiantes (columnas nombre, nombre_visible y grupo).'); return; }
  var rng = sh.getRange(2, 1, last - 1, 5); var vals = rng.getValues(); var used = {}; var n = 0, sinGrupo = 0;
  vals.forEach(function (r) { if (r[0]) used[String(r[0]).toUpperCase()] = 1; });
  vals.forEach(function (r) {
    if (!r[0] && !r[2]) return;
    if (!r[0]) { var c; do { c = randCode_(6); } while (used[c]); used[c] = 1; r[0] = c; n++; }
    if (!r[1]) r[1] = 'E' + randCode_(8);
    if (!r[3] && r[2]) r[3] = nombreVisible_(r[2]);
    if (!salaDeGrupo_(r[4])) sinGrupo++;
  });
  rng.setValues(vals);
  CacheService.getScriptCache().remove('roster');
  alert_('Códigos nuevos: ' + n + '.' + (sinGrupo ? ' Atención: ' + sinGrupo + ' fila(s) sin grupo válido (use 11°1, 11°2, 11°3 u 11°4).' : '') + ' Entregue a cada estudiante su código en mano (panel docente > Contraseñas y códigos > Tarjetas de acceso).');
}
/* «ACEVEDO OSORIO SAMUEL» → «Samuel Acevedo» (orden del listado: apellidos y luego nombres) */
function nombreVisible_(full) {
  var t = String(full || '').trim().split(/\s+/); if (!t.length) return '';
  var cap = function (w) { w = w.toLowerCase(); return w.charAt(0).toUpperCase() + w.slice(1); };
  if (t.length === 1) return cap(t[0]);
  if (t.length === 2) return cap(t[1]) + ' ' + cap(t[0]);
  return cap(t[2]) + ' ' + cap(t[0]);
}


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
    if (a === 'join') return join_(req.code);
    if (a === 'pin') return pin_(req.role, sala, req.pin);
    var auth = auth_(req, sala); if (!auth.ok) return auth;
    var adm = ['roster', 'setCode', 'setDocPin', 'setMesaPin', 'resumen'];
    if (adm.indexOf(a) >= 0) { if (auth.role !== 'doc') return { ok: false, error: 'Solo la docente puede hacer esto.' }; return admin_(a, sala, req); }
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
function hashPin_(tipo, pin) {
  var pr = PropertiesService.getScriptProperties(); var salt = pr.getProperty('PIN_SALT');
  if (!salt) { salt = Utilities.getUuid(); pr.setProperty('PIN_SALT', salt); }
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + tipo + '|' + String(pin), Utilities.Charset.UTF_8);
  return d.map(function (b) { var h = (b < 0 ? b + 256 : b).toString(16); return h.length === 1 ? '0' + h : h; }).join('');
}
function bloqueado_(k, max) { var v = Number(CacheService.getScriptCache().get('fail:' + k) || 0); return v >= max; }
function fallo_(k) { var c = CacheService.getScriptCache(); var v = Number(c.get('fail:' + k) || 0) + 1; c.put('fail:' + k, String(v), 300); }
function docOk_(pin) { var h = PropertiesService.getScriptProperties().getProperty('PIN_DOC_HASH'); return !!h && !!pin && hashPin_('doc', pin) === h; }
function mesaOk_(sala, pin) { var h = PropertiesService.getScriptProperties().getProperty('PIN_MESA_' + sala); return !!h && !!pin && hashPin_('mesa', pin) === h; }

function join_(code) {
  if (bloqueado_('join', 40)) return { ok: false, error: 'Demasiados intentos fallidos. Espera unos minutos.' };
  var r = roster_().byCode[String(code || '').toUpperCase().trim()];
  if (!r) { fallo_('join'); return { ok: false, error: 'No encontramos ese código personal. Revísalo con tu docente.' }; }
  if (!r.sala || !salaInfo_(r.sala)) return { ok: false, error: 'Tu código no tiene un grupo asignado. Avísale a tu docente.' };
  return { ok: true, sid: r.id, sala: r.sala, nombre: r.vis, grupo: grupoN_(r.sala), now: Date.now() };
}
function pin_(role, sala, pin) {
  if (bloqueado_('pin:' + role, 10)) return { ok: false, error: 'Demasiados intentos fallidos. Espere 5 minutos.' };
  if (role === 'doc') {
    if (!docOk_(pin)) { fallo_('pin:doc'); return { ok: false, error: 'PIN docente incorrecto.' }; }
    return { ok: true, salas: GRUPOS_SRV };
  }
  if (role === 'mesa') {
    if (!salaInfo_(sala)) return { ok: false, error: 'Elige el grupo.' };
    if (!mesaOk_(sala, pin)) { fallo_('pin:mesa'); return { ok: false, error: 'PIN de la Mesa Técnica incorrecto.' }; }
    return { ok: true, sala: sala };
  }
  return { ok: false, error: 'Rol desconocido.' };
}
function auth_(req, sala) {
  var info = salaInfo_(sala); if (!info) return { ok: false, error: 'No existe el grupo ' + sala + '.' };
  var role = req.role;
  if (info.lib && role !== 'doc') return { ok: false, error: 'Solo la docente puede abrir la biblioteca de casos.' };
  if (role === 'est') {
    var r = roster_().byCode[String(req.code || '').toUpperCase().trim()];
    if (!r) return { ok: false, error: 'Código personal no válido. Vuelve a entrar.' };
    if (r.sala !== sala) return { ok: false, error: 'Tu código pertenece a otro grupo.' };
    return { ok: true, role: 'est', sid: r.id, vis: r.vis, info: info };
  }
  if (role === 'doc') { if (!docOk_(req.pin)) return { ok: false, error: 'Sesión docente no válida. Vuelva a entrar.' }; return { ok: true, role: 'doc', info: info }; }
  if (role === 'mesa') { if (!mesaOk_(sala, req.pin)) return { ok: false, error: 'Sesión de la Mesa no válida. Vuelva a entrar.' }; return { ok: true, role: 'mesa', info: info }; }
  if (role === 'proy') return { ok: true, role: 'proy', info: info };
  return { ok: false, error: 'Rol desconocido.' };
}

/* ---------------- administración (solo docente) ---------------- */
function admin_(a, sala, req) {
  if (a === 'roster') {
    var byId = roster_().byId; var rows = [];
    for (var id in byId) if (byId[id].sala === sala) rows.push({ id: id, nombre: byId[id].nombre, vis: byId[id].vis, codigo: byId[id].codigo });
    rows.sort(function (x, y) { return x.nombre.localeCompare(y.nombre, 'es'); });
    return { ok: true, rows: rows };
  }
  if (a === 'setCode') {
    var sh = sheet_(SH.EST); var vals = sh.getDataRange().getValues(); var nc = String(req.nuevo || '').trim().toUpperCase(); var row = -1; var used = {};
    for (var i = 1; i < vals.length; i++) { used[String(vals[i][0]).toUpperCase()] = i; if (String(vals[i][1]) === String(req.id)) row = i; }
    if (row < 0 || salaDeGrupo_(vals[row][4]) !== sala) return { ok: false, error: 'Estudiante no encontrado en este grupo.' };
    if (!nc) { do { nc = randCode_(6); } while (used[nc] !== undefined); }
    if (!/^[A-Z0-9-]{4,12}$/.test(nc)) return { ok: false, error: 'El código debe tener de 4 a 12 letras, números o guiones.' };
    if (used[nc] !== undefined && used[nc] !== row) return { ok: false, error: 'Ese código ya lo tiene otra persona.' };
    sh.getRange(row + 1, 1).setValues([[nc]]);
    CacheService.getScriptCache().remove('roster');
    return { ok: true, codigo: nc };
  }
  if (a === 'setDocPin') {
    var np = String(req.nuevo || ''); if (!/^\d{4,8}$/.test(np)) return { ok: false, error: 'El PIN debe tener de 4 a 8 números.' };
    PropertiesService.getScriptProperties().setProperty('PIN_DOC_HASH', hashPin_('doc', np)); return { ok: true };
  }
  if (a === 'setMesaPin') {
    var g = String(req.grupo || sala); var mp = String(req.nuevo || '');
    if (!salaInfo_(g)) return { ok: false, error: 'Grupo no válido.' };
    if (!/^\d{4,8}$/.test(mp)) return { ok: false, error: 'El PIN debe tener de 4 a 8 números.' };
    PropertiesService.getScriptProperties().setProperty('PIN_MESA_' + g, hashPin_('mesa', mp)); return { ok: true };
  }
  if (a === 'resumen') { var out = {}; GRUPOS_SRV.forEach(function (x) { out[x.sala] = loadSala_(x.sala).kv; }); return { ok: true, salas: out }; }
  return { ok: false, error: 'Acción desconocida.' };
}

/* ---------------- estado ---------------- */
function poll_(sala, req, auth) {
  var st = loadSala_(sala); var now = Date.now();
  var cache = CacheService.getScriptCache(); var pk = 'p:' + sala;
  var pres = {}; try { pres = JSON.parse(cache.get(pk) || '{}'); } catch (e) { pres = {}; }
  if (auth.role === 'est') { pres[auth.sid] = now; cache.put(pk, JSON.stringify(pres), 21600); }
  var out = { ok: true, v: st.v, pres: pres, now: now };
  if (auth.role === 'doc' || auth.role === 'proy') out.url = ScriptApp.getService().getUrl();
  if (req.v && Number(req.v) === Number(st.v)) { out.same = true; return out; }
  var meta = st.kv['caso:meta'];
  if (req.ch && meta && meta.hash === req.ch) { var kv2 = {}; for (var k in st.kv) if (!/^caso:\d+$/.test(k)) kv2[k] = st.kv[k]; out.kv = kv2; return out; }
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
      if (auth.role === 'est' && op.k === 's:' + auth.sid) {
        var pp = op.p || [];
        if (pp[0] === 'seud') return { ok: false, error: 'Tu nombre lo asigna la docente.' };
        if (!pp.length && op.v && typeof op.v === 'object') { op.v.seud = auth.vis; op.v.sid = auth.sid; }
      }
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
    return [now, grupoN_(sala), r.posicion, ros.codigo || '', ros.nombre || '', ros.vis || r.seudonimo, r.mesa, r.rol, r.nodo, r.C1, r.C2, r.C3, r.C4, r.C5, r.C6, r.indice, r.nivel, r.revision_docente];
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
  if (sala === 'LIB') return { grupo: 'Biblioteca de casos', lib: true };
  for (var i = 0; i < GRUPOS_SRV.length; i++) if (GRUPOS_SRV[i].sala === sala) return { grupo: GRUPOS_SRV[i].n };
  return null;
}
function roster_() {
  var c = CacheService.getScriptCache(); var v = c.get('roster');
  if (v) return JSON.parse(v);
  var vals = sheet_(SH.EST).getDataRange().getValues(); var byCode = {}, byId = {};
  for (var i = 1; i < vals.length; i++) {
    var code = String(vals[i][0] || '').toUpperCase().trim(); var id = String(vals[i][1] || '').trim();
    if (!code || !id) continue;
    var sala = salaDeGrupo_(vals[i][4]); var vis = String(vals[i][3] || '') || nombreVisible_(vals[i][2]);
    byCode[code] = { id: id, sala: sala, vis: vis }; byId[id] = { codigo: code, nombre: String(vals[i][2] || ''), vis: vis, sala: sala };
  }
  var r = { byCode: byCode, byId: byId };
  c.put('roster', JSON.stringify(r), 300);
  return r;
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
