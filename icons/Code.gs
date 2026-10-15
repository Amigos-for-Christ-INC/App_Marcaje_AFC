/**
 * MARCAJE DE ENTRADA Y SALIDA — Google Apps Script  (v2)
 * Creada y de autoría del Ing. Moisés Mendoza
 *
 * - Ingreso con correo y contraseña (guardada cifrada, nunca en texto).
 * - Cada empleado crea, cambia o recupera su contraseña con un código por correo.
 * - Solo se puede marcar dentro del lugar de trabajo asignado (geocerca).
 * - Jornada corrida: horas = salida - entrada (sin descontar almuerzo).
 * Hojas: Empleados, Lugares, Marcas, Permisos, Config (se crean solas).
 */

const H = { EMP: 'Empleados', LUG: 'Lugares', MAR: 'Marcas', PER: 'Permisos', CFG: 'Config' };
const EMP_COLS = ['Correo', 'Nombre', 'Cargo', 'Lugar', 'Activo', 'Hash', 'Sal', 'DebeCambiar', 'Rol'];
const MAR_COLS = ['ID', 'Fecha', 'Correo', 'Nombre', 'Entrada', 'Salida', 'Horas', 'LatEntrada', 'LngEntrada', 'MetrosEntrada', 'LatSalida', 'LngSalida', 'MetrosSalida', 'Lugar', 'Observacion', 'IdsCliente'];
const PER_COLS = ['ID', 'Solicitado', 'Correo', 'Nombre', 'Tipo', 'Desde', 'Hasta', 'HoraDesde', 'HoraHasta', 'Motivo', 'Estado', 'Respuesta', 'Revisado', 'RevisadoPor'];
// Roles: Empleado (solo marca), Supervisor (asistencia y permisos de su lugar),
// RRHH (todo el personal, crea empleados), Admin (todo, incluidos roles, lugares y PIN)
const ROLES = ['Empleado', 'Supervisor', 'RRHH', 'Admin'];
const TIPOS_PERMISO = ['Permiso personal', 'Cita médica', 'Vacaciones', 'Llegada tarde', 'Salida temprana', 'Otro'];
const LIBRE = 'LIBRE';             // Lugar = "Libre": puede marcar desde cualquier sitio (se guarda el GPS igual)

/* ================= Web app ================= */
function doGet() {
  setup_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Control de asistencia | Amigos for Christ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================= API para la app instalable (Cloudflare Pages) =================
 * La app publicada fuera de Google llama aquí con POST {fn, args}.
 * Solo se pueden llamar las funciones de esta lista; todo lo demás se rechaza.
 */
function apiPublica_() {
  return {
    obtenerEmpresa, login, sesion, cerrarSesion, pedirCodigo, restablecerClave, cambiarClave,
    marcar, sincronizarMarcas, misMarcas, solicitarPermiso, misPermisos,
    adminLogin, adminResumen, responderPermiso, guardarEmpleado, guardarLugar, cambiarPinAdmin,
    reporte, reporteADrive, estadoRespaldo, respaldarAhora, activarRespaldoMensual,
    guardarAvisos, probarAviso
  };
}

function doPost(e) {
  let res;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = String(req.fn || '');
    const api = apiPublica_();
    if (!Object.prototype.hasOwnProperty.call(api, fn)) throw new Error('Función no permitida.');
    const args = Array.isArray(req.args) ? req.args.slice(0, 10) : [];
    setup_();
    res = { ok: true, data: api[fn].apply(null, args) };
  } catch (err) {
    res = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= Utilidades ================= */
function ss_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('Marcaje - Base de datos');
  p.setProperty('SS_ID', ss.getId());
  return ss;
}
function sh_(n) { return ss_().getSheetByName(n); }
function cache_() { return CacheService.getScriptCache(); }
function tz_() { return Session.getScriptTimeZone(); }
function hoy_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function hora_() { return Utilities.formatDate(new Date(), tz_(), 'HH:mm:ss'); }
function ahora_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'); }
function id_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase(); }
function correo_(c) { return String(c || '').trim().toLowerCase(); }

function setup_() {
  const ss = ss_();
  ensure_(ss, H.EMP, EMP_COLS);
  ensure_(ss, H.LUG, ['Nombre', 'Lat', 'Lng', 'RadioMetros']);
  ensure_(ss, H.MAR, MAR_COLS);
  ensure_(ss, H.PER, PER_COLS);
  migrar_(H.EMP, EMP_COLS);
  migrar_(H.PER, PER_COLS);
  migrar_(H.MAR, MAR_COLS);
  const cfg = ensure_(ss, H.CFG, ['Clave', 'Valor']);
  if (cfg.getLastRow() < 2) {
    writeText_(cfg, 2, [['PIN_ADMIN', '1234'], ['EMPRESA', 'Mi Empresa'], ['PRECISION_MAX_METROS', '100']]);
  }
}

function ensure_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Agrega encabezados nuevos a hojas creadas con versiones anteriores
function migrar_(name, cols) {
  const sh = sh_(name);
  if (sh.getMaxColumns() < cols.length) sh.insertColumnsAfter(sh.getMaxColumns(), cols.length - sh.getMaxColumns());
  const h = sh.getRange(1, 1, 1, cols.length).getDisplayValues()[0];
  cols.forEach((c, i) => { if (!h[i]) sh.getRange(1, i + 1).setValue(c).setFontWeight('bold'); });
}

// Escribe como texto para que Sheets no convierta fechas ni horas
function writeText_(sh, row, values, col) {
  const rg = sh.getRange(row, col || 1, values.length, values[0].length);
  rg.setNumberFormat('@');
  rg.setValues(values.map(r => r.map(v => v === null || v === undefined ? '' : String(v))));
}

function rows_(name) {
  const v = sh_(name).getDataRange().getDisplayValues();
  v.shift();
  return v;
}

function cfg_(k) {
  const r = rows_(H.CFG).find(x => x[0] === k);
  return r ? String(r[1]).trim() : '';
}

function lock_(fn) {
  const l = LockService.getScriptLock();
  if (!l.tryLock(30000)) throw new Error('Hay mucha gente marcando en este momento. Intenta de nuevo en unos segundos.');
  try { return fn(); } finally { l.releaseLock(); }
}

function difHoras_(a, b) {
  const s = t => { const p = String(t).split(':').map(Number); return p[0] * 3600 + p[1] * 60 + (p[2] || 0); };
  let d = s(b) - s(a);
  if (d < 0) d += 86400;
  return Math.round(d / 36) / 100;
}

function dist_(lat1, lng1, lat2, lng2) {
  const R = 6371000, r = x => x * Math.PI / 180;
  const a = Math.sin(r(lat2 - lat1) / 2) ** 2 + Math.cos(r(lat1)) * Math.cos(r(lat2)) * Math.sin(r(lng2 - lng1) / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

/* ================= Contraseñas ================= */
function hash_(clave, sal) {
  let h = sal + '|' + clave;
  for (let i = 0; i < 300; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8));
  }
  return h;
}
function validarClave_(c) {
  c = String(c || '');
  if (c.length < 6) throw new Error('La contraseña debe tener al menos 6 caracteres.');
  return c;
}
function bloqueo_(k) { if (Number(cache_().get(k) || 0) >= 5) throw new Error('Demasiados intentos fallidos. Espera 15 minutos.'); }
function fallo_(k) { const c = cache_(); c.put(k, String(Number(c.get(k) || 0) + 1), 900); }

/* ================= Empleados y lugares ================= */
function empleados_() {
  return rows_(H.EMP).map((r, i) => ({
    row: i + 2, correo: correo_(r[0]), nombre: r[1], cargo: r[2], lugar: r[3],
    activo: String(r[4]).trim().toUpperCase() === 'NO' ? 'NO' : 'SI', hash: r[5], sal: r[6],
    debeCambiar: String(r[7]).trim().toUpperCase() === 'SI', rol: normRol_(r[8])
  })).filter(e => e.correo);
}
function normRol_(r) {
  r = String(r || '').trim().toUpperCase();
  return r === 'ADMIN' ? 'Admin' : r === 'RRHH' ? 'RRHH' : r === 'SUPERVISOR' ? 'Supervisor' : 'Empleado';
}
function empPorCorreo_(correo) {
  const c = correo_(correo);
  return empleados_().find(e => e.correo === c) || null;
}
function publico_(e) { return { correo: e.correo, nombre: e.nombre, cargo: e.cargo, lugar: e.lugar, debeCambiar: e.debeCambiar, rol: e.rol }; }

function lugares_() {
  return rows_(H.LUG).filter(r => r[0]).map(r => ({
    nombre: String(r[0]).trim(), lat: Number(String(r[1]).replace(',', '.')),
    lng: Number(String(r[2]).replace(',', '.')), radio: Number(r[3]) || 150
  }));
}
// "Oficina Central, Bodega" -> lista de lugares; "Libre" -> sin restricción
function lugaresDe_(texto) {
  const nombres = String(texto || '').split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
  if (nombres.indexOf(LIBRE) >= 0) return LIBRE;
  const todos = lugares_();
  return todos.filter(l => nombres.indexOf(l.nombre.toUpperCase()) >= 0);
}

function validarUbicacion_(e, lat, lng, prec) {
  lat = Number(lat); lng = Number(lng); prec = Number(prec) || 0;
  const sinGps = !isFinite(lat) || !isFinite(lng) || (lat === 0 && lng === 0);
  const ls = lugaresDe_(e.lugar);
  if (ls === LIBRE) return { nombre: 'Libre', d: '', lat: sinGps ? '' : lat, lng: sinGps ? '' : lng };
  if (!ls.length) throw new Error('No tienes un lugar de trabajo asignado. Avisa a administración.');
  if (sinGps) throw new Error('No se pudo obtener tu ubicación. Activa el GPS y permite el acceso a la ubicación.');
  const max = Number(cfg_('PRECISION_MAX_METROS')) || 100;
  if (prec > max) throw new Error('La señal GPS es débil (±' + Math.round(prec) + ' m). Espera unos segundos cerca de una ventana o al aire libre e intenta de nuevo.');
  let mejor = null;
  ls.forEach(l => {
    const d = dist_(lat, lng, l.lat, l.lng);
    if (!mejor || d - l.radio < mejor.d - mejor.radio) mejor = { nombre: l.nombre, d: d, radio: l.radio };
  });
  if (mejor.d > mejor.radio) {
    const txt = mejor.d >= 1000 ? (mejor.d / 1000).toFixed(1) + ' km' : Math.round(mejor.d) + ' m';
    throw new Error('Estás a ' + txt + ' de ' + mejor.nombre + '. Solo puedes marcar dentro de ' + mejor.radio + ' m de tu lugar de trabajo.');
  }
  return { nombre: mejor.nombre, d: Math.round(mejor.d), lat: lat.toFixed(6), lng: lng.toFixed(6) };
}

/* ================= Sesiones ================= */
// Empleados: sesión de 30 días en el teléfono (no tienen que entrar de nuevo para marcar la salida).
// Administración: 6 horas.
const DIAS_SESION = 30;
function props_() { return PropertiesService.getScriptProperties(); }

function crearSesionEmp_(correo) {
  const t = Utilities.getUuid().replace(/-/g, '') + id_();
  props_().setProperty('S_' + t, correo + '|' + (Date.now() + DIAS_SESION * 86400000));
  return t;
}
function limpiarSesiones_() {
  const p = props_(), todas = p.getProperties(), ya = Date.now();
  Object.keys(todas).forEach(k => {
    if (k.indexOf('S_') === 0 && Number(todas[k].split('|')[1]) < ya) p.deleteProperty(k);
  });
}
function sesionEmp_(token) {
  const v = token ? props_().getProperty('S_' + token) : null;
  if (!v || Number(v.split('|')[1]) < Date.now()) throw new Error('SESION: Tu sesión venció. Vuelve a entrar.');
  const e = empPorCorreo_(v.split('|')[0]);
  if (!e || e.activo === 'NO') throw new Error('SESION: Tu usuario está inactivo. Habla con administración.');
  return e;
}
function crearSesionAdmin_() {
  const t = Utilities.getUuid() + id_();
  cache_().put('A_' + t, '1', 21600);
  return t;
}
// Quién está usando el panel: el PIN (Admin total) o un empleado con rol
function actor_(token, roles) {
  let a;
  if (token && cache_().get('A_' + token)) {
    cache_().put('A_' + token, '1', 21600);
    a = { rol: 'Admin', nombre: 'Administrador (PIN)', correo: '', lugares: null };
  } else {
    const e = sesionEmp_(token);
    if (e.rol === 'Empleado') throw new Error('No tienes permisos de administración.');
    a = { rol: e.rol, nombre: e.nombre, correo: e.correo, lugares: e.rol === 'Supervisor' ? nombresLugar_(e.lugar) : null };
  }
  if (roles && roles.indexOf(a.rol) < 0) throw new Error('Tu rol (' + a.rol + ') no permite hacer esto.');
  return a;
}
function nombresLugar_(t) {
  return String(t || '').split(',').map(s => s.trim().toUpperCase()).filter(s => s && s !== LIBRE);
}
// El supervisor solo ve a quienes comparten alguno de sus lugares
function enAlcance_(a, lugarEmp) {
  if (!a.lugares) return true;
  return nombresLugar_(lugarEmp).some(l => a.lugares.indexOf(l) >= 0);
}

/* ================= Acceso del empleado ================= */
function datosInicio_(e, token) {
  const ls = lugaresDe_(e.lugar);
  return {
    token: token, emp: publico_(e), estado: estado_(e.correo), hoy: hoy_(), empresa: cfg_('EMPRESA'),
    lugares: ls === LIBRE ? 'LIBRE' : ls, precisionMax: Number(cfg_('PRECISION_MAX_METROS')) || 100,
    ahoraMs: Date.now()
  };
}

function login(correo, clave) {
  const k = 'F_' + correo_(correo);
  bloqueo_(k);
  const e = empPorCorreo_(correo);
  if (!e) { fallo_(k); throw new Error('Ese correo no está registrado. Pide a administración que te agregue.'); }
  if (e.activo === 'NO') throw new Error('Tu usuario está inactivo. Habla con administración.');
  if (!e.hash) throw new Error('Todavía no tienes contraseña. Toca "Crear o recuperar contraseña".');
  if (hash_(String(clave || ''), e.sal) !== e.hash) { fallo_(k); throw new Error('Contraseña incorrecta.'); }
  cache_().remove(k);
  limpiarSesiones_();
  return datosInicio_(e, crearSesionEmp_(e.correo));
}

function sesion(token) {
  const e = sesionEmp_(token);
  return datosInicio_(e, token);
}

function cerrarSesion(token) { if (token) props_().deleteProperty('S_' + token); return true; }

function pedirCodigo(correo) {
  const c = correo_(correo);
  const e = empPorCorreo_(c);
  if (!e) throw new Error('Ese correo no está registrado. Pide a administración que te agregue.');
  if (e.activo === 'NO') throw new Error('Tu usuario está inactivo. Habla con administración.');
  const cc = cache_();
  if (cc.get('E_' + c)) throw new Error('Ya te enviamos un código hace menos de un minuto. Revisa tu correo (y la carpeta de spam).');
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('Hoy ya no se pueden enviar más correos. Pide a administración una contraseña temporal.');
  const codigo = String(Math.floor(100000 + Math.random() * 900000));
  cc.put('C_' + c, JSON.stringify({ codigo: codigo, n: 0 }), 900);
  cc.put('E_' + c, '1', 60);
  const empresa = cfg_('EMPRESA') || 'Marcaje';
  MailApp.sendEmail({
    to: c, name: empresa, subject: 'Tu código para la app de marcaje: ' + codigo,
    htmlBody: '<p>Hola ' + e.nombre + ',</p><p>Tu código para crear o cambiar tu contraseña es:</p>' +
      '<p style="font-size:28px;font-weight:bold;letter-spacing:4px">' + codigo + '</p>' +
      '<p>Vence en 15 minutos. Si no lo pediste, ignora este correo.</p><p>' + empresa + '</p>'
  });
  return 'Te enviamos un código a ' + c + '. Revisa también la carpeta de spam.';
}

function restablecerClave(correo, codigo, nueva) {
  const c = correo_(correo), cc = cache_();
  const raw = cc.get('C_' + c);
  if (!raw) throw new Error('El código venció o no existe. Pide uno nuevo.');
  const d = JSON.parse(raw);
  if (d.n >= 5) { cc.remove('C_' + c); throw new Error('Demasiados intentos. Pide un código nuevo.'); }
  if (String(codigo || '').trim() !== d.codigo) {
    d.n++; cc.put('C_' + c, JSON.stringify(d), 900);
    throw new Error('Código incorrecto.');
  }
  validarClave_(nueva);
  lock_(() => {
    const e = empPorCorreo_(c);
    const sal = id_();
    writeText_(sh_(H.EMP), e.row, [[hash_(nueva, sal), sal, 'NO']], 6);
  });
  cc.remove('C_' + c); cc.remove('F_' + c);
  return login(c, nueva);
}

function cambiarClave(token, actual, nueva) {
  const e = sesionEmp_(token);
  if (hash_(String(actual || ''), e.sal) !== e.hash) throw new Error('La contraseña actual no es correcta.');
  validarClave_(nueva);
  if (actual === nueva) throw new Error('La nueva contraseña debe ser distinta a la actual.');
  lock_(() => { const sal = id_(); writeText_(sh_(H.EMP), e.row, [[hash_(nueva, sal), sal, 'NO']], 6); });
  return 'Contraseña cambiada.';
}

/* ================= Marcaje ================= */
// Lee solo el final de la hoja (las marcas de hoy siempre están al final)
function filaHoy_(correo) { return filaDe_(correo, hoy_()); }

// Busca la marca de un empleado en una fecha (revisa las últimas 2000 filas;
// las marcas sin conexión pueden llegar fuera de orden)
function filaDe_(correo, fecha) {
  const sh = sh_(H.MAR), last = sh.getLastRow();
  if (last < 2) return null;
  const n = Math.min(last - 1, 2000), start = last - n + 1;
  const v = sh.getRange(start, 1, n, MAR_COLS.length).getDisplayValues();
  for (let i = v.length - 1; i >= 0; i--) {
    if (v[i][1] === fecha && correo_(v[i][2]) === correo) return { row: start + i, data: v[i] };
  }
  return null;
}

function marcasDeHoy_() {
  const sh = sh_(H.MAR), last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(last - 1, 1500), start = last - n + 1;
  const h = hoy_();
  return sh.getRange(start, 1, n, MAR_COLS.length).getDisplayValues().filter(r => r[1] === h);
}

function estado_(correo) {
  const f = filaHoy_(correo);
  if (!f) return { fase: 'entrada' };
  if (!f.data[5]) return { fase: 'salida', entrada: f.data[4] };
  return { fase: 'listo', entrada: f.data[4], salida: f.data[5], horas: f.data[6] };
}

function marcar(token, lat, lng, prec) {
  const e = sesionEmp_(token);
  if (e.debeCambiar) throw new Error('Primero cambia tu contraseña temporal.');
  const u = validarUbicacion_(e, lat, lng, prec);   // fuera del candado: no frena a los demás
  return lock_(() => {
    const sh = sh_(H.MAR);
    const st = estado_(e.correo);
    const h = hora_();
    let msg;
    if (st.fase === 'entrada') {
      writeText_(sh, sh.getLastRow() + 1, [[id_(), hoy_(), e.correo, e.nombre, h, '', '', u.lat, u.lng, u.d, '', '', '', u.nombre]]);
      msg = 'Entrada marcada a las ' + h.slice(0, 5);
    } else if (st.fase === 'salida') {
      const f = filaHoy_(e.correo);
      writeText_(sh, f.row, [[h, difHoras_(st.entrada, h).toFixed(2)]], 6);
      writeText_(sh, f.row, [[u.lat, u.lng, u.d]], 11);
      msg = 'Salida marcada a las ' + h.slice(0, 5);
    } else {
      throw new Error('Ya marcaste entrada y salida hoy.');
    }
    return { msg: msg, estado: estado_(e.correo), ahoraMs: Date.now() };
  });
}

/* ================= Marcas hechas sin conexión ================= */
// El teléfono guarda las marcas cuando no hay internet y las manda aquí al reconectarse.
// Cada marca trae: id (único), tipo (entrada/salida), tCorr (hora del teléfono corregida con la
// diferencia que tenía con el servidor), tMono (hora calculada con el reloj interno del navegador),
// lat, lng, prec. Se guardan con la observación "Sin conexión" para que el Admin las pueda revisar.
function sincronizarMarcas(token, lista) {
  const e = sesionEmp_(token);
  if (!Array.isArray(lista) || !lista.length) return { ok: [], estado: estado_(e.correo), ahoraMs: Date.now() };
  const ahora = Date.now(), ok = [];
  const fmt = (t, f) => Utilities.formatDate(new Date(t), tz_(), f);
  lock_(() => {
    const sh = sh_(H.MAR);
    const ids = idsSincronizados_();
    lista.slice(0, 20).sort((x, y) => Number(x.tCorr) - Number(y.tCorr)).forEach(m => {
      const id = String(m.id || '').replace(/[^\w-]/g, '').slice(0, 40);
      if (!id) return;
      if (ids[id]) { ok.push(id); return; }
      let t = Number(m.tCorr) || 0;
      if (!(t > 0)) return;
      const obs = [];
      if (t > ahora + 120000) { obs.push('hora del teléfono en el futuro, se usó la hora de llegada'); t = ahora; }
      if (ahora - t > 7 * 86400000) obs.push('marca de hace más de 7 días');
      const mono = Number(m.tMono) || 0;
      if (mono && Math.abs(mono - t) > 300000) obs.push('revisar: el reloj del teléfono difiere ' + Math.round(Math.abs(mono - t) / 60000) + ' min');
      let u;
      try { u = validarUbicacion_(e, m.lat, m.lng, m.prec); }
      catch (err) { u = { nombre: '', d: '', lat: m.lat || '', lng: m.lng || '' }; obs.push('ubicación: ' + err.message); }
      const fecha = fmt(t, 'yyyy-MM-dd'), hora = fmt(t, 'HH:mm:ss');
      const nota = (m.tipo === 'salida' ? 'Salida' : 'Entrada') + ' sin conexión, subida ' + fmt(ahora, 'dd/MM HH:mm') + (obs.length ? ' (' + obs.join('; ') + ')' : '');
      const f = filaDe_(e.correo, fecha);
      if (m.tipo === 'entrada') {
        if (!f) writeText_(sh, sh.getLastRow() + 1, [[id_(), fecha, e.correo, e.nombre, hora, '', '', u.lat, u.lng, u.d, '', '', '', u.nombre, nota, id]]);
      } else if (f && !f.data[5]) {
        writeText_(sh, f.row, [[hora, difHoras_(f.data[4], hora).toFixed(2)]], 6);
        writeText_(sh, f.row, [[u.lat, u.lng, u.d]], 11);
        writeText_(sh, f.row, [[(f.data[14] ? f.data[14] + '. ' : '') + nota, ((f.data[15] || '') + ' ' + id).trim()]], 15);
      } else if (!f) {
        writeText_(sh, sh.getLastRow() + 1, [[id_(), fecha, e.correo, e.nombre, '', hora, '', '', '', '', u.lat, u.lng, u.d, u.nombre, nota + ' (salida sin entrada)', id]]);
      }
      ids[id] = true;
      ok.push(id);
    });
  });
  return { ok: ok, estado: estado_(e.correo), ahoraMs: Date.now() };
}

function idsSincronizados_() {
  const sh = sh_(H.MAR), last = sh.getLastRow(), mapa = {};
  if (last < 2) return mapa;
  const n = Math.min(last - 1, 3000);
  sh.getRange(last - n + 1, 16, n, 1).getDisplayValues().forEach(r => String(r[0]).split(' ').forEach(x => { if (x) mapa[x] = true; }));
  return mapa;
}

function misMarcas(token) {
  const e = sesionEmp_(token);
  const sh = sh_(H.MAR), last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(last - 1, 15000), start = last - n + 1;   // ~3 meses con 200 empleados
  return sh.getRange(start, 1, n, MAR_COLS.length).getDisplayValues()
    .filter(r => correo_(r[2]) === e.correo).slice(-45).reverse()
    .map(r => ({ fecha: r[1], entrada: r[4], salida: r[5], horas: r[6], lugar: r[13] }));
}

/* ================= Permisos ================= */
function permisoObj_(r) {
  return { id: r[0], solicitado: r[1], correo: r[2], nombre: r[3], tipo: r[4], desde: r[5], hasta: r[6],
    horaDesde: r[7], horaHasta: r[8], motivo: r[9], estado: r[10], respuesta: r[11], revisado: r[12], revisadoPor: r[13] || '' };
}

function solicitarPermiso(token, p) {
  const e = sesionEmp_(token);
  if (TIPOS_PERMISO.indexOf(p.tipo) < 0) throw new Error('Elige el tipo de permiso.');
  if (!p.desde) throw new Error('Indica la fecha del permiso.');
  if (!String(p.motivo || '').trim()) throw new Error('Escribe el motivo.');
  if (p.hasta && p.hasta < p.desde) throw new Error('La fecha "hasta" no puede ser antes de "desde".');
  lock_(() => {
    const sh = sh_(H.PER);
    writeText_(sh, sh.getLastRow() + 1, [[id_(), ahora_(), e.correo, e.nombre, p.tipo, p.desde, p.hasta || p.desde,
      p.horaDesde || '', p.horaHasta || '', String(p.motivo).trim().slice(0, 500), 'Pendiente', '', '', '']]);
  });
  avisarPermiso_(e, p);   // si el correo falla, la solicitud igual queda guardada
  return 'Solicitud enviada. Queda pendiente de aprobación.';
}

/* ================= Avisos por correo ================= */
function correosAviso_() {
  return cfg_('AVISO_PERMISOS').split(/[,;\s]+/).map(correo_).filter(c => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c));
}

function avisarPermiso_(e, p) {
  try {
    const para = correosAviso_();
    if (!para.length || MailApp.getRemainingDailyQuota() < para.length) return;
    const x = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const fechas = (!p.hasta || p.hasta === p.desde) ? p.desde : p.desde + ' al ' + p.hasta;
    const horas = p.horaDesde ? (p.horaDesde + (p.horaHasta ? ' a ' + p.horaHasta : '')) : 'Todo el día';
    const fila = (k, v) => '<tr><td style="padding:6px 12px 6px 0;color:#6D6E71">' + k + '</td><td style="padding:6px 0;color:#414042"><b>' + x(v) + '</b></td></tr>';
    let url = '';
    url = cfg_('URL_APP');   // enlace de la app instalable (Config: URL_APP)
    if (!url) { try { url = ScriptApp.getService().getUrl(); } catch (err) {} }
    const html =
      '<div style="font-family:Raleway,Arial,sans-serif;max-width:520px">' +
      '<div style="background:#0A2240;color:#FFFFFF;padding:16px 20px;font-size:18px;font-weight:bold">Nueva solicitud de permiso</div>' +
      '<div style="padding:16px 20px;background:#FFFFFF;border:1px solid #E3E0E7">' +
      '<table style="border-collapse:collapse;font-size:15px">' +
      fila('Empleado', e.nombre) + fila('Correo', e.correo) + fila('Lugar', e.lugar) + fila('Tipo', p.tipo) +
      fila('Fecha', fechas) + fila('Horario', horas) + fila('Motivo', String(p.motivo).trim()) +
      '</table>' +
      (url ? '<p style="margin:18px 0 4px"><a href="' + url + '" style="background:#00758F;color:#FFFFFF;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:bold">Abrir la app para aprobar</a></p>' : '') +
      '</div><p style="font-size:12px;color:#6D6E71;padding:0 20px">Aviso automático del control de asistencia.</p></div>';
    MailApp.sendEmail({
      to: para.join(','), name: 'Control de asistencia',
      subject: 'Permiso pendiente: ' + e.nombre + ' (' + p.tipo + ', ' + fechas + ')',
      htmlBody: html
    });
  } catch (err) {
    console.error('No se pudo enviar el aviso de permiso: ' + err);
  }
}

function setCfg_(k, v) {
  const sh = sh_(H.CFG);
  const col = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getDisplayValues();
  for (let i = 1; i < col.length; i++) if (col[i][0] === k) { writeText_(sh, i + 1, [[v]], 2); return; }
  writeText_(sh, sh.getLastRow() + 1, [[k, v]]);
}

function guardarAvisos(tok, texto) {
  actor_(tok, ['Admin']);
  const lista = String(texto || '').split(/[,;\s]+/).map(correo_).filter(Boolean);
  const malos = lista.filter(c => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c));
  if (malos.length) throw new Error('Correo no válido: ' + malos.join(', '));
  const unicos = lista.filter((c, i) => lista.indexOf(c) === i);
  if (unicos.length > 10) throw new Error('Máximo 10 correos.');
  setCfg_('AVISO_PERMISOS', unicos.join(', '));
  return unicos.length ? 'Los avisos de permisos llegarán a ' + unicos.length + ' correo(s).' : 'Avisos por correo desactivados.';
}

function probarAviso(tok) {
  actor_(tok, ['Admin']);
  const para = correosAviso_();
  if (!para.length) throw new Error('Primero guarda al menos un correo.');
  avisarPermiso_({ nombre: 'Empleado de prueba', correo: 'prueba@ejemplo.com', lugar: 'Oficina' },
    { tipo: 'Permiso personal', desde: hoy_(), hasta: hoy_(), motivo: 'Este es un correo de prueba.' });
  return 'Correo de prueba enviado a ' + para.join(', ') + '.';
}

function misPermisos(token) {
  const e = sesionEmp_(token);
  return rows_(H.PER).filter(r => correo_(r[2]) === e.correo).slice(-25).reverse().map(permisoObj_);
}

/* ================= Administración ================= */
function adminLogin(pin) {
  bloqueo_('F_ADMIN');
  if (String(pin || '').trim() !== cfg_('PIN_ADMIN')) { fallo_('F_ADMIN'); throw new Error('PIN de administrador incorrecto.'); }
  cache_().remove('F_ADMIN');
  return crearSesionAdmin_();
}

function adminResumen(tok) {
  const a = actor_(tok);
  const h = hoy_();
  const todos = empleados_();
  const emps = todos.filter(e => enAlcance_(a, e.lugar));
  const visibles = {};
  emps.forEach(e => visibles[e.correo] = true);
  const marcas = marcasDeHoy_();
  const porCorreo = {};
  marcas.forEach(r => porCorreo[correo_(r[2])] = r);
  const permisos = rows_(H.PER).map(permisoObj_).filter(p => visibles[correo_(p.correo)]);
  const conPermiso = {};
  permisos.filter(p => p.estado === 'Aprobado' && p.desde <= h && p.hasta >= h).forEach(p => conPermiso[correo_(p.correo)] = p.tipo);

  const asistencia = emps.filter(e => e.activo === 'SI').map(e => {
    const m = porCorreo[e.correo];
    let estado = 'Sin marcar';
    if (m && m[5]) estado = 'Salió';
    else if (m) estado = 'Presente';
    else if (conPermiso[e.correo]) estado = 'Con permiso';
    return { correo: e.correo, nombre: e.nombre, cargo: e.cargo, estado: estado,
      entrada: m ? m[4] : '', salida: m ? m[5] : '', horas: m ? m[6] : '',
      latE: m ? m[7] : '', lngE: m ? m[8] : '', mE: m ? m[9] : '',
      latS: m ? m[10] : '', lngS: m ? m[11] : '', mS: m ? m[12] : '',
      lugar: m ? m[13] : e.lugar, permiso: conPermiso[e.correo] || '', sinConexion: m ? !!m[14] : false };
  });

  return {
    yo: { rol: a.rol, nombre: a.nombre, correo: a.correo },
    avisos: a.rol === 'Admin' ? cfg_('AVISO_PERMISOS') : '',
    hoy: h, empresa: cfg_('EMPRESA'), url: a.rol === 'Supervisor' ? '' : ss_().getUrl(),
    asistencia: asistencia,
    pendientes: permisos.filter(p => p.estado === 'Pendiente').reverse(),
    resueltos: permisos.filter(p => p.estado !== 'Pendiente').slice(-30).reverse(),
    empleados: a.rol === 'Supervisor' ? [] : emps.map(e => ({ correo: e.correo, nombre: e.nombre, cargo: e.cargo, lugar: e.lugar,
      activo: e.activo, tieneClave: !!e.hash, debeCambiar: e.debeCambiar, rol: e.rol })),
    lugares: lugares_()
  };
}

function responderPermiso(tok, id, estado, respuesta) {
  const a = actor_(tok);
  if (['Aprobado', 'Rechazado'].indexOf(estado) < 0) throw new Error('Estado no válido.');
  return lock_(() => {
    const sh = sh_(H.PER);
    const v = sh.getRange(1, 1, sh.getLastRow(), 3).getDisplayValues();
    for (let i = v.length - 1; i >= 1; i--) {
      if (v[i][0] === id) {
        const quien = correo_(v[i][2]);
        if (a.correo && quien === a.correo) throw new Error('No puedes aprobar ni rechazar tu propio permiso.');
        const emp = empPorCorreo_(quien);
        if (!emp || !enAlcance_(a, emp.lugar)) throw new Error('Este empleado no está a tu cargo.');
        writeText_(sh, i + 1, [[estado, respuesta || '', ahora_(), a.nombre]], 11);
        return 'Permiso ' + estado.toLowerCase();
      }
    }
    throw new Error('No se encontró la solicitud.');
  });
}

function guardarEmpleado(tok, d) {
  const a = actor_(tok, ['Admin', 'RRHH']);
  const rol = normRol_(d.rol);
  if (a.rol === 'RRHH' && (rol === 'Admin' || rol === 'RRHH')) throw new Error('Solo un Admin puede asignar el rol ' + rol + '.');
  const correo = correo_(d.correo);
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) throw new Error('Escribe un correo válido.');
  if (!String(d.nombre || '').trim()) throw new Error('Escribe el nombre.');
  const lugar = String(d.lugar || '').trim();
  if (!lugar) throw new Error('Asigna al menos un lugar de trabajo (o "Libre").');
  if (lugaresDe_(lugar) !== LIBRE) {
    const validos = lugares_().map(l => l.nombre.toUpperCase());
    const malos = lugar.split(',').map(s => s.trim()).filter(s => s && validos.indexOf(s.toUpperCase()) < 0);
    if (malos.length) throw new Error('No existe el lugar: ' + malos.join(', '));
  }
  const temporal = String(d.claveTemporal || '');
  if (temporal) validarClave_(temporal);
  return lock_(() => {
    const sh = sh_(H.EMP);
    const ex = empPorCorreo_(d.original || correo);
    if (ex && a.rol === 'RRHH' && (ex.rol === 'Admin' || ex.rol === 'RRHH')) throw new Error('Solo un Admin puede modificar a este usuario.');
    if (ex && a.correo && ex.correo === a.correo && ex.rol === 'Admin' && rol !== 'Admin') throw new Error('No puedes quitarte tu propio rol de Admin.');
    if (!ex || correo !== ex.correo) {
      if (empPorCorreo_(correo)) throw new Error('Ya existe un empleado con ese correo.');
    }
    let hash = ex ? ex.hash : '', sal = ex ? ex.sal : '', debe = ex && ex.debeCambiar ? 'SI' : 'NO';
    if (temporal) { sal = id_(); hash = hash_(temporal, sal); debe = 'SI'; }
    const fila = [[correo, String(d.nombre).trim(), String(d.cargo || '').trim(), lugar, d.activo === 'NO' ? 'NO' : 'SI', hash, sal, debe, rol]];
    writeText_(sh, ex ? ex.row : sh.getLastRow() + 1, fila);
    return ex ? 'Empleado actualizado' : 'Empleado agregado';
  });
}

function guardarLugar(tok, l) {
  actor_(tok, ['Admin']);
  const nombre = String(l.nombre || '').trim();
  const lat = Number(l.lat), lng = Number(l.lng), radio = Math.round(Number(l.radio));
  if (!nombre || nombre.indexOf(',') >= 0) throw new Error('Escribe un nombre para el lugar (sin comas).');
  if (nombre.toUpperCase() === LIBRE) throw new Error('"Libre" es un nombre reservado.');
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180 || (!lat && !lng)) throw new Error('Coordenadas no válidas.');
  if (!(radio >= 20 && radio <= 5000)) throw new Error('El radio debe estar entre 20 y 5000 metros.');
  return lock_(() => {
    const sh = sh_(H.LUG);
    const v = sh.getDataRange().getDisplayValues();
    const buscar = String(l.original || nombre).toUpperCase();
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][0]).trim().toUpperCase() === buscar) {
        writeText_(sh, i + 1, [[nombre, lat.toFixed(6), lng.toFixed(6), radio]]);
        return 'Lugar actualizado';
      }
    }
    writeText_(sh, sh.getLastRow() + 1, [[nombre, lat.toFixed(6), lng.toFixed(6), radio]]);
    return 'Lugar agregado';
  });
}

function cambiarPinAdmin(tok, nuevo) {
  actor_(tok, ['Admin']);
  if (!/^\d{4,8}$/.test(String(nuevo || ''))) throw new Error('El PIN debe tener de 4 a 8 números.');
  const sh = sh_(H.CFG);
  const v = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === 'PIN_ADMIN') { writeText_(sh, i + 1, [[nuevo]], 2); return 'PIN de administrador cambiado'; }
}

/* ================= Reportes (solo Admin) ================= */
const TIPOS_REPORTE = {
  resumen: 'Resumen por empleado',
  marcas: 'Detalle de marcas',
  incompletas: 'Marcas sin salida',
  permisos: 'Permisos'
};

// Filtro de lugar: '' = todos, 'LIBRE' = empleados libres, o el nombre de un lugar
function enLugar_(lugarEmp, filtro) {
  if (!filtro) return true;
  const f = String(filtro).trim().toUpperCase();
  const ls = String(lugarEmp || '').split(',').map(x => x.trim().toUpperCase());
  return ls.indexOf(f) >= 0;
}

function generarReporte_(a, tipo, desde, hasta, lugar) {
  if (!TIPOS_REPORTE[tipo]) throw new Error('Tipo de reporte no válido.');
  if (!desde || !hasta || hasta < desde) throw new Error('Revisa el rango de fechas.');
  const emps = {};
  empleados_().filter(e => enAlcance_(a, e.lugar) && enLugar_(e.lugar, lugar)).forEach(e => emps[e.correo] = e);
  const titulo = TIPOS_REPORTE[tipo] + ' del ' + desde + ' al ' + hasta + (lugar ? ' (' + lugar + ')' : '');
  const marcas = () => rows_(H.MAR).filter(r => r[1] >= desde && r[1] <= hasta && emps[correo_(r[2])]);

  if (tipo === 'marcas') {
    return { titulo: titulo, encabezados: ['Fecha', 'Empleado', 'Correo', 'Entrada', 'Salida', 'Horas', 'Lugar', 'Metros entrada', 'Metros salida', 'Observación'],
      filas: marcas().map(r => [r[1], emps[correo_(r[2])].nombre, correo_(r[2]), r[4], r[5], r[6] ? Number(r[6]) : '', r[13], r[9], r[12], r[14] || '']) };
  }
  if (tipo === 'incompletas') {
    return { titulo: titulo, encabezados: ['Fecha', 'Empleado', 'Correo', 'Entrada', 'Lugar'],
      filas: marcas().filter(r => !r[5]).map(r => [r[1], emps[correo_(r[2])].nombre, correo_(r[2]), r[4], r[13]]) };
  }
  if (tipo === 'permisos') {
    return { titulo: titulo, encabezados: ['Solicitado', 'Empleado', 'Correo', 'Tipo', 'Desde', 'Hasta', 'Hora inicio', 'Hora fin', 'Motivo', 'Estado', 'Respuesta', 'Revisado por', 'Revisado el'],
      filas: rows_(H.PER).filter(r => r[5] <= hasta && r[6] >= desde && emps[correo_(r[2])])
        .map(r => [r[1], emps[correo_(r[2])].nombre, correo_(r[2]), r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[13] || '', r[12]]) };
  }
  // resumen
  const acc = {};
  Object.keys(emps).forEach(c => acc[c] = { e: emps[c], dias: 0, horas: 0, inc: 0, perm: 0 });
  marcas().forEach(r => {
    const x = acc[correo_(r[2])];
    x.dias++;
    if (r[5]) x.horas += Number(r[6]) || 0; else x.inc++;
  });
  rows_(H.PER).filter(r => r[10] === 'Aprobado' && r[5] <= hasta && r[6] >= desde).forEach(r => {
    const x = acc[correo_(r[2])]; if (x) x.perm++;
  });
  const filas = Object.keys(acc).map(c => {
    const x = acc[c], completos = x.dias - x.inc;
    return [x.e.nombre, c, x.e.lugar, x.dias, Math.round(x.horas * 100) / 100,
      completos ? Math.round(x.horas / completos * 100) / 100 : 0, x.inc, x.perm];
  }).sort((p, q) => String(p[0]).localeCompare(String(q[0])));
  return { titulo: titulo, encabezados: ['Empleado', 'Correo', 'Lugar', 'Días con marca', 'Horas totales', 'Promedio horas/día', 'Sin salida', 'Permisos aprobados'], filas: filas };
}

function reporte(tok, desde, hasta, tipo, lugar) {
  const a = actor_(tok, ['Admin']);
  return generarReporte_(a, tipo || 'resumen', desde, hasta, lugar || '');
}

// Crea el reporte como hoja de Google en la carpeta "Marcaje - Reportes"
function reporteADrive(tok, desde, hasta, tipo, lugar) {
  const a = actor_(tok, ['Admin']);
  const r = generarReporte_(a, tipo || 'resumen', desde, hasta, lugar || '');
  const ss = SpreadsheetApp.create(r.titulo);
  const sh = ss.getSheets()[0];
  sh.setName('Reporte');
  const datos = [r.encabezados].concat(r.filas);
  sh.getRange(1, 1, datos.length, r.encabezados.length).setValues(datos);
  sh.getRange(1, 1, 1, r.encabezados.length).setFontWeight('bold').setBackground('#0A2240').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, r.encabezados.length);
  const archivo = DriveApp.getFileById(ss.getId());
  archivo.moveTo(carpeta_('CARPETA_REPORTES', 'Marcaje - Reportes'));
  if (a.correo) { try { archivo.addEditor(a.correo); } catch (e) { /* el correo no es cuenta de Google */ } }
  return { url: ss.getUrl(), filas: r.filas.length };
}

/* ================= Respaldo mensual ================= */
function carpeta_(prop, nombre) {
  const p = props_(), id = p.getProperty(prop);
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const f = DriveApp.createFolder(nombre);
  p.setProperty(prop, f.getId());
  return f;
}

// La ejecuta el activador automático el día 1 de cada mes (y el botón "Hacer respaldo ahora")
function respaldoMensual() {
  SpreadsheetApp.flush();
  const carpeta = carpeta_('CARPETA_RESPALDOS', 'Marcaje - Respaldos');
  const nombre = 'Respaldo Marcaje ' + Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HHmm');
  const copia = DriveApp.getFileById(ss_().getId()).makeCopy(nombre, carpeta);
  props_().setProperty('ULTIMO_RESPALDO', ahora_() + '|' + copia.getUrl());
  // Conserva solo los últimos N respaldos (Config: RESPALDOS_A_GUARDAR, 12 por defecto)
  const guardar = Number(cfg_('RESPALDOS_A_GUARDAR')) || 12;
  const lista = [], it = carpeta.getFiles();
  while (it.hasNext()) lista.push(it.next());
  lista.sort((x, y) => y.getDateCreated() - x.getDateCreated());
  lista.slice(guardar).forEach(f => f.setTrashed(true));
  return copia.getUrl();
}

function respaldoActivo_() {
  return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'respaldoMensual');
}

function activarRespaldo_() {
  if (respaldoActivo_()) return;
  ScriptApp.newTrigger('respaldoMensual').timeBased().onMonthDay(1).atHour(2).create();
}

function estadoRespaldo(tok) {
  actor_(tok, ['Admin']);
  const u = props_().getProperty('ULTIMO_RESPALDO') || '';
  const idc = props_().getProperty('CARPETA_RESPALDOS');
  return { activo: respaldoActivo_(), ultimo: u.split('|')[0] || '', urlUltimo: u.split('|')[1] || '',
    urlCarpeta: idc ? 'https://drive.google.com/drive/folders/' + idc : '' };
}

function respaldarAhora(tok) {
  actor_(tok, ['Admin']);
  return respaldoMensual();
}

function activarRespaldoMensual(tok) {
  actor_(tok, ['Admin']);
  activarRespaldo_();
  return 'Respaldo automático activado: el día 1 de cada mes a las 2 a. m.';
}

function obtenerEmpresa() { return cfg_('EMPRESA'); }

/* Ejecuta esto una vez desde el editor: autoriza correos y Drive, y activa el respaldo mensual */
function autorizar() {
  setup_();
  activarRespaldo_();
  Logger.log('Respaldo mensual activo: ' + respaldoActivo_());
  Logger.log('Hoja de datos: ' + ss_().getUrl());
  Logger.log('Correos restantes hoy: ' + MailApp.getRemainingDailyQuota());
}
