/**
 * MARCAJE DE ENTRADA Y SALIDA — Google Apps Script  (v2)
 * Creada por el área de Desarrollo Tecnológico. Autoría: Amigos for Christ, Inc.
 *
 * - Ingreso con correo y contraseña (guardada cifrada, nunca en texto).
 * - Cada empleado crea, cambia o recupera su contraseña con un código por correo.
 * - Solo se puede marcar dentro del lugar de trabajo asignado (geocerca).
 * - Jornada corrida: horas = salida - entrada (sin descontar almuerzo).
 * Hojas: Empleados, Lugares, Marcas, Permisos, Config (se crean solas).
 */

const H = { EMP: 'Empleados', LUG: 'Lugares', MAR: 'Marcas', PER: 'Permisos', CFG: 'Config', DIS: 'Dispositivos' };
const EMP_COLS = ['Correo', 'Nombre', 'Cargo', 'Lugar', 'Activo', 'Hash', 'Sal', 'DebeCambiar', 'Rol', 'Lider'];
const MAR_COLS = ['ID', 'Fecha', 'Correo', 'Nombre', 'Entrada', 'Salida', 'Horas', 'LatEntrada', 'LngEntrada', 'MetrosEntrada', 'LatSalida', 'LngSalida', 'MetrosSalida', 'Lugar', 'Observacion', 'IdsCliente'];
const PER_COLS = ['ID', 'Solicitado', 'Correo', 'Nombre', 'Tipo', 'Desde', 'Hasta', 'HoraDesde', 'HoraHasta', 'Motivo', 'Estado', 'Respuesta', 'Revisado', 'RevisadoPor',
  'Lider', 'RevisadoLider', 'RevisadoLiderPor', 'RespuestaLider', 'ReponerFecha', 'ReponerHoras'];
// Estados de un permiso: 'Pendiente líder' (espera al líder) -> 'Pendiente' (espera a Administración) -> 'Aprobado' o 'Rechazado'
const PEND_LIDER = 'Pendiente líder';
// Roles: Empleado (solo marca), Supervisor (asistencia y permisos de su lugar),
// RRHH (todo el personal, crea empleados), Admin (todo, incluidos roles, lugares y PIN)
const ROLES = ['Empleado', 'Supervisor', 'RRHH', 'Admin'];
const TIPOS_PERMISO = ['Permiso personal', 'Cita médica', 'Vacaciones', 'Llegada tarde', 'Salida temprana', 'Otro'];
const LIBRE = 'LIBRE';             // Lugar = "Libre": puede marcar desde cualquier sitio (se guarda el GPS igual)

/* ================= Web app ================= */
function doGet() {
  setupUnaVez_();
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
    obtenerEmpresa, login, loginMicrosoft, sesion, cerrarSesion, pedirCodigo, restablecerClave, cambiarClave,
    marcar, sincronizarMarcas, misMarcas, solicitarPermiso, misPermisos, equipoLider, responderPermisoLider,
    adminLogin, adminResumen, responderPermiso, guardarEmpleado, guardarLugar, cambiarPinAdmin,
    reporte, reporteADrive, estadoRespaldo, respaldarAhora, activarRespaldoMensual,
    guardarAvisos, probarAviso,
    registrarDispositivo, quitarDispositivo, probarNotificacion, estadoRecordatorios, activarRecordatorios
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
    setupUnaVez_();
    res = { ok: true, data: api[fn].apply(null, args) };
  } catch (err) {
    res = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= Utilidades ================= */
/* Rendimiento: lo que se lee una vez se reutiliza durante la misma petición */
let _ss = null, _hojas = {}, _cfg = null, _emps = null, _lugs = null;
function olvidarMemoria_() { _cfg = null; _emps = null; _lugs = null; }

// Revisa/crea las hojas solo una vez cada 6 horas (no en cada petición)
function setupUnaVez_() {
  const c = cache_();
  if (c.get('SETUP_V6')) return;
  setup_();
  c.put('SETUP_V6', '1', 21600);
}

function ss_() {
  if (_ss) return _ss;
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SS_ID');
  if (id) return (_ss = SpreadsheetApp.openById(id));
  const ss = SpreadsheetApp.create('Marcaje - Base de datos');
  p.setProperty('SS_ID', ss.getId());
  return (_ss = ss);
}
function sh_(n) { return _hojas[n] || (_hojas[n] = ss_().getSheetByName(n)); }
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
  ensure_(ss, H.DIS, ['Correo', 'Token', 'Plataforma', 'Actualizado']);
  migrar_(H.EMP, EMP_COLS);
  migrar_(H.PER, PER_COLS);
  migrar_(H.MAR, MAR_COLS);
  const cfg = ensure_(ss, H.CFG, ['Clave', 'Valor']);
  if (cfg.getLastRow() < 2) {
    writeText_(cfg, 2, [['PIN_ADMIN', '1234'], ['EMPRESA', 'Mi Empresa'], ['PRECISION_MAX_METROS', '100']]);
  }
  // Horario de los recordatorios (se pueden cambiar en la hoja Config)
  const claves = cfg.getRange(1, 1, Math.max(cfg.getLastRow(), 1), 1).getDisplayValues().map(r => r[0]);
  [['DIAS_LABORALES', '1,2,3,4,5'], ['AVISOS_ENTRADA', '08:00,08:15,08:30'], ['AVISOS_SALIDA', '17:00'], ['FERIADOS', ''], ['DOMINIO_MICROSOFT', 'amigosforchrist.org'], ['HORAS_DIA', '8']]
    .forEach(([k, v]) => { if (claves.indexOf(k) < 0) writeText_(cfg, cfg.getLastRow() + 1, [[k, v]]); });
  _cfg = null;
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
  if (!_cfg) { _cfg = {}; rows_(H.CFG).forEach(r => { if (r[0]) _cfg[r[0]] = String(r[1]).trim(); }); }
  return _cfg[k] || '';
}

function lock_(fn) {
  const l = LockService.getScriptLock();
  if (!l.tryLock(30000)) throw new Error('Hay mucha gente marcando en este momento. Intenta de nuevo en unos segundos.');
  olvidarMemoria_();   // dentro del candado se trabaja con datos frescos
  try { return fn(); } finally { l.releaseLock(); olvidarMemoria_(); }
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
  if (_emps) return _emps;
  return _emps = rows_(H.EMP).map((r, i) => ({
    row: i + 2, correo: correo_(r[0]), nombre: r[1], cargo: r[2], lugar: r[3],
    activo: String(r[4]).trim().toUpperCase() === 'NO' ? 'NO' : 'SI', hash: r[5], sal: r[6],
    debeCambiar: String(r[7]).trim().toUpperCase() === 'SI', rol: normRol_(r[8]), liderTxt: String(r[9] || '').trim()
  })).filter(e => e.correo);
}
// Líder del empleado (columna Lider: su correo; también se acepta el nombre exacto).
// Devuelve el empleado líder activo, o null si no tiene o no es válido.
function liderDe_(e) {
  const t = String(e && e.liderTxt || '').trim();
  if (!t) return null;
  const c = correo_(t);
  const l = empleados_().find(x => x.correo === c) || empleados_().find(x => String(x.nombre).trim().toLowerCase() === t.toLowerCase());
  if (!l || l.activo === 'NO' || l.correo === e.correo) return null;
  return l;
}
// Correos de las personas que tienen a este empleado como líder
function equipoDe_(correoLider) {
  return empleados_().filter(x => x.activo === 'SI' && (liderDe_(x) || {}).correo === correoLider);
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
  if (_lugs) return _lugs;
  return _lugs = rows_(H.LUG).filter(r => r[0]).map(r => ({
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

function crearSesionEmp_(correo, viaMicrosoft) {
  const t = Utilities.getUuid().replace(/-/g, '') + id_();
  props_().setProperty('S_' + t, correo + '|' + (Date.now() + DIAS_SESION * 86400000) + (viaMicrosoft ? '|ms' : ''));
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
  // Quien entró con Microsoft no usa la contraseña de la app: no se le pide cambiar la temporal
  return v.split('|')[2] === 'ms' ? Object.assign({}, e, { debeCambiar: false }) : e;
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
    ahoraMs: Date.now(), esLider: equipoDe_(e.correo).length > 0
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

// Entrar con la cuenta de Microsoft 365: el teléfono manda el permiso que Microsoft le dio,
// y el servidor le pregunta a Microsoft quién es esa persona (no se puede falsificar).
function loginMicrosoft(accessToken) {
  if (!accessToken || String(accessToken).length < 50) throw new Error('No se recibió la autorización de Microsoft.');
  const r = UrlFetchApp.fetch('https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName,displayName',
    { headers: { Authorization: 'Bearer ' + accessToken }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Microsoft no confirmó tu identidad. Intenta entrar de nuevo.');
  const me = JSON.parse(r.getContentText() || '{}');
  const opciones = [me.mail, me.userPrincipalName].map(correo_).filter(Boolean);
  if (!opciones.length) throw new Error('Tu cuenta de Microsoft no tiene un correo asignado.');
  const dominio = (cfg_('DOMINIO_MICROSOFT') || '').toLowerCase().replace(/^@/, '');
  if (dominio && !opciones.some(c => c.split('@')[1] === dominio)) throw new Error('Entra con tu cuenta de @' + dominio + '.');
  const e = opciones.map(empPorCorreo_).find(Boolean);
  if (!e) throw new Error('Tu cuenta ' + opciones[0] + ' no está registrada en la app. Pide a administración que te agregue.');
  if (e.activo === 'NO') throw new Error('Tu usuario está inactivo. Habla con administración.');
  limpiarSesiones_();
  const datos = datosInicio_(Object.assign({}, e, { debeCambiar: false }), crearSesionEmp_(e.correo, true));
  datos.viaMicrosoft = true;
  return datos;
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
// Para ir rápido recuerda en caché qué fila tiene la marca de cada persona en el día,
// y si no la tiene en caché revisa solo las columnas Fecha y Correo.
function filaDe_(correo, fecha) {
  const sh = sh_(H.MAR), c = cache_(), k = claveFila_(correo, fecha);
  const guardada = Number(c.get(k)) || 0;
  if (guardada >= 2) {
    const d = sh.getRange(guardada, 1, 1, MAR_COLS.length).getDisplayValues()[0];
    if (d[1] === fecha && correo_(d[2]) === correo) return { row: guardada, data: d };
  }
  const last = sh.getLastRow();
  if (last < 2) return null;
  const n = Math.min(last - 1, 2000), start = last - n + 1;
  const v = sh.getRange(start, 2, n, 2).getDisplayValues();
  for (let i = v.length - 1; i >= 0; i--) {
    if (v[i][0] === fecha && correo_(v[i][1]) === correo) {
      const row = start + i;
      c.put(k, String(row), 21600);
      return { row: row, data: sh.getRange(row, 1, 1, MAR_COLS.length).getDisplayValues()[0] };
    }
  }
  return null;
}
function claveFila_(correo, fecha) { return 'R_' + fecha + '|' + correo; }
function recordarFila_(correo, fecha, row) { cache_().put(claveFila_(correo, fecha), String(row), 21600); }
function estadoDeFila_(f) {
  if (!f) return { fase: 'entrada' };
  if (!f.data[5]) return { fase: 'salida', entrada: f.data[4] };
  return { fase: 'listo', entrada: f.data[4], salida: f.data[5], horas: f.data[6] };
}

function marcasDeHoy_() {
  const sh = sh_(H.MAR), last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(last - 1, 1500), start = last - n + 1;
  const h = hoy_();
  return sh.getRange(start, 1, n, MAR_COLS.length).getDisplayValues().filter(r => r[1] === h);
}

function estado_(correo) { return estadoDeFila_(filaHoy_(correo)); }

function marcar(token, lat, lng, prec) {
  const e = sesionEmp_(token);
  if (e.debeCambiar) throw new Error('Primero cambia tu contraseña temporal.');
  const u = validarUbicacion_(e, lat, lng, prec);   // fuera del candado: no frena a los demás
  return lock_(() => {
    const sh = sh_(H.MAR);
    const hoy = hoy_(), f = filaDe_(e.correo, hoy), st = estadoDeFila_(f);
    const h = hora_();
    let msg, nuevo;
    if (st.fase === 'entrada') {
      const row = sh.getLastRow() + 1;
      writeText_(sh, row, [[id_(), hoy, e.correo, e.nombre, h, '', '', u.lat, u.lng, u.d, '', '', '', u.nombre]]);
      recordarFila_(e.correo, hoy, row);
      msg = 'Entrada marcada a las ' + h.slice(0, 5);
      nuevo = { fase: 'salida', entrada: h };
    } else if (st.fase === 'salida') {
      const horas = difHoras_(st.entrada, h).toFixed(2);
      writeText_(sh, f.row, [[h, horas, f.data[7], f.data[8], f.data[9], u.lat, u.lng, u.d]], 6);
      msg = 'Salida marcada a las ' + h.slice(0, 5);
      nuevo = { fase: 'listo', entrada: st.entrada, salida: h, horas: horas };
    } else {
      throw new Error('Ya marcaste entrada y salida hoy.');
    }
    return { msg: msg, estado: nuevo, ahoraMs: Date.now() };
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
      const forma = m.motivo === 'ocupado' ? ' guardada en el teléfono (servidor ocupado), subida ' : ' sin conexión, subida ';
      const nota = (m.tipo === 'salida' ? 'Salida' : 'Entrada') + forma + fmt(ahora, 'dd/MM HH:mm') + (obs.length ? ' (' + obs.join('; ') + ')' : '');
      const f = filaDe_(e.correo, fecha);
      if (m.tipo === 'entrada') {
        if (!f) { const row = sh.getLastRow() + 1; writeText_(sh, row, [[id_(), fecha, e.correo, e.nombre, hora, '', '', u.lat, u.lng, u.d, '', '', '', u.nombre, nota, id]]); recordarFila_(e.correo, fecha, row); }
      } else if (f && !f.data[5]) {
        writeText_(sh, f.row, [[hora, difHoras_(f.data[4], hora).toFixed(2), f.data[7], f.data[8], f.data[9], u.lat, u.lng, u.d, f.data[13],
          (f.data[14] ? f.data[14] + '. ' : '') + nota, ((f.data[15] || '') + ' ' + id).trim()]], 6);
      } else if (!f) {
        const row = sh.getLastRow() + 1;
        writeText_(sh, row, [[id_(), fecha, e.correo, e.nombre, '', hora, '', '', '', '', u.lat, u.lng, u.d, u.nombre, nota + ' (salida sin entrada)', id]]);
        recordarFila_(e.correo, fecha, row);
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
  const n = Math.min(last - 1, 12000), start = last - n + 1;   // ~2 meses con 200 empleados
  const v = sh.getRange(start, 2, n, 6).getDisplayValues();     // Fecha..Horas
  const lug = sh.getRange(start, 14, n, 1).getDisplayValues();  // Lugar
  const out = [];
  for (let i = v.length - 1; i >= 0 && out.length < 45; i--) {
    if (correo_(v[i][1]) === e.correo) out.push({ fecha: v[i][0], entrada: v[i][3], salida: v[i][4], horas: v[i][5], lugar: lug[i][0] });
  }
  return out;
}

/* ================= Permisos ================= */
function permisoObj_(r) {
  return { id: r[0], solicitado: r[1], correo: r[2], nombre: r[3], tipo: r[4], desde: r[5], hasta: r[6],
    horaDesde: r[7], horaHasta: r[8], motivo: r[9], estado: r[10], respuesta: r[11], revisado: r[12], revisadoPor: r[13] || '',
    lider: r[14] || '', revisadoLider: r[15] || '', revisadoLiderPor: r[16] || '', respuestaLider: r[17] || '',
    reponerFecha: r[18] || '', reponerHoras: r[19] || '' };
}

function solicitarPermiso(token, p) {
  const e = sesionEmp_(token);
  if (TIPOS_PERMISO.indexOf(p.tipo) < 0) throw new Error('Elige el tipo de permiso.');
  if (!p.desde) throw new Error('Indica la fecha del permiso.');
  if (!String(p.motivo || '').trim()) throw new Error('Escribe el motivo.');
  if (p.hasta && p.hasta < p.desde) throw new Error('La fecha "hasta" no puede ser antes de "desde".');
  const lider = liderDe_(e);
  const estado = lider ? PEND_LIDER : 'Pendiente';
  const id = id_();
  lock_(() => {
    const sh = sh_(H.PER);
    writeText_(sh, sh.getLastRow() + 1, [[id, ahora_(), e.correo, e.nombre, p.tipo, p.desde, p.hasta || p.desde,
      p.horaDesde || '', p.horaHasta || '', String(p.motivo).trim().slice(0, 500), estado, '', '', '',
      lider ? lider.correo : '', '', '', '', '', '']]);
  });
  // Si el correo o la notificación fallan, la solicitud igual queda guardada
  if (lider) {
    avisarPermiso_(e, p, { para: [lider.correo], titulo: 'Un colaborador de tu equipo pidió permiso',
      asunto: 'Permiso por aprobar (líder): ', boton: 'Abrir la app para aprobar', link: '#equipo' });
    pushA_([lider.correo], 'Permiso por aprobar', e.nombre + ' pidió ' + p.tipo.toLowerCase() + ' (' + rangoTxt_(p) + ').', 'permiso', '#equipo');
    avisarPermiso_(e, p, { titulo: 'Nueva solicitud de permiso (primero la revisa el líder)', asunto: 'Permiso solicitado (espera al líder): ',
      filasExtra: [['Líder', lider.nombre]], boton: 'Abrir la app' });
    return 'Solicitud enviada. Primero la revisa tu líder (' + lider.nombre + ') y luego Administración.';
  }
  avisarPermiso_(e, p);
  pushA_(correosAviso_(), 'Permiso por aprobar', e.nombre + ' pidió ' + p.tipo.toLowerCase() + ' (' + rangoTxt_(p) + ').', 'permiso');
  return 'Solicitud enviada. Queda pendiente de aprobación.';
}

function rangoTxt_(p) {
  const h = p.hasta || p.desde;
  return (!h || h === p.desde) ? p.desde : p.desde + ' al ' + h;
}

/* ================= Avisos por correo ================= */
function correosAviso_() {
  return cfg_('AVISO_PERMISOS').split(/[,;\s]+/).map(correo_).filter(c => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c));
}

// o (opcional): { para, titulo, asunto, filasExtra: [[k, v]], boton, link }
// Sin "para" se envía a la lista AVISO_PERMISOS (Administración).
function avisarPermiso_(e, p, o) {
  o = o || {};
  try {
    const para = (o.para || correosAviso_()).filter(c => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c));
    if (!para.length || MailApp.getRemainingDailyQuota() < para.length) return;
    const x = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const fechas = rangoTxt_(p);
    const horas = p.horaDesde ? (p.horaDesde + (p.horaHasta ? ' a ' + p.horaHasta : '')) : 'Todo el día';
    const fila = (k, v) => '<tr><td style="padding:6px 12px 6px 0;color:#6D6E71">' + x(k) + '</td><td style="padding:6px 0;color:#414042"><b>' + x(v) + '</b></td></tr>';
    let url = cfg_('URL_APP');   // enlace de la app instalable (Config: URL_APP)
    if (!url) { try { url = ScriptApp.getService().getUrl(); } catch (err) {} }
    if (url && o.link) url = url.replace(/#.*$/, '') + o.link;
    const extra = (o.filasExtra || []).filter(f => f[1] !== '' && f[1] != null).map(f => fila(f[0], f[1])).join('');
    const html =
      '<div style="font-family:Raleway,Arial,sans-serif;max-width:520px">' +
      '<div style="background:#0A2240;color:#FFFFFF;padding:16px 20px;font-size:18px;font-weight:bold">' + x(o.titulo || 'Nueva solicitud de permiso') + '</div>' +
      '<div style="padding:16px 20px;background:#FFFFFF;border:1px solid #E3E0E7">' +
      '<table style="border-collapse:collapse;font-size:15px">' +
      fila('Empleado', e.nombre) + fila('Correo', e.correo) + fila('Lugar', e.lugar) + fila('Tipo', p.tipo) +
      fila('Fecha', fechas) + fila('Horario', horas) + fila('Motivo', String(p.motivo || '').trim()) + extra +
      '</table>' +
      (url ? '<p style="margin:18px 0 4px"><a href="' + x(url) + '" style="background:#00758F;color:#FFFFFF;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:bold">' + x(o.boton || 'Abrir la app para aprobar') + '</a></p>' : '') +
      '</div><p style="font-size:12px;color:#6D6E71;padding:0 20px">Aviso automático del control de asistencia.</p></div>';
    MailApp.sendEmail({
      to: para.join(','), name: 'Control de asistencia',
      subject: (o.asunto || 'Permiso pendiente: ') + e.nombre + ' (' + p.tipo + ', ' + fechas + ')',
      htmlBody: html
    });
  } catch (err) {
    console.error('No se pudo enviar el aviso de permiso: ' + err);
  }
}

// Notificación al teléfono de estas personas (si tienen recordatorios activados). Nunca detiene el proceso.
function pushA_(correos, titulo, cuerpo, tipo, link) {
  try {
    if (!props_().getProperty('FCM_CUENTA_SERVICIO')) return;
    const dev = dispositivos_(), lista = [];
    (correos || []).map(correo_).filter((c, i, a) => c && a.indexOf(c) === i).forEach(c => (dev[c] || []).forEach(t =>
      lista.push({ token: t, titulo: titulo, cuerpo: cuerpo, tipo: tipo || 'permiso', link: link || '' })));
    if (lista.length) enviarPush_(lista);
  } catch (err) {
    console.error('No se pudo enviar la notificación: ' + err);
  }
}

function setCfg_(k, v) {
  const sh = sh_(H.CFG);
  const col = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getDisplayValues();
  _cfg = null;
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

/* ================= Líderes: horas del equipo y primera aprobación ================= */
// Fechas como texto yyyy-MM-dd; se calculan en UTC para no depender de la zona horaria.
function fUTC_(f) { const p = String(f).split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]); }
function sumarDias_(f, n) { return Utilities.formatDate(new Date(fUTC_(f) + n * 86400000), 'UTC', 'yyyy-MM-dd'); }
function diaSemana_(f) { const d = new Date(fUTC_(f)).getUTCDay(); return d === 0 ? 7 : d; }   // 1 = lunes ... 7 = domingo
function diasLaborales_() { return (cfg_('DIAS_LABORALES') || '1,2,3,4,5').split(/[,;\s]+/).map(x => Number(x.trim())).filter(Boolean); }
function horasDia_() { const h = Number(String(cfg_('HORAS_DIA') || '8').replace(',', '.')); return h > 0 && h <= 24 ? h : 8; }
function esLaborable_(f, dias, feriados) { return dias.indexOf(diaSemana_(f)) >= 0 && feriados.indexOf(f) < 0; }
function r2_(n) { return Math.round(n * 100) / 100; }

// Horas que cubre un permiso aprobado en un día laborable (permiso por horas o de día completo)
function horasPermisoDia_(p, hd) {
  if (p.horaDesde && p.horaHasta) return Math.min(hd, difHoras_(p.horaDesde, p.horaHasta));
  if (p.horaDesde || p.horaHasta) return 0;   // llegada tarde o salida temprana sin hora de fin: no se descuenta
  return hd;
}

// Duración de lo que pide un permiso, en horas laborables (sugerencia para la reposición)
function duracionPermiso_(p) {
  const dias = diasLaborales_(), fer = listaFechas_(cfg_('FERIADOS')), hd = horasDia_();
  let total = 0, f = p.desde, n = 0;
  const hasta = p.hasta || p.desde;
  while (f <= hasta && n++ < 400) {
    if (esLaborable_(f, dias, fer)) total += (p.horaDesde && p.horaHasta) ? Math.min(hd, difHoras_(p.horaDesde, p.horaHasta)) : (p.horaDesde || p.horaHasta) ? 0 : hd;
    f = sumarDias_(f, 1);
  }
  return r2_(total);
}

// Horas trabajadas y esperadas de la semana en curso y del mes, para una lista de correos.
// Trabajadas: horas completas (con salida) desde el lunes (o el día 1) hasta hoy.
// Esperadas: HORAS_DIA por cada día laborable desde el lunes (o el día 1) hasta AYER, sin contar feriados
// ni lo cubierto por permisos aprobados. Un permiso aprobado con reposición no se descuenta: esas horas se deben.
function horasDe_(correos) {
  const hoy = hoy_(), hd = horasDia_(), dias = diasLaborales_(), fer = listaFechas_(cfg_('FERIADOS'));
  const lunes = sumarDias_(hoy, 1 - diaSemana_(hoy)), mes1 = hoy.slice(0, 8) + '01';
  const desde = lunes < mes1 ? lunes : mes1, ayer = sumarDias_(hoy, -1);
  const quiere = {};
  correos.forEach(c => quiere[c] = { semana: { trab: 0, esp: 0 }, mes: { trab: 0, esp: 0 }, hoy: '' });
  // Marcas: se leen solo Fecha..Horas de las últimas filas
  const sh = sh_(H.MAR), last = sh.getLastRow();
  if (last >= 2) {
    const n = Math.min(last - 1, 15000), v = sh.getRange(last - n + 1, 2, n, 6).getDisplayValues();
    v.forEach(r => {
      const x = quiere[correo_(r[1])];
      if (!x || r[0] < desde || r[0] > hoy) return;
      const h = r[4] ? Number(String(r[5]).replace(',', '.')) || 0 : 0;
      if (r[0] >= lunes) x.semana.trab += h;
      if (r[0] >= mes1) x.mes.trab += h;
      if (r[0] === hoy) x.hoy = r[4] ? 'salió' : (r[3] ? 'presente' : '');
    });
  }
  // Permisos aprobados que excusan horas (sin reposición)
  const exc = {};
  rows_(H.PER).forEach(r => {
    const c = correo_(r[2]);
    if (!quiere[c] || r[10] !== 'Aprobado' || r[18] || r[6] < desde || r[5] > ayer) return;
    const p = permisoObj_(r), hp = horasPermisoDia_(p, hd);
    if (!hp) return;
    let f = p.desde < desde ? desde : p.desde, k = 0;
    while (f <= p.hasta && f <= ayer && k++ < 400) { const key = c + '|' + f; exc[key] = Math.min(hd, (exc[key] || 0) + hp); f = sumarDias_(f, 1); }
  });
  let f = desde, k = 0;
  while (f <= ayer && k++ < 62) {
    if (esLaborable_(f, dias, fer)) {
      correos.forEach(c => {
        const esp = Math.max(0, hd - (exc[c + '|' + f] || 0));
        if (f >= lunes) quiere[c].semana.esp += esp;
        if (f >= mes1) quiere[c].mes.esp += esp;
      });
    }
    f = sumarDias_(f, 1);
  }
  Object.keys(quiere).forEach(c => ['semana', 'mes'].forEach(t => {
    const o = quiere[c][t];
    o.trab = r2_(o.trab); o.esp = r2_(o.esp); o.debe = r2_(Math.max(0, o.esp - o.trab)); o.cumple = o.trab >= o.esp;
  }));
  return { porCorreo: quiere, lunes: lunes, mes1: mes1, horasDia: hd };
}

// Lo que ve el líder: su equipo con las horas de la semana y del mes, y los permisos que esperan su decisión
function equipoLider(token) {
  const yo = sesionEmp_(token);
  const equipo = equipoDe_(yo.correo);
  const H_ = horasDe_(equipo.map(e => e.correo));
  const enEquipo = {};
  equipo.forEach(e => enEquipo[e.correo] = true);
  const mios = rows_(H.PER).map(permisoObj_).filter(p => correo_(p.lider) === yo.correo);
  const pendientes = mios.filter(p => p.estado === PEND_LIDER).reverse().map(p => Object.assign(p, {
    horas: (H_.porCorreo[correo_(p.correo)] || horasDe_([correo_(p.correo)]).porCorreo[correo_(p.correo)]),
    duracion: duracionPermiso_(p) }));
  return {
    horasDia: H_.horasDia, lunes: H_.lunes, mes1: H_.mes1, hoy: hoy_(), diasLaborales: diasLaborales_(),
    miembros: equipo.map(e => Object.assign({ correo: e.correo, nombre: e.nombre, cargo: e.cargo }, H_.porCorreo[e.correo]))
      .sort((a, b) => String(a.nombre).localeCompare(String(b.nombre))),
    pendientes: pendientes,
    recientes: mios.filter(p => p.estado !== PEND_LIDER).slice(-20).reverse()
  };
}

// Decisión del líder. Aprobar pasa la solicitud a Administración; rechazar la cierra.
// reponerFecha (opcional): día NO laborable (fin de semana) en que el colaborador repone las horas.
function responderPermisoLider(token, id, decision, nota, reponerFecha, reponerHoras) {
  const yo = sesionEmp_(token);
  if (['Aprobado', 'Rechazado'].indexOf(decision) < 0) throw new Error('Decisión no válida.');
  nota = String(nota || '').trim().slice(0, 300);
  reponerFecha = String(reponerFecha || '').trim();
  let horasR = '';
  if (decision === 'Aprobado' && reponerFecha) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(reponerFecha)) throw new Error('Fecha de reposición no válida.');
    if (reponerFecha <= hoy_()) throw new Error('La reposición debe ser en un día futuro.');
    if (esLaborable_(reponerFecha, diasLaborales_(), [])) throw new Error('La reposición debe ser en un día no laboral (fin de semana).');
    horasR = Number(String(reponerHoras || '').replace(',', '.'));
    if (!(horasR >= 0.5 && horasR <= 12)) throw new Error('Las horas a reponer deben estar entre 0.5 y 12.');
    horasR = r2_(horasR);
  } else reponerFecha = '';
  const res = lock_(() => {
    const sh = sh_(H.PER), last = sh.getLastRow();
    if (last < 2) throw new Error('No se encontró la solicitud.');
    const v = sh.getRange(2, 1, last - 1, PER_COLS.length).getDisplayValues();
    for (let i = v.length - 1; i >= 0; i--) {
      if (v[i][0] !== id) continue;
      const p = permisoObj_(v[i]);
      if (correo_(p.lider) !== yo.correo) throw new Error('Esta solicitud no está a tu cargo como líder.');
      if (correo_(p.correo) === yo.correo) throw new Error('No puedes aprobar tu propio permiso.');
      if (p.estado !== PEND_LIDER) throw new Error('Esta solicitud ya fue respondida (' + p.estado + ').');
      const cuando = ahora_();
      if (decision === 'Aprobado') {
        writeText_(sh, i + 2, [['Pendiente']], 11);
      } else {
        writeText_(sh, i + 2, [['Rechazado', nota, cuando, yo.nombre + ' (líder)']], 11);
      }
      writeText_(sh, i + 2, [[cuando, yo.nombre, nota, reponerFecha, horasR]], 16);
      return p;
    }
    throw new Error('No se encontró la solicitud.');
  });
  // Avisos (fuera del candado)
  const emp = empPorCorreo_(res.correo) || { nombre: res.nombre, correo: res.correo, lugar: '' };
  const rep = reponerFecha ? 'Reponer el ' + reponerFecha + ' (' + horasR + ' h)' : '';
  if (decision === 'Aprobado') {
    pushA_([res.correo], 'Tu líder aprobó tu permiso', 'Ahora falta la aprobación de Administración.' + (rep ? ' ' + rep + '.' : ''), 'permiso');
    avisarPermiso_(emp, res, { titulo: 'Permiso aprobado por el líder: falta la aprobación final', asunto: 'Aprobado por el líder, falta Administración: ',
      filasExtra: [['Aprobado por', yo.nombre + ' (líder)'], ['Nota del líder', nota], ['Reposición', rep]] });
    pushA_(correosAviso_(), 'Permiso listo para aprobación final', emp.nombre + ': aprobado por su líder ' + yo.nombre + '.', 'permiso');
    return 'Aprobado. Pasó a Administración para la aprobación final.' + (rep ? ' ' + rep + '.' : '');
  }
  pushA_([res.correo], 'Tu líder rechazó tu permiso', nota || 'Abre la app para ver el detalle.', 'permiso');
  avisarPermiso_(emp, res, { titulo: 'Permiso rechazado por el líder', asunto: 'Rechazado por el líder: ',
    filasExtra: [['Rechazado por', yo.nombre + ' (líder)'], ['Motivo', nota]], boton: 'Abrir la app' });
  return 'Permiso rechazado.';
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

  // Para decidir: horas de la semana y del mes de quienes tienen permisos por aprobar
  const pend = permisos.filter(p => p.estado === 'Pendiente').reverse();
  if (pend.length) {
    const hs = horasDe_(pend.map(p => correo_(p.correo)).filter((c, i, x) => x.indexOf(c) === i)).porCorreo;
    pend.forEach(p => { p.horas = hs[correo_(p.correo)]; p.duracion = duracionPermiso_(p); });
  }
  return {
    yo: { rol: a.rol, nombre: a.nombre, correo: a.correo }, horasDia: horasDia_(),
    avisos: a.rol === 'Admin' ? cfg_('AVISO_PERMISOS') : '',
    hoy: h, empresa: cfg_('EMPRESA'), url: a.rol === 'Supervisor' ? '' : ss_().getUrl(),
    asistencia: asistencia,
    pendientes: pend,
    esperandoLider: permisos.filter(p => p.estado === PEND_LIDER).reverse(),
    resueltos: permisos.filter(p => p.estado !== 'Pendiente' && p.estado !== PEND_LIDER).slice(-30).reverse(),
    empleados: a.rol === 'Supervisor' ? [] : emps.map(e => {
      const l = liderDe_(e);
      return { correo: e.correo, nombre: e.nombre, cargo: e.cargo, lugar: e.lugar, activo: e.activo,
        tieneClave: !!e.hash, debeCambiar: e.debeCambiar, rol: e.rol, lider: l ? l.correo : '', liderNombre: l ? l.nombre : '' };
    }),
    lugares: lugares_()
  };
}

function responderPermiso(tok, id, estado, respuesta) {
  const a = actor_(tok);
  if (['Aprobado', 'Rechazado'].indexOf(estado) < 0) throw new Error('Estado no válido.');
  const r = lock_(() => {
    const sh = sh_(H.PER);
    const v = sh.getRange(1, 1, sh.getLastRow(), 3).getDisplayValues();
    for (let i = v.length - 1; i >= 1; i--) {
      if (v[i][0] === id) {
        const quien = correo_(v[i][2]);
        if (a.correo && quien === a.correo) throw new Error('No puedes aprobar ni rechazar tu propio permiso.');
        const emp = empPorCorreo_(quien);
        if (!emp || !enAlcance_(a, emp.lugar)) throw new Error('Este empleado no está a tu cargo.');
        const fila = permisoObj_(sh.getRange(i + 1, 1, 1, PER_COLS.length).getDisplayValues()[0]);
        if (fila.estado === PEND_LIDER && a.rol !== 'Admin') throw new Error('Esta solicitud todavía espera la aprobación del líder.');
        if (fila.estado !== 'Pendiente' && fila.estado !== PEND_LIDER) throw new Error('Esta solicitud ya fue respondida (' + fila.estado + ').');
        writeText_(sh, i + 1, [[estado, String(respuesta || '').slice(0, 300), ahora_(), a.nombre + (fila.estado === PEND_LIDER ? ' (sin esperar al líder)' : '')]], 11);
        return { p: fila, emp: emp };
      }
    }
    throw new Error('No se encontró la solicitud.');
  });
  // Avisos: al colaborador y a su líder
  const p = r.p, rep = p.reponerFecha ? ' Reponer el ' + p.reponerFecha + ' (' + p.reponerHoras + ' h).' : '';
  pushA_([p.correo], 'Tu permiso fue ' + estado.toLowerCase(), (respuesta ? String(respuesta).slice(0, 120) : rangoTxt_(p)) + (estado === 'Aprobado' ? rep : ''), 'permiso');
  if (p.lider) pushA_([p.lider], 'Permiso ' + estado.toLowerCase() + ' por Administración', p.nombre + ' (' + rangoTxt_(p) + ').' + (estado === 'Aprobado' ? rep : ''), 'permiso', '#equipo');
  return 'Permiso ' + estado.toLowerCase();
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
  let lider = d.lider === undefined ? null : correo_(d.lider);
  if (lider) {
    if (lider === correo) throw new Error('Una persona no puede ser su propio líder.');
    const l = empPorCorreo_(lider);
    if (!l) throw new Error('El líder elegido no está registrado como empleado.');
    if (l.activo === 'NO') throw new Error('El líder elegido está inactivo.');
  }
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
    if (lider === null) lider = ex ? ex.liderTxt : '';
    const fila = [[correo, String(d.nombre).trim(), String(d.cargo || '').trim(), lugar, d.activo === 'NO' ? 'NO' : 'SI', hash, sal, debe, rol, lider]];
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
  // Lee solo hasta la columna que el reporte necesita (Observación = 15; el resumen solo hasta Horas = 7)
  const marcas = () => {
    const sh = sh_(H.MAR), last = sh.getLastRow();
    if (last < 2) return [];
    const cols = tipo === 'resumen' ? 7 : 15;
    return sh.getRange(2, 1, last - 1, cols).getDisplayValues().filter(r => r[1] >= desde && r[1] <= hasta && emps[correo_(r[2])]);
  };

  if (tipo === 'marcas') {
    return { titulo: titulo, encabezados: ['Fecha', 'Empleado', 'Correo', 'Entrada', 'Salida', 'Horas', 'Lugar', 'Metros entrada', 'Metros salida', 'Observación'],
      filas: marcas().map(r => [r[1], emps[correo_(r[2])].nombre, correo_(r[2]), r[4], r[5], r[6] ? Number(r[6]) : '', r[13], r[9], r[12], r[14] || '']) };
  }
  if (tipo === 'incompletas') {
    return { titulo: titulo, encabezados: ['Fecha', 'Empleado', 'Correo', 'Entrada', 'Lugar'],
      filas: marcas().filter(r => !r[5]).map(r => [r[1], emps[correo_(r[2])].nombre, correo_(r[2]), r[4], r[13]]) };
  }
  if (tipo === 'permisos') {
    return { titulo: titulo, encabezados: ['Solicitado', 'Empleado', 'Correo', 'Tipo', 'Desde', 'Hasta', 'Hora inicio', 'Hora fin', 'Motivo', 'Estado',
        'Líder', 'Líder revisó el', 'Nota del líder', 'Reponer el', 'Horas a reponer', 'Respuesta', 'Revisado por', 'Revisado el'],
      filas: rows_(H.PER).filter(r => r[5] <= hasta && r[6] >= desde && emps[correo_(r[2])])
        .map(r => [r[1], emps[correo_(r[2])].nombre, correo_(r[2]), r[4], r[5], r[6], r[7], r[8], r[9], r[10],
          r[16] || r[14] || '', r[15] || '', r[17] || '', r[18] || '', r[19] ? Number(r[19]) : '', r[11], r[13] || '', r[12]]) };
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
  activarRecordatorios_();
  Logger.log('Recordatorios activos: ' + recordatoriosActivos_() + ' | Firebase configurado: ' + !!props_().getProperty('FCM_CUENTA_SERVICIO'));
  Logger.log('Respaldo mensual activo: ' + respaldoActivo_());
  Logger.log('Hoja de datos: ' + ss_().getUrl());
  Logger.log('Correos restantes hoy: ' + MailApp.getRemainingDailyQuota());
}

/* ================= Recordatorios por notificación (Firebase Cloud Messaging) =================
 * Configuración (una sola vez):
 *  - Propiedad del script FCM_CUENTA_SERVICIO: el JSON completo de la cuenta de servicio de Firebase.
 *  - Hoja Config: DIAS_LABORALES (1 = lunes ... 7 = domingo), AVISOS_ENTRADA, AVISOS_SALIDA (horas HH:mm),
 *    FERIADOS (fechas yyyy-MM-dd separadas por coma).
 * El activador revisarRecordatorios corre cada 5 minutos y envía cada aviso una sola vez al día:
 *  - Avisos de entrada: solo a quien no ha marcado entrada ni tiene un permiso aprobado hoy.
 *  - Aviso de salida: solo a quien marcó entrada y todavía no marca salida.
 */
function registrarDispositivo(token, pushToken, plataforma) {
  const e = sesionEmp_(token);
  pushToken = String(pushToken || '').trim();
  if (pushToken.length < 20 || pushToken.length > 4096) throw new Error('Identificador de notificaciones no válido.');
  return lock_(() => {
    const sh = sh_(H.DIS);
    const v = sh.getDataRange().getDisplayValues();
    const fila = [[e.correo, pushToken, String(plataforma || '').slice(0, 40), ahora_()]];
    for (let i = 1; i < v.length; i++) if (v[i][1] === pushToken) { writeText_(sh, i + 1, fila); return 'Recordatorios activados.'; }
    writeText_(sh, sh.getLastRow() + 1, fila);
    return 'Recordatorios activados.';
  });
}

function quitarDispositivo(token, pushToken) {
  sesionEmp_(token);
  borrarDispositivos_([String(pushToken || '')]);
  return 'Recordatorios desactivados en este teléfono.';
}

function dispositivos_() {
  const mapa = {};
  rows_(H.DIS).forEach(r => { const c = correo_(r[0]); if (c && r[1]) (mapa[c] = mapa[c] || []).push(r[1]); });
  return mapa;
}

function borrarDispositivos_(tokens) {
  if (!tokens.length) return;
  lock_(() => {
    const sh = sh_(H.DIS), last = sh.getLastRow();
    if (last < 2) return;
    const v = sh.getRange(2, 1, last - 1, 4).getDisplayValues();
    const quedan = v.filter(r => tokens.indexOf(r[1]) < 0);
    if (quedan.length === v.length) return;
    sh.getRange(2, 1, last - 1, 4).clearContent();
    if (quedan.length) writeText_(sh, 2, quedan);
  });
}

// Token de acceso a Firebase firmado con la cuenta de servicio (se guarda 50 minutos)
function tokenFcm_() {
  const c = cache_(), guardado = c.get('FCM_ACCESO');
  if (guardado) return guardado;
  let sa;
  try { sa = JSON.parse(props_().getProperty('FCM_CUENTA_SERVICIO') || '{}'); } catch (e) { sa = {}; }
  if (!sa.private_key || !sa.client_email || !sa.project_id) throw new Error('Falta configurar FCM_CUENTA_SERVICIO en las propiedades del script.');
  const b64 = t => Utilities.base64EncodeWebSafe(t, Utilities.Charset.UTF_8).replace(/=+$/, '');
  const ahora = Math.floor(Date.now() / 1000);
  const cab = b64(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const datos = b64(JSON.stringify({ iss: sa.client_email, scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token', iat: ahora, exp: ahora + 3600 }));
  const firma = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(cab + '.' + datos, sa.private_key)).replace(/=+$/, '');
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', { method: 'post', muteHttpExceptions: true,
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: cab + '.' + datos + '.' + firma } });
  const j = JSON.parse(r.getContentText() || '{}');
  if (!j.access_token) throw new Error('Firebase rechazó la cuenta de servicio: ' + (j.error_description || j.error || r.getResponseCode()));
  c.put('FCM_ACCESO', j.access_token, 3000);
  c.put('FCM_PROYECTO', sa.project_id, 3000);
  return j.access_token;
}

// Envía notificaciones: lista de { token, titulo, cuerpo, tipo }. Borra los teléfonos que ya no existen.
function enviarPush_(lista) {
  if (!lista.length) return { enviadas: 0, fallidas: 0 };
  const acceso = tokenFcm_(), proyecto = cache_().get('FCM_PROYECTO');
  let base = cfg_('URL_APP');
  if (base && base.slice(-1) !== '/') base += '/';
  const invalidos = []; let enviadas = 0, fallidas = 0;
  for (let i = 0; i < lista.length; i += 100) {
    const lote = lista.slice(i, i + 100);
    const pedidos = lote.map(m => {
      const webpush = { notification: { tag: 'marcaje-' + m.tipo, renotify: true } };
      if (base) { webpush.fcm_options = { link: base + (m.link || '') }; webpush.notification.icon = base + 'icons/icon-192.png'; }
      return {
        url: 'https://fcm.googleapis.com/v1/projects/' + proyecto + '/messages:send', method: 'post',
        contentType: 'application/json', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + acceso },
        payload: JSON.stringify({ message: { token: m.token, notification: { title: m.titulo, body: m.cuerpo }, webpush: webpush, data: { tipo: m.tipo } } })
      };
    });
    UrlFetchApp.fetchAll(pedidos).forEach((r, k) => {
      const code = r.getResponseCode();
      if (code === 200) { enviadas++; return; }
      fallidas++;
      const t = r.getContentText();
      if (code === 404 || /UNREGISTERED|registration-token-not-registered|INVALID_ARGUMENT/.test(t)) invalidos.push(lote[k].token);
    });
  }
  borrarDispositivos_(invalidos);
  return { enviadas: enviadas, fallidas: fallidas };
}

function minutos_(hhmm) { const p = String(hhmm).split(':').map(Number); return p[0] * 60 + (p[1] || 0); }
// Acepta 08:00, 8:00, 8:00:00, 8:00 a. m., 5:00 p. m. (Sheets a veces convierte lo que se escribe a mano)
function listaHoras_(t) {
  return String(t || '').split(/[,;]/).map(x => x.trim().toLowerCase()).map(x => {
    const m = x.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([ap])?/);
    if (!m) return '';
    let h = Number(m[1]);
    if (m[3] === 'p' && h < 12) h += 12;
    if (m[3] === 'a' && h === 12) h = 0;
    return (h < 10 ? '0' : '') + h + ':' + m[2];
  }).filter(Boolean).sort();
}
// Acepta 2026-12-25, 25/12/2026 o 25-12-2026
function listaFechas_(t) {
  return String(t || '').split(/[,;\s]+/).map(x => x.trim()).map(x => {
    let m = x.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
    m = x.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
    if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
    return '';
  }).filter(Boolean);
}

// Lo ejecuta el activador cada 5 minutos
function revisarRecordatorios() {
  if (!props_().getProperty('FCM_CUENTA_SERVICIO')) return;
  const ahora = new Date(), hoy = hoy_();
  const dia = Number(Utilities.formatDate(ahora, tz_(), 'u'));               // 1 = lunes ... 7 = domingo
  const dias = (cfg_('DIAS_LABORALES') || '1,2,3,4,5').split(/[,;\s]+/).map(x => Number(x.trim()));
  // En días no laborales solo se avisa a quien tiene una reposición aprobada para hoy
  let solo = null;
  if (dias.indexOf(dia) < 0 || listaFechas_(cfg_('FERIADOS')).indexOf(hoy) >= 0) {
    solo = {};
    rows_(H.PER).forEach(r => { if (r[10] === 'Aprobado' && r[18] === hoy) solo[correo_(r[2])] = true; });
    if (!Object.keys(solo).length) return;
  }
  const hm = minutos_(Utilities.formatDate(ahora, tz_(), 'HH:mm'));
  const avisos = listaHoras_(cfg_('AVISOS_ENTRADA') || '08:00,08:15,08:30').map((t, i) => ({ t: t, tipo: 'entrada', n: i }))
    .concat(listaHoras_(cfg_('AVISOS_SALIDA') || '17:00').map((t, i) => ({ t: t, tipo: 'salida', n: i })));
  const p = props_();
  avisos.forEach(a => {
    const k = 'AVISO_' + hoy + '_' + a.tipo + '_' + a.t, dif = hm - minutos_(a.t);
    if (dif < 0 || dif > 25 || p.getProperty(k)) return;   // a su hora (hasta 25 min de margen) y una sola vez
    p.setProperty(k, ahora_());
    try { enviarAviso_(a, hoy, solo); } catch (e) { console.error('Recordatorio ' + k + ': ' + e); }
  });
  // Limpia las marcas de avisos de días anteriores
  const todas = p.getProperties();
  Object.keys(todas).forEach(k => { if (k.indexOf('AVISO_') === 0 && k.indexOf('AVISO_' + hoy) !== 0) p.deleteProperty(k); });
}

function enviarAviso_(a, hoy, solo) {
  const dev = dispositivos_();
  const marcas = {};
  marcasDeHoy_().forEach(r => { marcas[correo_(r[2])] = r; });
  const permiso = {};
  rows_(H.PER).forEach(r => { if (r[10] === 'Aprobado' && r[5] <= hoy && r[6] >= hoy) permiso[correo_(r[2])] = true; });
  const lista = [];
  empleados_().filter(e => e.activo === 'SI' && dev[e.correo] && (!solo || solo[e.correo])).forEach(e => {
    const m = marcas[e.correo], nombre = String(e.nombre || '').trim().split(' ')[0] || '';
    let titulo, cuerpo;
    if (a.tipo === 'entrada' && solo) {
      if (m || a.n > 0) return;
      titulo = 'Hoy repones horas'; cuerpo = 'Buenos días, ' + nombre + '. Hoy tienes una reposición aprobada: marca tu entrada al llegar.';
    } else if (a.tipo === 'entrada') {
      if (m || permiso[e.correo]) return;
      if (a.n === 0) { titulo = 'Recuerda marcar tu entrada'; cuerpo = 'Buenos días, ' + nombre + '. Marca tu entrada al llegar.'; }
      else if (a.n === 1) { titulo = 'Ya pasó la tolerancia de entrada'; cuerpo = nombre + ', todavía no has marcado tu entrada. Márcala ahora.'; }
      else { titulo = 'Todavía no has marcado tu entrada'; cuerpo = nombre + ', marca tu entrada apenas llegues. Si hoy no vienes, pide un permiso en la app.'; }
    } else {
      if (!m || !m[4] || m[5]) return;
      titulo = 'No olvides marcar tu salida'; cuerpo = 'Terminó la jornada, ' + nombre + '. Cuando salgas, marca tu salida.';
    }
    dev[e.correo].forEach(t => lista.push({ token: t, titulo: titulo, cuerpo: cuerpo, tipo: a.tipo }));
  });
  const r = enviarPush_(lista);
  console.log('Aviso ' + a.tipo + ' ' + a.t + ': ' + r.enviadas + ' enviadas, ' + r.fallidas + ' fallidas');
  return r;
}

function probarNotificacion(token) {
  const e = sesionEmp_(token);
  const t = dispositivos_()[e.correo] || [];
  if (!t.length) throw new Error('Este usuario no tiene recordatorios activados en ningún teléfono.');
  const r = enviarPush_(t.map(x => ({ token: x, titulo: 'Recordatorios activados', cuerpo: 'Así te llegará el aviso si se te olvida marcar.', tipo: 'prueba' })));
  if (!r.enviadas) throw new Error('No se pudo enviar la prueba. Vuelve a activar los recordatorios.');
  return 'Notificación de prueba enviada.';
}

function recordatoriosActivos_() {
  return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'revisarRecordatorios');
}
function activarRecordatorios_() {
  if (!recordatoriosActivos_()) ScriptApp.newTrigger('revisarRecordatorios').timeBased().everyMinutes(5).create();
}

function estadoRecordatorios(tok) {
  actor_(tok, ['Admin']);
  const dev = dispositivos_();
  return { firebase: !!props_().getProperty('FCM_CUENTA_SERVICIO'), activo: recordatoriosActivos_(),
    personas: Object.keys(dev).length, telefonos: Object.keys(dev).reduce((s, k) => s + dev[k].length, 0),
    dias: cfg_('DIAS_LABORALES') || '1,2,3,4,5', entrada: cfg_('AVISOS_ENTRADA') || '08:00,08:15,08:30', salida: cfg_('AVISOS_SALIDA') || '17:00' };
}

function activarRecordatorios(tok) {
  actor_(tok, ['Admin']);
  activarRecordatorios_();
  return 'Recordatorios automáticos activados (se revisan cada 5 minutos).';
}
