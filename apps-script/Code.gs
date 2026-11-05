/**
 * App Trinidad de las Huertas — API (Google Apps Script)
 * Comité Vecinal del Fraccionamiento Trinidad de las Huertas, Oaxaca de Juárez.
 *
 * Implementar como "Aplicación web": Ejecutar como = Yo, Quién tiene acceso = Cualquier usuario.
 * La primera vez ejecuta la función instalar() desde el editor para crear la hoja y la carpeta de fotos.
 * La clave del comité se guarda en Propiedades del script (CLAVE); nunca en este archivo.
 */

const PROPS = PropertiesService.getScriptProperties();
const H_USUARIOS = ['usuario', 'nombre', 'calle', 'casa', 'telefono', 'rol', 'estado', 'salt', 'hash', 'creado', 'ultimoAcceso'];
const H_ITEMS = ['id', 'categoria', 'num', 'codigo', 'falla', 'estado', 'ubicacion', 'entre', 'lat', 'lon', 'foto', 'actualizado', 'por', 'data'];
const H_REPORTES = ['rid', 'fecha', 'categoria', 'itemId', 'itemCodigo', 'tipo', 'comentario', 'usuario', 'nombre', 'casa', 'telefono',
  'lat', 'lon', 'foto', 'estado', 'fechaEstado', 'por', 'notas', 'cuadrilla', 'apoyos', 'historial'];
const H_BITACORA = ['fecha', 'usuario', 'nombre', 'accion', 'detalle'];
const H_LUGARES = ['id', 'clase', 'nombre', 'giro', 'estado', 'dueno', 'lat', 'lon', 'foto', 'actualizado', 'por', 'data'];
const H_PEDIDOS = ['pid', 'fecha', 'lugarId', 'lugarNombre', 'cliente', 'nombre', 'casa', 'telefono', 'items', 'total', 'entrega', 'nota',
  'estado', 'fechaEstado', 'dlat', 'dlon', 'historial'];
const ESTADOS_PED = ['recibido', 'preparando', 'camino', 'entregado', 'cancelado'];
const H_CONTENIDO = ['clave', 'data', 'actualizado', 'por'];
const CONTENIDOS = ['comite', 'reglamentos', 'numeros'];
const MIMES = { 'application/pdf': 'pdf', 'image/jpeg': 'jpg', 'image/png': 'png' };
const CATEGORIAS = ['luminarias', 'drenaje', 'baches', 'fugas', 'apagones', 'basura', 'poda'];
const INVENTARIO = ['luminarias', 'drenaje'];
const ROLES = ['vecino', 'operador', 'comite'];
const ESTADOS_REP = ['nuevo', 'proceso', 'atendido', 'descartado'];
const DIAS_TOKEN = 180;

// ---------------- Infraestructura ----------------
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss_() {
  const id = PROPS.getProperty('SS_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { } }
  const s = SpreadsheetApp.create('App Trinidad de las Huertas · Datos');
  PROPS.setProperty('SS_ID', s.getId());
  return s;
}
function hoja_(nombre, headers) {
  const s = ss_(); let sh = s.getSheetByName(nombre);
  if (!sh) {
    sh = s.insertSheet(nombre);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#7a1f2b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    const def = s.getSheetByName('Hoja 1') || s.getSheetByName('Sheet1');
    if (def && s.getSheets().length > 1 && def.getLastRow() === 0) s.deleteSheet(def);
  }
  return sh;
}
const hUsuarios_ = () => hoja_('Usuarios', H_USUARIOS);
const hItems_ = () => hoja_('Inventario', H_ITEMS);
const hReportes_ = () => hoja_('Reportes', H_REPORTES);
const hBitacora_ = () => hoja_('Bitácora', H_BITACORA);
const hLugares_ = () => hoja_('Lugares', H_LUGARES);
const hPedidos_ = () => hoja_('Pedidos', H_PEDIDOS);
const hContenido_ = () => hoja_('Contenido', H_CONTENIDO);
function contenido_() {
  const o = {}; filas_(hContenido_(), H_CONTENIDO).forEach(r => { try { o[r.clave] = JSON.parse(r.data || 'null'); } catch (e) { } });
  return o;
}

function filas_(sh, headers) {
  const n = sh.getLastRow() - 1; if (n <= 0) return [];
  const v = sh.getRange(2, 1, n, headers.length).getValues();
  return v.map((r, i) => { const o = { _row: i + 2 }; headers.forEach((h, j) => o[h] = r[j]); return o; });
}
function escribirFila_(sh, headers, row, obj) {
  const vals = headers.map(h => obj[h] === undefined || obj[h] === null ? '' : obj[h]);
  if (row) sh.getRange(row, 1, 1, headers.length).setValues([vals]);
  else sh.appendRow(vals);
}
function carpetaFotos_() {
  const id = PROPS.getProperty('FOTOS_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  const f = DriveApp.createFolder('App Trinidad de las Huertas · Fotos');
  PROPS.setProperty('FOTOS_ID', f.getId());
  return f;
}
function secreto_() {
  let s = PROPS.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); PROPS.setProperty('SECRET', s); }
  return s;
}
function hex_(bytes) { return bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }
function hashPw_(salt, pw) {
  let h = salt + '|' + pw;
  for (let i = 0; i < 300; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  return h;
}
function b64_(s) { return Utilities.base64EncodeWebSafe(s, Utilities.Charset.UTF_8).replace(/=+$/, ''); }
function unb64_(s) { return Utilities.newBlob(Utilities.base64DecodeWebSafe(s + '==='.slice((s.length + 3) % 4))).getDataAsString('UTF-8'); }
function firma_(txt) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(txt, secreto_())).replace(/=+$/, ''); }
function token_(usuario) {
  const exp = Date.now() + DIAS_TOKEN * 864e5;
  const p = usuario + '|' + exp + '|' + (PROPS.getProperty('VER_' + usuario) || '0');
  return b64_(p) + '.' + firma_(p);
}
function limpiar_(s, n) { return String(s === undefined || s === null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n || 120); }
function normUsuario_(u) { return String(u || '').trim().toLowerCase(); }
function ahora_() { return new Date().toISOString(); }
function nuevoId_(pref) { return pref + Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36); }

function usuarios_() { return filas_(hUsuarios_(), H_USUARIOS); }
function buscarUsuario_(u) { u = normUsuario_(u); return usuarios_().find(x => normUsuario_(x.usuario) === u) || null; }
function publicoUsuario_(x) { return { usuario: x.usuario, nombre: x.nombre, calle: x.calle, casa: String(x.casa), telefono: String(x.telefono || ''), rol: x.rol, estado: x.estado }; }

function verificar_(token) {
  const parts = String(token || '').split('.'); if (parts.length !== 2) return null;
  let p; try { p = unb64_(parts[0]); } catch (e) { return null; }
  if (firma_(p) !== parts[1]) return null;
  const [usuario, exp, ver] = p.split('|');
  if (Date.now() > Number(exp)) return null;
  if ((PROPS.getProperty('VER_' + usuario) || '0') !== ver) return null;
  const u = buscarUsuario_(usuario);
  if (!u || u.estado === 'bloqueado') return null;
  return u;
}
function bitacora_(u, accion, detalle) {
  try { hBitacora_().appendRow([ahora_(), u ? u.usuario : '', u ? u.nombre : '', accion, limpiar_(detalle, 400)]); } catch (e) { }
}
function guardarFoto_(b64, nombre) {
  if (!b64 || b64.length > 2.8e6) throw new Error('foto');
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', nombre + '.jpg');
  const f = carpetaFotos_().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f.getId();
}
function intentos_(clave, sumar) {
  const c = CacheService.getScriptCache(); const k = 'int_' + clave; let n = Number(c.get(k) || 0);
  if (sumar) { n++; c.put(k, String(n), 900); }
  return n;
}

// ---------------- Lectura ----------------
function itemDe_(r) {
  let d = {}; try { d = JSON.parse(r.data || '{}'); } catch (e) { }
  return Object.assign(d, { id: String(r.id), categoria: r.categoria, num: Number(r.num) || null, lat: Number(r.lat), lon: Number(r.lon), foto: r.foto ? String(r.foto) : (d.foto || '') });
}
function reporteDe_(r) {
  let hist = [], apoyos = [];
  try { hist = JSON.parse(r.historial || '[]'); } catch (e) { }
  try { apoyos = JSON.parse(r.apoyos || '[]'); } catch (e) { }
  return {
    rid: String(r.rid), fecha: r.fecha instanceof Date ? r.fecha.toISOString() : String(r.fecha), categoria: r.categoria, itemId: String(r.itemId || ''),
    itemCodigo: String(r.itemCodigo || ''), tipo: r.tipo, comentario: r.comentario, usuario: String(r.usuario), nombre: r.nombre, casa: String(r.casa),
    telefono: String(r.telefono || ''), lat: Number(r.lat), lon: Number(r.lon), foto: String(r.foto || ''), estado: r.estado || 'nuevo',
    fechaEstado: r.fechaEstado instanceof Date ? r.fechaEstado.toISOString() : String(r.fechaEstado || ''), por: String(r.por || ''), notas: String(r.notas || ''),
    cuadrilla: String(r.cuadrilla || ''), apoyos: apoyos, historial: hist, _row: r._row
  };
}
function datos_(u) {
  const staff = u.rol === 'comite' || u.rol === 'operador';
  const items = filas_(hItems_(), H_ITEMS).map(itemDe_);
  const hace30 = Date.now() - 30 * 864e5;
  const reps = filas_(hReportes_(), H_REPORTES).map(reporteDe_);
  const out = [];
  reps.forEach(r => {
    const mio = normUsuario_(r.usuario) === normUsuario_(u.usuario);
    const abierto = r.estado === 'nuevo' || r.estado === 'proceso';
    const reciente = new Date(r.fechaEstado || r.fecha).getTime() > hace30;
    if (!(staff || mio || abierto || (r.estado === 'atendido' && reciente))) return;
    if (staff && !abierto && !mio && new Date(r.fechaEstado || r.fecha).getTime() < Date.now() - 400 * 864e5) return;
    const o = Object.assign({}, r); delete o._row;
    o.mio = mio; o.apoyo = o.apoyos.indexOf(u.usuario) >= 0; o.nApoyos = o.apoyos.length;
    if (!staff) {
      delete o.apoyos; delete o.telefono; delete o.notas; delete o.por;
      if (!mio) { delete o.usuario; delete o.nombre; delete o.casa; o.historial = []; }
    }
    out.push(o);
  });
  const lugares = filas_(hLugares_(), H_LUGARES).map(lugarDe_)
    .filter(l => l.estado === 'aprobado' || u.rol === 'comite' || normUsuario_(l.dueno) === normUsuario_(u.usuario))
    .map(l => { l.mio = normUsuario_(l.dueno) === normUsuario_(u.usuario); if (u.rol !== 'comite' && !l.mio) delete l.dueno; return l; });
  const mios = {}; lugares.forEach(l => { if (l.mio) mios[l.id] = 1; });
  const hace60 = Date.now() - 60 * 864e5;
  const pedidos = filas_(hPedidos_(), H_PEDIDOS).map(pedidoDe_).filter(p => {
    const soyCliente = normUsuario_(p.cliente) === normUsuario_(u.usuario), soyNegocio = !!mios[p.lugarId];
    if (!(soyCliente || soyNegocio)) return false;
    const abierto = ['recibido', 'preparando', 'camino'].indexOf(p.estado) >= 0;
    return abierto || new Date(p.fecha).getTime() > hace60;
  }).map(p => { p.soyCliente = normUsuario_(p.cliente) === normUsuario_(u.usuario); p.soyNegocio = !!mios[p.lugarId]; delete p._row; return p; });
  return { items: items, reportes: out, lugares: lugares, pedidos: pedidos, contenido: contenido_() };
}
function lugarDe_(r) {
  let d = {}; try { d = JSON.parse(r.data || '{}'); } catch (e) { }
  return Object.assign(d, { id: String(r.id), clase: r.clase, nombre: String(r.nombre), giro: String(r.giro || ''), estado: r.estado || 'aprobado', dueno: String(r.dueno || ''),
    lat: Number(r.lat), lon: Number(r.lon), foto: String(r.foto || ''), actualizado: String(r.actualizado || ''), _row: r._row });
}
function pedidoDe_(r) {
  let items = [], hist = []; try { items = JSON.parse(r.items || '[]'); } catch (e) { } try { hist = JSON.parse(r.historial || '[]'); } catch (e) { }
  const iso = v => v instanceof Date ? v.toISOString() : String(v || '');
  return { pid: String(r.pid), fecha: iso(r.fecha), lugarId: String(r.lugarId), lugarNombre: String(r.lugarNombre), cliente: String(r.cliente), nombre: String(r.nombre),
    casa: String(r.casa), telefono: String(r.telefono || ''), items: items, total: Number(r.total) || 0, entrega: String(r.entrega), nota: String(r.nota || ''),
    estado: String(r.estado), fechaEstado: iso(r.fechaEstado), dlat: r.dlat === '' ? null : Number(r.dlat), dlon: r.dlon === '' ? null : Number(r.dlon), historial: hist, _row: r._row };
}
const enZona_ = (lat, lon) => lat > 16.9 && lat < 17.2 && lon > -96.9 && lon < -96.5;
function limpiarMenu_(m) {
  return (Array.isArray(m) ? m : []).slice(0, 80).map(x => ({ n: limpiarTxt_(x && x.n, 60), p: Math.max(0, Math.min(99999, Number(x && x.p) || 0)), d: limpiarTxt_(x && x.d, 120) })).filter(x => x.n);
}
function limpiarTxt_(s, n) { return limpiar_(s, n); }

// ---------------- Entrada ----------------
function doGet() { return json_({ ok: true, app: 'Trinidad de las Huertas', v: 1 }); }

function doPost(e) {
  let req; try { req = JSON.parse(e.postData.contents); } catch (er) { return json_({ ok: false, error: 'json' }); }
  try { return json_(atender_(req)); }
  catch (er) { return json_({ ok: false, error: String(er.message || er).slice(0, 80) }); }
}

function atender_(req) {
  const a = req.action;

  // ---- Cuentas ----
  if (a === 'registro') {
    const usuario = normUsuario_(req.usuario);
    if (!/^[a-z0-9._-]{3,30}$/.test(usuario)) return { ok: false, error: 'usuario_invalido' };
    const pw = String(req.password || ''); if (pw.length < 6) return { ok: false, error: 'pw_corta' };
    const nombre = limpiar_(req.nombre, 80), calle = limpiar_(req.calle, 80), casa = limpiar_(req.casa, 20);
    if (nombre.length < 3 || !calle || !casa) return { ok: false, error: 'faltan_datos' };
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      if (buscarUsuario_(usuario)) return { ok: false, error: 'usuario_existe' };
      if (usuarios_().length > 3000) return { ok: false, error: 'lleno' };
      const salt = Utilities.getUuid();
      const u = { usuario: usuario, nombre: nombre, calle: calle, casa: casa, telefono: limpiar_(req.telefono, 20), rol: 'vecino', estado: 'activo', salt: salt, hash: hashPw_(salt, pw), creado: ahora_(), ultimoAcceso: ahora_() };
      escribirFila_(hUsuarios_(), H_USUARIOS, null, u);
      bitacora_(u, 'Registro', `${nombre} · ${calle} ${casa}`);
      return { ok: true, token: token_(usuario), user: publicoUsuario_(u) };
    } finally { lock.releaseLock(); }
  }
  if (a === 'login') {
    const usuario = normUsuario_(req.usuario);
    if (intentos_(usuario) >= 8) return { ok: false, error: 'bloqueo_temporal' };
    const u = buscarUsuario_(usuario);
    if (!u || hashPw_(u.salt, String(req.password || '')) !== u.hash) { intentos_(usuario, true); return { ok: false, error: 'credenciales' }; }
    if (u.estado === 'bloqueado') return { ok: false, error: 'bloqueado' };
    hUsuarios_().getRange(u._row, H_USUARIOS.indexOf('ultimoAcceso') + 1).setValue(ahora_());
    return { ok: true, token: token_(u.usuario), user: publicoUsuario_(u) };
  }

  const u = verificar_(req.token);
  if (!u) return { ok: false, error: 'sesion' };
  const comite = u.rol === 'comite', staff = comite || u.rol === 'operador';

  if (a === 'yo') return { ok: true, user: publicoUsuario_(u) };
  if (a === 'datos') return Object.assign({ ok: true, user: publicoUsuario_(u) }, datos_(u));

  if (a === 'cambiarPassword') {
    if (hashPw_(u.salt, String(req.actual || '')) !== u.hash) return { ok: false, error: 'credenciales' };
    const pw = String(req.nueva || ''); if (pw.length < 6) return { ok: false, error: 'pw_corta' };
    const salt = Utilities.getUuid(); const sh = hUsuarios_();
    sh.getRange(u._row, H_USUARIOS.indexOf('salt') + 1, 1, 2).setValues([[salt, hashPw_(salt, pw)]]);
    bitacora_(u, 'Cambió su contraseña', '');
    return { ok: true };
  }
  if (a === 'perfil') {
    const nombre = limpiar_(req.nombre, 80), calle = limpiar_(req.calle, 80), casa = limpiar_(req.casa, 20);
    if (nombre.length < 3 || !calle || !casa) return { ok: false, error: 'faltan_datos' };
    hUsuarios_().getRange(u._row, 2, 1, 4).setValues([[nombre, calle, casa, limpiar_(req.telefono, 20)]]);
    bitacora_(u, 'Actualizó su perfil', `${nombre} · ${calle} ${casa}`);
    return { ok: true, user: publicoUsuario_(Object.assign(u, { nombre: nombre, calle: calle, casa: casa, telefono: limpiar_(req.telefono, 20) })) };
  }
  if (a === 'hacermeComite') {
    if (intentos_('comite_' + u.usuario) >= 6) return { ok: false, error: 'bloqueo_temporal' };
    const clave = PROPS.getProperty('CLAVE');
    if (!clave || String(req.clave || '') !== clave) { intentos_('comite_' + u.usuario, true); return { ok: false, error: 'clave' }; }
    hUsuarios_().getRange(u._row, H_USUARIOS.indexOf('rol') + 1).setValue('comite');
    bitacora_(u, 'Entró como comité con la clave del comité', '');
    return { ok: true, user: publicoUsuario_(Object.assign(u, { rol: 'comite' })) };
  }

  // ---- Reportes de vecinos ----
  if (a === 'reportar') {
    const cat = String(req.categoria); if (CATEGORIAS.indexOf(cat) < 0) return { ok: false, error: 'categoria' };
    const tipo = limpiar_(req.tipo, 60); if (!tipo) return { ok: false, error: 'tipo' };
    const lat = Number(req.lat), lon = Number(req.lon);
    if (!(lat > 16.9 && lat < 17.2 && lon > -96.9 && lon < -96.5)) return { ok: false, error: 'ubicacion' };
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const sh = hReportes_();
      if (sh.getLastRow() > 20000) return { ok: false, error: 'lleno' };
      const abiertosMios = filas_(sh, H_REPORTES).filter(r => normUsuario_(r.usuario) === normUsuario_(u.usuario) && new Date(r.fecha).getTime() > Date.now() - 864e5).length;
      if (abiertosMios >= 25) return { ok: false, error: 'demasiados' };
      const rid = nuevoId_('r');
      const foto = req.foto ? guardarFoto_(req.foto, cat + '_' + rid) : '';
      const r = {
        rid: rid, fecha: ahora_(), categoria: cat, itemId: limpiar_(req.itemId, 40), itemCodigo: limpiar_(req.itemCodigo, 40), tipo: tipo,
        comentario: limpiar_(req.comentario, 400), usuario: u.usuario, nombre: u.nombre, casa: `${u.calle} ${u.casa}`, telefono: u.telefono,
        lat: lat, lon: lon, foto: foto, estado: 'nuevo', fechaEstado: ahora_(), por: '', notas: '', cuadrilla: '', apoyos: '[]',
        historial: JSON.stringify([{ f: ahora_(), t: 'Reporte recibido', u: u.nombre }])
      };
      escribirFila_(sh, H_REPORTES, null, r);
      bitacora_(u, 'Reportó ' + cat, `${tipo}${r.itemCodigo ? ' · ' + r.itemCodigo : ''} · ${rid}`);
      return { ok: true, rid: rid, foto: foto };
    } finally { lock.releaseLock(); }
  }
  if (a === 'apoyar') {
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const sh = hReportes_(); const r = filas_(sh, H_REPORTES).map(reporteDe_).find(x => x.rid === String(req.rid));
      if (!r) return { ok: false, error: 'no_existe' };
      const i = r.apoyos.indexOf(u.usuario); if (i >= 0) r.apoyos.splice(i, 1); else r.apoyos.push(u.usuario);
      sh.getRange(r._row, H_REPORTES.indexOf('apoyos') + 1).setValue(JSON.stringify(r.apoyos));
      if (i < 0) bitacora_(u, 'Confirmó un reporte', `${r.categoria} · ${r.tipo} · ${r.rid}`);
      return { ok: true, apoyo: i < 0, nApoyos: r.apoyos.length };
    } finally { lock.releaseLock(); }
  }
  if (a === 'cancelarReporte') {
    const sh = hReportes_(); const r = filas_(sh, H_REPORTES).map(reporteDe_).find(x => x.rid === String(req.rid));
    if (!r || normUsuario_(r.usuario) !== normUsuario_(u.usuario) || r.estado !== 'nuevo') return { ok: false, error: 'permiso' };
    r.historial.push({ f: ahora_(), t: 'Cancelado por quien reportó', u: u.nombre });
    sh.getRange(r._row, H_REPORTES.indexOf('estado') + 1, 1, 3).setValues([['descartado', ahora_(), u.usuario]]);
    sh.getRange(r._row, H_REPORTES.indexOf('historial') + 1).setValue(JSON.stringify(r.historial.slice(-60)));
    bitacora_(u, 'Canceló su reporte', r.rid);
    return { ok: true };
  }

  // ---- Negocios y lugares ----
  if (a === 'lugarGuardar') {
    const d = req.data || {}; const clase = req.clase === 'clave' ? 'clave' : 'negocio';
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const sh = hLugares_(); const filas = filas_(sh, H_LUGARES).map(lugarDe_);
      const prev = req.id ? filas.find(x => x.id === String(req.id)) : null;
      if (req.id && !prev) return { ok: false, error: 'no_existe' };
      const dueno = prev ? prev.dueno : u.usuario;
      const esDueno = normUsuario_(dueno) === normUsuario_(u.usuario);
      if ((prev ? prev.clase : clase) === 'clave' && !comite) return { ok: false, error: 'permiso' };
      if (prev && !esDueno && !comite) return { ok: false, error: 'permiso' };
      if (!prev && filas.filter(x => normUsuario_(x.dueno) === normUsuario_(u.usuario)).length >= 10 && !comite) return { ok: false, error: 'demasiados' };
      const base = prev ? Object.assign({}, prev) : {};
      delete base._row; delete base.mio;
      const next = Object.assign(base, {
        nombre: limpiar_(d.nombre !== undefined ? d.nombre : base.nombre, 80), giro: limpiar_(d.giro !== undefined ? d.giro : base.giro, 30),
        descripcion: limpiar_(d.descripcion !== undefined ? d.descripcion : base.descripcion, 600), horario: limpiar_(d.horario !== undefined ? d.horario : base.horario, 120),
        telefono: limpiar_(d.telefono !== undefined ? d.telefono : base.telefono, 20), whatsapp: limpiar_(d.whatsapp !== undefined ? d.whatsapp : base.whatsapp, 20),
        direccion: limpiar_(d.direccion !== undefined ? d.direccion : base.direccion, 160), domicilio: d.domicilio !== undefined ? !!d.domicilio : !!base.domicilio,
        maps: limpiar_(d.maps !== undefined ? d.maps : base.maps, 300)
      });
      if (d.menu !== undefined) next.menu = limpiarMenu_(d.menu);
      if (d.lat !== undefined) { next.lat = Number(d.lat); next.lon = Number(d.lon); }
      if (!next.nombre || next.nombre.length < 2) return { ok: false, error: 'faltan_datos' };
      if (!enZona_(Number(next.lat), Number(next.lon))) return { ok: false, error: 'ubicacion' };
      const id = prev ? prev.id : nuevoId_(clase === 'clave' ? 'k' : 'n');
      const estado = prev ? prev.estado : (clase === 'clave' || comite ? 'aprobado' : 'pendiente');
      const row = { id: id, clase: prev ? prev.clase : clase, nombre: next.nombre, giro: next.giro, estado: estado, dueno: dueno, lat: Number(next.lat), lon: Number(next.lon),
        foto: next.foto || '', actualizado: ahora_(), por: u.usuario, data: JSON.stringify(next).slice(0, 45000) };
      escribirFila_(sh, H_LUGARES, prev ? prev._row : null, row);
      bitacora_(u, (prev ? 'Editó ' : 'Registró ') + (row.clase === 'clave' ? 'lugar clave' : 'negocio'), `${row.nombre}${!prev && estado === 'pendiente' ? ' · por aprobar' : ''}`);
      return { ok: true, lugar: Object.assign(next, { id: id, clase: row.clase, estado: estado, dueno: dueno, mio: normUsuario_(dueno) === normUsuario_(u.usuario), actualizado: row.actualizado }) };
    } finally { lock.releaseLock(); }
  }
  if (a === 'lugarFoto') {
    const sh = hLugares_(); const l = filas_(sh, H_LUGARES).map(lugarDe_).find(x => x.id === String(req.id));
    if (!l) return { ok: false, error: 'no_existe' };
    if (!comite && normUsuario_(l.dueno) !== normUsuario_(u.usuario)) return { ok: false, error: 'permiso' };
    const foto = guardarFoto_(req.data, 'lugar_' + l.id + '_' + Date.now());
    const campo = req.campo === 'menuFoto' ? 'menuFoto' : 'foto';
    const d = Object.assign({}, l); delete d._row; d[campo] = foto;
    if (campo === 'foto') sh.getRange(l._row, H_LUGARES.indexOf('foto') + 1).setValue(foto);
    sh.getRange(l._row, H_LUGARES.indexOf('data') + 1).setValue(JSON.stringify(d).slice(0, 45000));
    bitacora_(u, campo === 'foto' ? 'Foto de lugar' : 'Foto de menú', l.nombre);
    return { ok: true, foto: foto, campo: campo };
  }
  if (a === 'lugarEstado') {
    if (!comite) return { ok: false, error: 'permiso' };
    const est = String(req.estado); if (['aprobado', 'pendiente', 'rechazado'].indexOf(est) < 0) return { ok: false, error: 'estado' };
    const sh = hLugares_(); const l = filas_(sh, H_LUGARES).map(lugarDe_).find(x => x.id === String(req.id));
    if (!l) return { ok: false, error: 'no_existe' };
    sh.getRange(l._row, H_LUGARES.indexOf('estado') + 1).setValue(est);
    bitacora_(u, est === 'aprobado' ? 'Aprobó negocio' : est === 'rechazado' ? 'Rechazó negocio' : 'Regresó negocio a revisión', l.nombre);
    return { ok: true };
  }
  if (a === 'lugarBorrar') {
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const sh = hLugares_(); const l = filas_(sh, H_LUGARES).map(lugarDe_).find(x => x.id === String(req.id));
      if (!l) return { ok: false, error: 'no_existe' };
      if (!comite && normUsuario_(l.dueno) !== normUsuario_(u.usuario)) return { ok: false, error: 'permiso' };
      sh.deleteRow(l._row);
      bitacora_(u, 'Eliminó ' + (l.clase === 'clave' ? 'lugar clave' : 'negocio'), l.nombre);
      return { ok: true };
    } finally { lock.releaseLock(); }
  }

  // ---- Pedidos ----
  if (a === 'pedidoCrear') {
    const l = filas_(hLugares_(), H_LUGARES).map(lugarDe_).find(x => x.id === String(req.lugarId));
    if (!l || l.clase !== 'negocio' || l.estado !== 'aprobado') return { ok: false, error: 'no_existe' };
    const menu = limpiarMenu_(l.menu); const precio = {}; menu.forEach(m => precio[m.n] = m.p);
    const items = (Array.isArray(req.items) ? req.items : []).slice(0, 60).map(x => {
      const n = limpiar_(x && x.n, 80), q = Math.max(1, Math.min(99, Math.round(Number(x && x.q) || 1)));
      return { n: n, q: q, p: precio[n] !== undefined ? precio[n] : 0 };
    }).filter(x => x.n);
    const nota = limpiar_(req.nota, 500);
    if (!items.length && !nota) return { ok: false, error: 'faltan_datos' };
    const total = items.reduce((s, x) => s + x.p * x.q, 0);
    const entrega = req.entrega === 'domicilio' && l.domicilio ? 'domicilio' : 'recoger';
    const dl = Number(req.dlat), dn = Number(req.dlon); const conDest = entrega === 'domicilio' && enZona_(dl, dn);
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const sh = hPedidos_();
      const hoy = filas_(sh, H_PEDIDOS).filter(p => normUsuario_(p.cliente) === normUsuario_(u.usuario) && new Date(p.fecha).getTime() > Date.now() - 864e5).length;
      if (hoy >= 30) return { ok: false, error: 'demasiados' };
      const pid = nuevoId_('o');
      const p = { pid: pid, fecha: ahora_(), lugarId: l.id, lugarNombre: l.nombre, cliente: u.usuario, nombre: u.nombre, casa: `${u.calle} ${u.casa}`, telefono: u.telefono,
        items: JSON.stringify(items), total: total, entrega: entrega, nota: nota, estado: 'recibido', fechaEstado: ahora_(), dlat: conDest ? dl : '', dlon: conDest ? dn : '',
        historial: JSON.stringify([{ f: ahora_(), t: 'Pedido enviado', u: u.nombre }]) };
      escribirFila_(sh, H_PEDIDOS, null, p);
      bitacora_(u, 'Hizo un pedido', `${l.nombre} · ${pid}`);
      return { ok: true, pid: pid, total: total, items: items, entrega: entrega };
    } finally { lock.releaseLock(); }
  }
  if (a === 'pedidoEstado' || a === 'pedidoUbicacion' || a === 'pedidoSeguir') {
    const sh = hPedidos_(); const p = filas_(sh, H_PEDIDOS).map(pedidoDe_).find(x => x.pid === String(req.pid));
    if (!p) return { ok: false, error: 'no_existe' };
    const l = filas_(hLugares_(), H_LUGARES).map(lugarDe_).find(x => x.id === p.lugarId);
    const soyNegocio = !!l && normUsuario_(l.dueno) === normUsuario_(u.usuario);
    const soyCliente = normUsuario_(p.cliente) === normUsuario_(u.usuario);
    if (!soyNegocio && !soyCliente) return { ok: false, error: 'permiso' };
    const cache = CacheService.getScriptCache();
    if (a === 'pedidoSeguir') {
      let loc = null; try { loc = JSON.parse(cache.get('loc_' + p.pid) || 'null'); } catch (e) { }
      return { ok: true, estado: p.estado, fechaEstado: p.fechaEstado, loc: p.estado === 'camino' ? loc : null };
    }
    if (a === 'pedidoUbicacion') {
      if (!soyNegocio) return { ok: false, error: 'permiso' };
      if (p.estado !== 'camino') return { ok: false, error: 'estado' };
      const lat = Number(req.lat), lon = Number(req.lon); if (!enZona_(lat, lon)) return { ok: false, error: 'ubicacion' };
      cache.put('loc_' + p.pid, JSON.stringify({ lat: lat, lon: lon, acc: Math.min(5000, Number(req.acc) || 30), f: ahora_() }), 21600);
      return { ok: true };
    }
    const est = String(req.estado); if (ESTADOS_PED.indexOf(est) < 0) return { ok: false, error: 'estado' };
    if (!soyNegocio && !(soyCliente && est === 'cancelado' && p.estado === 'recibido')) return { ok: false, error: 'permiso' };
    const txt = { recibido: 'Recibido', preparando: 'Preparando', camino: 'En camino', entregado: 'Entregado', cancelado: 'Cancelado' }[est];
    p.historial.push({ f: ahora_(), t: txt, u: u.nombre });
    sh.getRange(p._row, H_PEDIDOS.indexOf('estado') + 1, 1, 2).setValues([[est, ahora_()]]);
    sh.getRange(p._row, H_PEDIDOS.indexOf('historial') + 1).setValue(JSON.stringify(p.historial.slice(-30)));
    if (est === 'entregado' || est === 'cancelado') cache.remove('loc_' + p.pid);
    bitacora_(u, 'Pedido: ' + txt.toLowerCase(), `${p.lugarNombre} · ${p.pid}`);
    return { ok: true };
  }

  // ---- Seguimiento (comité y operadores) ----
  if (!staff) return { ok: false, error: 'permiso' };

  if (a === 'reporteEstado') {
    const est = String(req.estado || '');
    if (ESTADOS_REP.indexOf(est) < 0) return { ok: false, error: 'estado' };
    if (est === 'descartado' && !comite) return { ok: false, error: 'permiso' };
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const sh = hReportes_(); const r = filas_(sh, H_REPORTES).map(reporteDe_).find(x => x.rid === String(req.rid));
      if (!r) return { ok: false, error: 'no_existe' };
      const txt = { nuevo: 'Regresó a nuevo', proceso: r.itemId || INVENTARIO.indexOf(r.categoria) >= 0 ? 'Aceptado, en atención' : 'En atención', atendido: 'Atendido', descartado: 'Descartado' }[est];
      const nota = limpiar_(req.nota, 300);
      r.historial.push({ f: ahora_(), t: txt + (nota ? ': ' + nota : ''), u: u.nombre });
      const notas = req.notas !== undefined ? limpiar_(req.notas, 600) : r.notas;
      const cuad = req.cuadrilla !== undefined ? limpiar_(req.cuadrilla, 10) : r.cuadrilla;
      const base = H_REPORTES.indexOf('estado') + 1;
      sh.getRange(r._row, base, 1, 5).setValues([[est, ahora_(), u.usuario, notas, cuad]]);
      sh.getRange(r._row, H_REPORTES.indexOf('historial') + 1).setValue(JSON.stringify(r.historial.slice(-60)));
      if (req.itemId !== undefined) sh.getRange(r._row, H_REPORTES.indexOf('itemId') + 1, 1, 2).setValues([[limpiar_(req.itemId, 40), limpiar_(req.itemCodigo, 40)]]);
      bitacora_(u, `Reporte ${txt.toLowerCase()}`, `${r.categoria} · ${r.tipo} · ${r.rid}${nota ? ' · ' + nota : ''}`);
      return { ok: true };
    } finally { lock.releaseLock(); }
  }
  if (a === 'reporteNotas') {
    const sh = hReportes_(); const r = filas_(sh, H_REPORTES).map(reporteDe_).find(x => x.rid === String(req.rid));
    if (!r) return { ok: false, error: 'no_existe' };
    sh.getRange(r._row, H_REPORTES.indexOf('notas') + 1, 1, 2).setValues([[limpiar_(req.notas, 600), limpiar_(req.cuadrilla, 10)]]);
    bitacora_(u, 'Notas de reporte', r.rid);
    return { ok: true };
  }
  if (a === 'reporteFoto') {
    const sh = hReportes_(); const r = filas_(sh, H_REPORTES).map(reporteDe_).find(x => x.rid === String(req.rid));
    if (!r) return { ok: false, error: 'no_existe' };
    const foto = guardarFoto_(req.data, r.categoria + '_' + r.rid + '_' + Date.now());
    sh.getRange(r._row, H_REPORTES.indexOf('foto') + 1).setValue(foto);
    bitacora_(u, 'Foto de reporte', r.rid);
    return { ok: true, foto: foto };
  }
  if (a === 'itemGuardar') {
    const cat = String(req.categoria || ''); const d = req.data || {};
    if (INVENTARIO.indexOf(cat) < 0) return { ok: false, error: 'categoria' };
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const sh = hItems_(); const filas = filas_(sh, H_ITEMS);
      let fila = req.id ? filas.find(x => String(x.id) === String(req.id)) : null;
      if (req.id && !fila && !comite) return { ok: false, error: 'no_existe' };
      if (!fila && !comite) return { ok: false, error: 'permiso' };
      const prev = fila ? itemDe_(fila) : {};
      let next;
      if (comite) next = Object.assign({}, prev, d);
      else { // operador: solo seguimiento, notas y cuadrilla
        next = Object.assign({}, prev);
        ['estado', 'notas', 'cuadrilla', 'historial', 'fechaReparacion', 'actualizado'].forEach(k => { if (d[k] !== undefined) next[k] = d[k]; });
      }
      const id = fila ? String(fila.id) : nuevoId_(cat === 'luminarias' ? 'p' : 'd');
      const num = fila ? Number(fila.num) : (d.num || (Math.max(0, ...filas.filter(x => x.categoria === cat).map(x => Number(x.num) || 0)) + 1));
      next.id = id; next.categoria = cat; next.num = num; next.actualizado = ahora_();
      const lat = Number(next.lat), lon = Number(next.lon);
      if (!(lat > 16.9 && lat < 17.2 && lon > -96.9 && lon < -96.5)) return { ok: false, error: 'ubicacion' };
      const row = {
        id: id, categoria: cat, num: num, codigo: next.codigo || next.poste || '', falla: next.falla || '', estado: next.estado || '',
        ubicacion: next.ubicacion || '', entre: next.entre || '', lat: lat, lon: lon, foto: next.foto || '', actualizado: next.actualizado, por: u.usuario,
        data: JSON.stringify(next).slice(0, 45000)
      };
      escribirFila_(sh, H_ITEMS, fila ? fila._row : null, row);
      // Al quedar reparado/funcionando, cerrar los reportes ligados que sigan abiertos.
      const cerrado = next.estado === 'atendido' || next.tipo === 'ok';
      if (cerrado && !(prev.estado === 'atendido' || prev.tipo === 'ok')) {
        const shR = hReportes_();
        filas_(shR, H_REPORTES).map(reporteDe_).filter(r => r.itemId === id && (r.estado === 'nuevo' || r.estado === 'proceso')).forEach(r => {
          r.historial.push({ f: ahora_(), t: 'Atendido (se reparó ' + (cat === 'luminarias' ? 'el poste' : 'el registro') + ')', u: u.nombre });
          shR.getRange(r._row, H_REPORTES.indexOf('estado') + 1, 1, 3).setValues([['atendido', ahora_(), u.usuario]]);
          shR.getRange(r._row, H_REPORTES.indexOf('historial') + 1).setValue(JSON.stringify(r.historial.slice(-60)));
        });
      }
      bitacora_(u, fila ? 'Editó ' + cat : 'Agregó ' + cat, `#${num} ${row.codigo} · ${row.falla}${next.estado ? ' · ' + next.estado : ''}`);
      return { ok: true, item: next };
    } finally { lock.releaseLock(); }
  }
  if (a === 'itemFoto') {
    const sh = hItems_(); const fila = filas_(sh, H_ITEMS).find(x => String(x.id) === String(req.id));
    if (!fila) return { ok: false, error: 'no_existe' };
    const foto = guardarFoto_(req.data, fila.categoria + '_' + fila.id + '_' + Date.now());
    const d = itemDe_(fila); d.foto = foto;
    sh.getRange(fila._row, H_ITEMS.indexOf('foto') + 1).setValue(foto);
    sh.getRange(fila._row, H_ITEMS.indexOf('data') + 1).setValue(JSON.stringify(d).slice(0, 45000));
    bitacora_(u, 'Foto de ' + fila.categoria, `#${fila.num} ${fila.codigo}`);
    return { ok: true, foto: foto };
  }

  // ---- Solo comité ----
  if (!comite) return { ok: false, error: 'permiso' };

  if (a === 'itemBorrar') {
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const sh = hItems_(); const fila = filas_(sh, H_ITEMS).find(x => String(x.id) === String(req.id));
      if (!fila) return { ok: false, error: 'no_existe' };
      sh.deleteRow(fila._row);
      bitacora_(u, 'Eliminó ' + fila.categoria, `#${fila.num} ${fila.codigo}`);
      return { ok: true };
    } finally { lock.releaseLock(); }
  }
  if (a === 'contenidoGuardar') {
    const clave = String(req.clave); if (CONTENIDOS.indexOf(clave) < 0) return { ok: false, error: 'clave_contenido' };
    const txt = JSON.stringify(req.data === undefined ? null : req.data); if (txt.length > 45000) return { ok: false, error: 'lleno' };
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const sh = hContenido_(); const fila = filas_(sh, H_CONTENIDO).find(r => r.clave === clave);
      escribirFila_(sh, H_CONTENIDO, fila ? fila._row : null, { clave: clave, data: txt, actualizado: ahora_(), por: u.usuario });
      bitacora_(u, 'Actualizó ' + ({ comite: 'Conoce a tu comité', reglamentos: 'Reglamentos', numeros: 'Números oficiales' })[clave], '');
      return { ok: true };
    } finally { lock.releaseLock(); }
  }
  if (a === 'archivoSubir') {
    const mime = String(req.mime || ''); if (!MIMES[mime]) return { ok: false, error: 'tipo_archivo' };
    if (!req.data || req.data.length > 14e6) return { ok: false, error: 'foto' };
    const nombre = limpiar_(req.nombre, 80).replace(/[^\w .()-]/g, '') || 'archivo';
    const blob = Utilities.newBlob(Utilities.base64Decode(req.data), mime, nombre.replace(/\.\w+$/, '') + '.' + MIMES[mime]);
    const f = carpetaFotos_().createFile(blob); f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    bitacora_(u, 'Subió archivo', nombre);
    return { ok: true, id: f.getId() };
  }
  if (a === 'usuarios') return { ok: true, usuarios: usuarios_().map(x => Object.assign(publicoUsuario_(x), { creado: String(x.creado), ultimoAcceso: String(x.ultimoAcceso) })) };
  if (a === 'usuarioCambiar') {
    const x = buscarUsuario_(req.usuario); if (!x) return { ok: false, error: 'no_existe' };
    if (normUsuario_(x.usuario) === normUsuario_(u.usuario) && (req.rol && req.rol !== 'comite' || req.estado === 'bloqueado')) return { ok: false, error: 'a_ti_mismo' };
    const sh = hUsuarios_();
    if (req.rol) { if (ROLES.indexOf(req.rol) < 0) return { ok: false, error: 'rol' }; sh.getRange(x._row, H_USUARIOS.indexOf('rol') + 1).setValue(req.rol); }
    if (req.estado) {
      if (['activo', 'bloqueado'].indexOf(req.estado) < 0) return { ok: false, error: 'estado' };
      sh.getRange(x._row, H_USUARIOS.indexOf('estado') + 1).setValue(req.estado);
      if (req.estado === 'bloqueado') PROPS.setProperty('VER_' + x.usuario, String(Date.now()));
    }
    bitacora_(u, 'Cambió usuario', `${x.usuario}${req.rol ? ' · rol ' + req.rol : ''}${req.estado ? ' · ' + req.estado : ''}`);
    return { ok: true };
  }
  if (a === 'usuarioReset') {
    const x = buscarUsuario_(req.usuario); if (!x) return { ok: false, error: 'no_existe' };
    const abc = 'abcdefghjkmnpqrstuvwxyz23456789'; let tmp = '';
    for (let i = 0; i < 8; i++) tmp += abc[Math.floor(Math.random() * abc.length)];
    const salt = Utilities.getUuid();
    hUsuarios_().getRange(x._row, H_USUARIOS.indexOf('salt') + 1, 1, 2).setValues([[salt, hashPw_(salt, tmp)]]);
    PROPS.setProperty('VER_' + x.usuario, String(Date.now()));
    bitacora_(u, 'Restableció contraseña', x.usuario);
    return { ok: true, temporal: tmp };
  }
  if (a === 'bitacora') {
    const sh = hBitacora_(); const n = sh.getLastRow() - 1; if (n <= 0) return { ok: true, bitacora: [] };
    const k = Math.min(n, 300);
    const v = sh.getRange(n - k + 2, 1, k, H_BITACORA.length).getValues().reverse();
    return { ok: true, bitacora: v.map(r => ({ f: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), usuario: r[1], nombre: r[2], accion: r[3], detalle: r[4] })) };
  }
  if (a === 'cambiarClaveComite') {
    const nueva = String(req.nueva || ''); if (nueva.length < 6) return { ok: false, error: 'pw_corta' };
    PROPS.setProperty('CLAVE', nueva);
    bitacora_(u, 'Cambió la clave del comité', '');
    return { ok: true };
  }
  if (a === 'importar') { // Una sola vez: postes y reportes de la página anterior.
    const lock = LockService.getScriptLock(); lock.waitLock(30000);
    try {
      const shI = hItems_(); const ya = {}; filas_(shI, H_ITEMS).forEach(x => ya[String(x.id)] = 1);
      const rowsI = (req.items || []).filter(p => !ya[String(p.id)]).map(p => {
        const d = Object.assign({}, p, { categoria: 'luminarias', codigo: p.poste || p.codigo || '' });
        return H_ITEMS.map(h => ({ id: String(p.id), categoria: 'luminarias', num: p.num || '', codigo: d.codigo, falla: p.falla || '', estado: p.estado || '', ubicacion: p.ubicacion || '', entre: p.entre || '', lat: p.lat, lon: p.lon, foto: p.foto || '', actualizado: p.actualizado || ahora_(), por: 'importado', data: JSON.stringify(d).slice(0, 45000) })[h]);
      });
      if (rowsI.length) shI.getRange(shI.getLastRow() + 1, 1, rowsI.length, H_ITEMS.length).setValues(rowsI);
      const shR = hReportes_(); const yaR = {}; filas_(shR, H_REPORTES).forEach(x => yaR[String(x.rid)] = 1);
      const rowsR = (req.reportes || []).filter(r => !yaR[String(r.rid)]).map(r => {
        const o = {
          rid: String(r.rid), fecha: r.fecha || ahora_(), categoria: 'luminarias', itemId: r.posteId || '', itemCodigo: r.posteNum || '', tipo: r.falla || '', comentario: r.comentario || '',
          usuario: 'anterior', nombre: r.contacto || 'Vecino (página anterior)', casa: '', telefono: '', lat: r.lat, lon: r.lon, foto: '',
          estado: r.estado === 'aceptado' ? 'proceso' : r.estado === 'descartado' ? 'descartado' : 'nuevo', fechaEstado: r.fechaRes || r.fecha || ahora_(), por: '', notas: '', cuadrilla: '',
          apoyos: '[]', historial: JSON.stringify([{ f: r.fecha || ahora_(), t: 'Reporte recibido (página anterior)', u: r.contacto || '' }])
        };
        return H_REPORTES.map(h => o[h]);
      });
      if (rowsR.length) shR.getRange(shR.getLastRow() + 1, 1, rowsR.length, H_REPORTES.length).setValues(rowsR);
      bitacora_(u, 'Importó datos de la página de postes', `${rowsI.length} postes · ${rowsR.length} reportes`);
      return { ok: true, items: rowsI.length, reportes: rowsR.length };
    } finally { lock.releaseLock(); }
  }
  return { ok: false, error: 'accion' };
}

/** Ejecutar una vez desde el editor: crea la hoja, las pestañas y la carpeta de fotos. */
function instalar() {
  hUsuarios_(); hItems_(); hReportes_(); hBitacora_(); hLugares_(); hPedidos_(); hContenido_(); carpetaFotos_(); secreto_();
  if (!PROPS.getProperty('CLAVE')) PROPS.setProperty('CLAVE', 'CAMBIAR-ESTA-CLAVE');
  Logger.log('Hoja: ' + ss_().getUrl());
}
