/**
 * POLAR APRENDE — anota en una hoja de Google TODAS las consultas a Leo, qué respondió y si el
 * usuario marcó 👍 / 👎. Gratis y sin IA: sirve para enseñarle a Leo lo que no sabe o contesta mal.
 *
 * INSTALACIÓN (una sola vez, 5 minutos):
 *  1. Crea una hoja de Google nueva, por ejemplo "Polar - preguntas sin responder".
 *  2. En esa hoja: Extensiones → Apps Script → borra lo que haya → pega este archivo → Guardar.
 *  3. Implementar → Nueva implementación → Tipo: Aplicación web
 *       Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario → Implementar → autorizar.
 *  4. Copia la URL que termina en /exec y pásasela a Claude (va en const POLAR_APRENDE_URL del index.html).
 *
 * Qué se guarda: fecha, vista (supervisor / logística / gerencia), la pregunta tal cual,
 * cómo la entendió Leo después de corregir tipeo y sinónimos, el motivo (respondió / no entendió /
 * no encontró / 👍 le sirvió / 👎 respuesta incorrecta), la fecha de los datos y un resumen de la respuesta.
 * Desde el 01.10.2026 también se guarda quién preguntó (correo y nombre), a pedido de Piero.
 * Desde el 06.10.2026 (v4): cuando alguien marca 👎, Leo le pregunta "¿Qué esperabas ver?" y la respuesta
 * se guarda en la columna J. Es lo que Leo necesita para aprender a contestar bien esa consulta.
 *
 * Desde el 07.10.2026 (v5): también guarda las TAREAS MANUALES del Daily call in en la pestaña "Daily tareas"
 * (quién, qué tarea, estado Pendiente / En proceso / Cumplida / No cumplida y comentario de cómo quedó).
 * Cada comprador solo anota y actualiza sus propias tareas (se valida con su correo: plinares@… = PLINARES).
 * La plataforma las lee con doGet?accion=tareas. Misma hoja, misma URL: solo hay que actualizar (abajo).
 *
 * Desde el 12.10.2026 (v6): guarda el DAILY CALL IN de cada colaborador en la pestaña "Daily call in": qué avanzó,
 * en qué trabajará, bloqueos, prioridad y los comentarios del equipo. Cada uno solo registra y edita lo suyo
 * (mismo control por correo); comentar puede cualquiera. La plataforma lo lee con doGet?accion=dci.
 *
 * Desde el 13.10.2026 (v7): cada tarea del Daily call in tiene PROYECTO / FRENTE y ESTADO (Pendiente / En proceso /
 * Finalizada, con la fecha en que se finalizó). Cambia el estado el dueño de la tarea o PLINARES (administrador).
 * También guarda la FOTO de cada colaborador (pestaña "Daily fotos", imagen chica en texto). Las columnas nuevas se
 * agregan solas al final de la hoja; no se pierde nada de lo ya registrado.
 *
 * Desde el 15.10.2026 (v8): guarda cada INGRESO A LA APP en la pestaña "Ingresos al app": fecha, hora (Lima), nombre,
 * correo, acceso, módulo con el que entró y dispositivo. Se anota una vez por cada vez que se abre la app. La plataforma lo
 * lee con doGet?accion=ingresos&dias=30 (lo ve el módulo Gerencial).
 *
 * ACTUALIZAR (si ya estaba instalado): pega este archivo encima del anterior → Guardar →
 *   Implementar → Administrar implementaciones → ✏️ editar → Versión: Nueva versión → Implementar.
 *   Así la URL /exec sigue siendo la misma y no hay que tocar el index.html.
 */

const HOJA = 'Preguntas';

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.tipo === 'tarea') return tareaGuardar(d);
    if (d.tipo === 'dci') return dciGuardar(d);
    if (d.tipo === 'ingreso') return ingresoGuardar(d);
    const pregunta = String(d.pregunta || '').trim().slice(0, 300);
    if (!pregunta) return salida({ ok: false });
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = libro.getSheetByName(HOJA);
    if (!hoja) {
      hoja = libro.insertSheet(HOJA);
      hoja.appendRow(['Fecha', 'Vista', 'Pregunta', 'Cómo la entendió Leo', 'Motivo', 'Datos al', '¿Ya se le enseñó?', 'Respuesta de Leo', 'Usuario', '¿Qué esperaba?']);
      hoja.setFrozenRows(1);
      hoja.setColumnWidths(1, 1, 150); hoja.setColumnWidth(3, 380); hoja.setColumnWidth(4, 320);
    }
    // Hojas creadas con versiones anteriores: se agregan las columnas H, I y J.
    if (!hoja.getRange('J1').getValue()) {
      hoja.getRange('J1').setValue('¿Qué esperaba?'); hoja.setColumnWidth(10, 380);
      hoja.getRange('J1').setFontWeight('bold').setBackground('#0a1e64').setFontColor('#ffffff');
    }
    if (!hoja.getRange('H1').getValue() || !hoja.getRange('I1').getValue()) {
      hoja.getRange('H1').setValue('Respuesta de Leo'); hoja.setColumnWidth(8, 420);
      hoja.getRange('I1').setValue('Usuario'); hoja.setColumnWidth(9, 260);
      hoja.getRange('A1:I1').setFontWeight('bold').setBackground('#0a1e64').setFontColor('#ffffff');
    }
    // Freno simple: como máximo 2,000 filas nuevas por día.
    const props = PropertiesService.getScriptProperties();
    const hoy = Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd');
    const n = Number(props.getProperty('dia_' + hoy) || 0);
    if (n >= 2000) return salida({ ok: false, motivo: 'límite diario' });
    props.setProperty('dia_' + hoy, String(n + 1));

    hoja.appendRow([
      Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HH:mm'),
      String(d.vista || ''), pregunta, String(d.entendido || '').slice(0, 300),
      String(d.motivo || ''), String(d.datos || ''), '', String(d.respuesta || '').slice(0, 400),
      String(d.usuario || '').slice(0, 120), String(d.esperaba || '').slice(0, 250),
    ]);
    return salida({ ok: true });
  } catch (err) {
    return salida({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.accion === 'tareas') return tareasLeer();
  if (e && e.parameter && e.parameter.accion === 'dci') return dciLeer();
  if (e && e.parameter && e.parameter.accion === 'ingresos') return ingresosLeer(e.parameter.dias);
  return salida({ ok: true, servicio: 'Polar aprende' });
}

// ---------- Daily call in: tareas manuales (07.10.2026) ----------
const HOJA_TAREAS = 'Daily tareas';
const COLS_TAREAS = ['ID', 'Día', 'Código', 'Comprador', 'Tarea', 'Estado', 'Comentario', 'Registró', 'Creada', 'Actualizada', 'Actualizó'];
const ESTADOS = ['Pendiente', 'En proceso', 'Cumplida', 'No cumplida'];

function hojaTareas() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(HOJA_TAREAS);
  if (!h) {
    h = libro.insertSheet(HOJA_TAREAS);
    h.appendRow(COLS_TAREAS);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, COLS_TAREAS.length).setFontWeight('bold').setBackground('#0a1e64').setFontColor('#ffffff');
    h.getRange('B:B').setNumberFormat('@');   // el día se guarda como texto aaaa-mm-dd
    h.setColumnWidth(5, 420); h.setColumnWidth(7, 320); h.setColumnWidth(4, 160);
  }
  return h;
}
function ahora() { return Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HH:mm'); }
function txt(v, n) { return String(v == null ? '' : v).trim().slice(0, n); }
function diaTxt(v) { return v instanceof Date ? Utilities.formatDate(v, 'America/Lima', 'yyyy-MM-dd') : String(v || ''); }

function tareasLeer() {
  const h = hojaTareas(), n = h.getLastRow();
  if (n < 2) return salida({ ok: true, tareas: [] });
  const lim = new Date(); lim.setDate(lim.getDate() - 120);
  const desde = Utilities.formatDate(lim, 'America/Lima', 'yyyy-MM-dd');
  const filas = h.getRange(2, 1, n - 1, COLS_TAREAS.length).getValues();
  const tareas = filas.filter(function (f) { return f[0] && diaTxt(f[1]) >= desde; }).map(function (f) {
    return { id: String(f[0]), dia: diaTxt(f[1]), cod: String(f[2]), tarea: String(f[4]), estado: String(f[5] || 'Pendiente'),
      com: String(f[6] || ''), por: String(f[7] || ''), creada: String(f[8] || ''), act: String(f[9] || ''), actpor: String(f[10] || '') };
  });
  return salida({ ok: true, tareas: tareas });
}

function tareaGuardar(d) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const h = hojaTareas(), quien = txt(d.usuario, 120);
    // Cada comprador solo anota y actualiza SUS tareas: el usuario del correo (plinares@…) debe ser el código.
    const yo = String(d.correo || '').trim().toLowerCase().split('@')[0].toUpperCase();
    if (!yo) return salida({ ok: false, error: 'sin usuario' });
    if (d.accion === 'crear') {
      const tarea = txt(d.tarea, 400), cod = txt(d.cod, 20), dia = txt(d.dia, 10);
      if (!tarea || !cod || !/^\d{4}-\d{2}-\d{2}$/.test(dia)) return salida({ ok: false, error: 'faltan datos' });
      if (cod !== yo) return salida({ ok: false, error: 'solo puedes anotar tareas a tu nombre' });
      const id = Utilities.getUuid().slice(0, 8);
      h.appendRow([id, dia, cod, txt(d.nombre, 80), tarea, 'Pendiente', '', quien, ahora(), '', '']);
      return salida({ ok: true, tarea: { id: id, dia: dia, cod: cod, tarea: tarea, estado: 'Pendiente', com: '', por: quien, creada: ahora(), act: '', actpor: '' } });
    }
    const id = txt(d.id, 20); if (!id) return salida({ ok: false, error: 'sin id' });
    const celda = h.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
    if (!celda) return salida({ ok: false, error: 'no existe' });
    const fila = celda.getRow();
    if (String(h.getRange(fila, 3).getValue()) !== yo) return salida({ ok: false, error: 'solo el dueño puede cambiar esta tarea' });
    if (d.accion === 'borrar') { h.deleteRow(fila); return salida({ ok: true }); }
    const estado = ESTADOS.indexOf(d.estado) !== -1 ? d.estado : h.getRange(fila, 6).getValue();
    h.getRange(fila, 6, 1, 2).setValues([[estado, txt(d.com, 400)]]);
    h.getRange(fila, 10, 1, 2).setValues([[ahora(), quien]]);
    return salida({ ok: true });
  } finally { lock.releaseLock(); }
}
// ---------- Daily Call In (12.10.2026, v6) ----------
const HOJA_DCI = 'Daily call in';
const COLS_DCI = ['ID', 'Día', 'Código', 'Colaborador', 'Área', '¿Qué avancé?', '¿En qué trabajaré?', 'Bloqueos', 'Prioridad', 'Comentarios', 'Registró', 'Creada', 'Actualizada', 'Proyecto / frente', 'Estado', 'Finalizada'];
const ESTADOS_DCI = ['Pendiente', 'En proceso', 'Finalizada'];
const DCI_ADMIN = ['PLINARES'];
const HOJA_FOTOS = 'Daily fotos';
const PRIORIDADES = ['Alta', 'Media', 'Baja'];
function hojaDci() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(HOJA_DCI);
  if (!h) {
    h = libro.insertSheet(HOJA_DCI);
    h.appendRow(COLS_DCI);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, COLS_DCI.length).setFontWeight('bold').setBackground('#0a1e64').setFontColor('#ffffff');
    h.getRange('B:B').setNumberFormat('@');
    h.setColumnWidth(6, 380); h.setColumnWidth(7, 380); h.setColumnWidth(8, 260); h.setColumnWidth(10, 320);
  }
  // v7: columnas nuevas al final (hojas creadas con la v6).
  if (h.getLastColumn() < COLS_DCI.length || !h.getRange(1, COLS_DCI.length).getValue()) {
    h.getRange(1, 1, 1, COLS_DCI.length).setValues([COLS_DCI]).setFontWeight('bold').setBackground('#0a1e64').setFontColor('#ffffff');
    h.setColumnWidth(14, 220);
  }
  return h;
}
function hojaFotos() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(HOJA_FOTOS);
  if (!h) { h = libro.insertSheet(HOJA_FOTOS); h.appendRow(['Código', 'Foto', 'Actualizada']); h.setFrozenRows(1); }
  return h;
}
function dciLeer() {
  const h = hojaDci(), n = h.getLastRow();
  if (n < 2) return salida({ ok: true, dci: [] });
  const lim = new Date(); lim.setDate(lim.getDate() - 120);
  const desde = Utilities.formatDate(lim, 'America/Lima', 'yyyy-MM-dd');
  const filas = h.getRange(2, 1, n - 1, COLS_DCI.length).getValues();
  const dci = filas.filter(function (f) { return f[0] && (diaTxt(f[1]) >= desde || String(f[14] || 'Pendiente') !== 'Finalizada'); }).map(function (f) {
    let com = []; try { com = JSON.parse(f[9] || '[]'); } catch (err) {}
    return { id: String(f[0]), dia: diaTxt(f[1]), cod: String(f[2]), nombre: String(f[3] || ''), area: String(f[4] || ''), av: String(f[5] || ''), tr: String(f[6] || ''),
      bl: String(f[7] || ''), pr: String(f[8] || 'Media'), com: com, por: String(f[10] || ''), creada: String(f[11] || ''), act: String(f[12] || ''),
      py: String(f[13] || ''), es: ESTADOS_DCI.indexOf(String(f[14])) !== -1 ? String(f[14]) : 'Pendiente', ff: String(f[15] || '') };
  });
  const fotos = {}, hf = hojaFotos(), nf = hf.getLastRow();
  if (nf >= 2) hf.getRange(2, 1, nf - 1, 2).getValues().forEach(function (f) { if (f[0] && f[1]) fotos[String(f[0])] = String(f[1]); });
  return salida({ ok: true, dci: dci, fotos: fotos });
}
function dciGuardar(d) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const h = hojaDci(), quien = txt(d.usuario, 120);
    const yo = String(d.correo || '').trim().toLowerCase().split('@')[0].toUpperCase();
    if (!yo) return salida({ ok: false, error: 'sin usuario' });
    const pr = PRIORIDADES.indexOf(d.pr) !== -1 ? d.pr : 'Media';
    if (d.accion === 'crear') {
      const cod = txt(d.cod, 20), dia = txt(d.dia, 10), av = txt(d.av, 1000), tr = txt(d.tr, 1000);
      if (!cod || (!av && !tr) || !/^\d{4}-\d{2}-\d{2}$/.test(dia)) return salida({ ok: false, error: 'faltan datos' });
      if (cod !== yo) return salida({ ok: false, error: 'solo puedes registrar tu propio daily' });
      const id = Utilities.getUuid().slice(0, 8), t = ahora(), py = txt(d.py, 80);
      h.appendRow([id, dia, cod, txt(d.nombre, 80), txt(d.area, 30), av, tr, txt(d.bl, 600), pr, '[]', quien, t, '', py, 'Pendiente', '']);
      return salida({ ok: true, dci: { id: id, dia: dia, cod: cod, nombre: txt(d.nombre, 80), area: txt(d.area, 30), av: av, tr: tr, bl: txt(d.bl, 600), pr: pr, com: [], por: quien, creada: t, act: '', py: py, es: 'Pendiente', ff: '' } });
    }
    if (d.accion === 'foto') {
      const foto = String(d.foto || '');
      if (!/^data:image\/(jpeg|png|webp);base64,/.test(foto) || foto.length > 48000) return salida({ ok: false, error: 'foto no válida' });
      const hf = hojaFotos(), c = hf.getRange('A:A').createTextFinder(yo).matchEntireCell(true).findNext();
      if (c) hf.getRange(c.getRow(), 2, 1, 2).setValues([[foto, ahora()]]); else hf.appendRow([yo, foto, ahora()]);
      return salida({ ok: true });
    }
    const id = txt(d.id, 20); if (!id) return salida({ ok: false, error: 'sin id' });
    const celda = h.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
    if (!celda) return salida({ ok: false, error: 'no existe' });
    const fila = celda.getRow();
    if (d.accion === 'comentar') {
      const tx = txt(d.tx, 500); if (!tx) return salida({ ok: false, error: 'vacío' });
      let com = []; try { com = JSON.parse(h.getRange(fila, 10).getValue() || '[]'); } catch (err) {}
      const c = { por: txt(d.nombre, 80) || quien, t: ahora(), tx: tx }; com.push(c);
      h.getRange(fila, 10).setValue(JSON.stringify(com));
      return salida({ ok: true, com: c });
    }
    if (d.accion === 'estado') {
      if (String(h.getRange(fila, 3).getValue()) !== yo && DCI_ADMIN.indexOf(yo) === -1) return salida({ ok: false, error: 'solo el dueño puede cambiar el estado' });
      const es = ESTADOS_DCI.indexOf(d.es) !== -1 ? d.es : 'Pendiente', t = ahora();
      h.getRange(fila, 15, 1, 2).setValues([[es, es === 'Finalizada' ? t : '']]);
      h.getRange(fila, 13).setValue(t);
      return salida({ ok: true, ff: es === 'Finalizada' ? t : '' });
    }
    if (String(h.getRange(fila, 3).getValue()) !== yo) return salida({ ok: false, error: 'solo el dueño puede cambiar esto' });
    if (d.accion === 'borrar') { h.deleteRow(fila); return salida({ ok: true }); }
    h.getRange(fila, 6, 1, 4).setValues([[txt(d.av, 1000), txt(d.tr, 1000), txt(d.bl, 600), pr]]);
    if (d.py != null) h.getRange(fila, 14).setValue(txt(d.py, 80));
    h.getRange(fila, 13).setValue(ahora());
    return salida({ ok: true });
  } finally { lock.releaseLock(); }
}
function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ---------- Ingresos a la app (15.10.2026, v8) ----------
const HOJA_ING = 'Ingresos al app';
const COLS_ING = ['Fecha', 'Hora', 'Nombre', 'Correo', 'Acceso', 'Módulo', 'Dispositivo'];
function hojaIngresos() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(HOJA_ING);
  if (!h) {
    h = libro.insertSheet(HOJA_ING);
    h.appendRow(COLS_ING);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, COLS_ING.length).setFontWeight('bold').setBackground('#0a1e64').setFontColor('#ffffff');
    h.getRange('A:B').setNumberFormat('@');   // fecha y hora como texto (aaaa-mm-dd · hh:mm:ss, hora de Lima)
    h.setColumnWidth(3, 260); h.setColumnWidth(4, 260); h.setColumnWidth(7, 220);
  }
  return h;
}
function ingresoGuardar(d) {
  const correo = txt(d.correo, 120).toLowerCase();
  if (!/^[^@\s]+@[^@\s]+$/.test(correo)) return salida({ ok: false });
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const t = new Date();
    hojaIngresos().appendRow([Utilities.formatDate(t, 'America/Lima', 'yyyy-MM-dd'), Utilities.formatDate(t, 'America/Lima', 'HH:mm:ss'),
      txt(d.nombre, 120), correo, txt(d.acceso, 30), txt(d.vista, 30), txt(d.dispositivo, 80)]);
    return salida({ ok: true });
  } finally { lock.releaseLock(); }
}
function ingresosLeer(dias) {
  const h = hojaIngresos(), n = h.getLastRow();
  if (n < 2) return salida({ ok: true, ingresos: [] });
  const d0 = new Date(); d0.setDate(d0.getDate() - Math.min(366, Math.max(1, Number(dias) || 30)));
  const desde = Utilities.formatDate(d0, 'America/Lima', 'yyyy-MM-dd'), ini = Math.max(2, n - 5000 + 1);
  const l = h.getRange(ini, 1, n - ini + 1, COLS_ING.length).getDisplayValues()
    .filter(function (f) { return f[0] && String(f[0]) >= desde; })
    .map(function (f) { return { f: f[0], h: f[1], n: f[2], c: f[3], a: f[4], v: f[5], d: f[6] }; });
  return salida({ ok: true, ingresos: l.reverse() });
}
