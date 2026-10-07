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
 * ACTUALIZAR (si ya estaba instalado): pega este archivo encima del anterior → Guardar →
 *   Implementar → Administrar implementaciones → ✏️ editar → Versión: Nueva versión → Implementar.
 *   Así la URL /exec sigue siendo la misma y no hay que tocar el index.html.
 */

const HOJA = 'Preguntas';

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.tipo === 'tarea') return tareaGuardar(d);
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
function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
