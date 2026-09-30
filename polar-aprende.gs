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
 * No se guarda quién preguntó ni ningún dato de la empresa.
 */

const HOJA = 'Preguntas';

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const pregunta = String(d.pregunta || '').trim().slice(0, 300);
    if (!pregunta) return salida({ ok: false });
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = libro.getSheetByName(HOJA);
    if (!hoja) {
      hoja = libro.insertSheet(HOJA);
      hoja.appendRow(['Fecha', 'Vista', 'Pregunta', 'Cómo la entendió Leo', 'Motivo', 'Datos al', '¿Ya se le enseñó?', 'Respuesta de Leo']);
      hoja.setFrozenRows(1);
      hoja.setColumnWidths(1, 1, 150); hoja.setColumnWidth(3, 380); hoja.setColumnWidth(4, 320);
    }
    // Hojas creadas con la versión anterior: se agrega la columna H.
    if (!hoja.getRange('H1').getValue()) {
      hoja.getRange('H1').setValue('Respuesta de Leo'); hoja.setColumnWidth(8, 420);
      hoja.getRange('A1:H1').setFontWeight('bold').setBackground('#0a1e64').setFontColor('#ffffff');
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
    ]);
    return salida({ ok: true });
  } catch (err) {
    return salida({ ok: false, error: String(err) });
  }
}

function doGet() { return salida({ ok: true, servicio: 'Polar aprende' }); }
function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
