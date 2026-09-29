/**
 * POLAR APRENDE — anota en una hoja de Google las preguntas que Polar no entendió.
 * Gratis y sin IA: solo guarda la pregunta para que luego se le enseñe a Polar a responderla.
 *
 * INSTALACIÓN (una sola vez, 5 minutos):
 *  1. Crea una hoja de Google nueva, por ejemplo "Polar - preguntas sin responder".
 *  2. En esa hoja: Extensiones → Apps Script → borra lo que haya → pega este archivo → Guardar.
 *  3. Implementar → Nueva implementación → Tipo: Aplicación web
 *       Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario → Implementar → autorizar.
 *  4. Copia la URL que termina en /exec y pásasela a Claude (va en const POLAR_APRENDE_URL del index.html).
 *
 * Qué se guarda: fecha, vista (supervisor / logística / gerencia), la pregunta tal cual,
 * cómo la entendió Polar después de corregir tipeo y sinónimos, el motivo y la fecha de los datos.
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
      hoja.appendRow(['Fecha', 'Vista', 'Pregunta', 'Cómo la entendió Polar', 'Motivo', 'Datos al', '¿Ya se le enseñó?']);
      hoja.setFrozenRows(1);
      hoja.getRange('A1:G1').setFontWeight('bold').setBackground('#0a1e64').setFontColor('#ffffff');
      hoja.setColumnWidths(1, 1, 150); hoja.setColumnWidth(3, 380); hoja.setColumnWidth(4, 320);
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
      String(d.motivo || ''), String(d.datos || ''), '',
    ]);
    return salida({ ok: true });
  } catch (err) {
    return salida({ ok: false, error: String(err) });
  }
}

function doGet() { return salida({ ok: true, servicio: 'Polar aprende' }); }
function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
