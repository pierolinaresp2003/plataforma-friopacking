/**
 * POLAR IA — intermediario entre la plataforma de reportes de Grupo Friopacking y la API de Claude.
 *
 * Por qué existe: la plataforma es una página pública (GitHub Pages), así que la API key de Anthropic
 * NO puede ir dentro del index.html. Este Apps Script la guarda en las Propiedades del script y solo
 * reenvía la conversación. Los datos (requerimientos, stock, OC, indicadores) nunca pasan por aquí
 * de golpe: Claude pide lo que necesita con "herramientas" y la página las responde con lo que ya
 * tiene cargado.
 *
 * INSTALACIÓN (una sola vez):
 *  1. https://script.google.com → Nuevo proyecto → pega este archivo → nómbralo "Polar IA".
 *  2. Engranaje (Configuración del proyecto) → Propiedades del script → Agregar:
 *       ANTHROPIC_API_KEY = sk-ant-...   (se crea en https://console.anthropic.com → API Keys)
 *  3. Implementar → Nueva implementación → Tipo: Aplicación web
 *       Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario → Implementar → autorizar.
 *  4. Copia la URL que termina en /exec y pégala en index.html, en  const POLAR_IA_URL = '...';
 *  5. En https://console.anthropic.com → Billing → pon un límite de gasto mensual.
 */

const MODELO = 'claude-sonnet-5-5';     // balance entre inteligencia y costo ($2 / $10 por millón de tokens)
const LIMITE_DIARIO = 400;              // consultas por día para toda la empresa (freno de costos)
const MAX_TOKENS = 1200;

const TOOLS = {
  consultar_requerimiento: {
    description: 'Detalle de un requerimiento: proyecto, solicitante, fecha y cada ítem con su estado (En Revisión, Pedido pendiente, Enviado a pedido — cuenta con OC, Recibido en Almacén, Enviado a Obra, Recibido en Obra, Archivado), pedido, OC, proveedor, fecha de entrega de la OC, cantidades recibidas, despachadas y recibidas en obra. Acepta "23303" o "0001-0023303".',
    input_schema: { type: 'object', properties: { codigo: { type: 'string' } }, required: ['codigo'] } },
  consultar_proyecto: {
    description: 'Resumen de un proyecto (código PRY… o nombre del cliente): total de ítems por estado y la lista de sus requerimientos con avance. Si el nombre coincide con varios proyectos devuelve las opciones.',
    input_schema: { type: 'object', properties: { proyecto: { type: 'string' } }, required: ['proyecto'] } },
  requerimientos_de_solicitante: {
    description: 'Requerimientos pedidos por una persona (el supervisor o responsable que hizo el requerimiento), por nombre.',
    input_schema: { type: 'object', properties: { nombre: { type: 'string' } }, required: ['nombre'] } },
  consultar_oc: {
    description: 'Detalle de una orden de compra: proveedor, estado, fechas (emisión, aprobación, entrega pactada), ítems y a qué requerimiento pertenecen.',
    input_schema: { type: 'object', properties: { codigo: { type: 'string' } }, required: ['codigo'] } },
  stock_producto: {
    description: 'Stock de un producto por almacén (Lima, Trujillo, Ica), tránsito, reservado en Lima, disponible real, rotación mensual, punto de reorden y OC en camino. Acepta código de 12 dígitos o palabras de la descripción.',
    input_schema: { type: 'object', properties: { consulta: { type: 'string' } }, required: ['consulta'] } },
  pendientes_compras: {
    description: 'Pendientes del área de Compras: OC por llegar, OC vencidas (días hábiles) y pedidos sin OC; de todas las personas o de una (Carolina Checa, Melissa Sihuairo, Jaime Altamirano, José Pomez, Manuel Luján, Piero Linares, Javier Centeno, Katherine Albujar, Camila Mora).',
    input_schema: { type: 'object', properties: { persona: { type: 'string' } } } },
  productos_por_reponer: {
    description: 'Productos en punto de reorden con la cantidad sugerida a pedir, ordenados por los que se quedan sin stock primero.',
    input_schema: { type: 'object', properties: { limite: { type: 'integer' } } } },
  indicadores_logistica: {
    description: 'Los 5 indicadores de tiempo (generación de pedido, emisión de OC, aprobación de OC, entrega en Lurín —solo nacional—, despacho de requerimiento): promedio en días hábiles, % que cumple la meta y casos. Filtros opcionales: persona de Compras, año, mes (1-12), origen (nacional/importado).',
    input_schema: { type: 'object', properties: { persona: { type: 'string' }, anio: { type: 'integer' }, mes: { type: 'integer' }, origen: { type: 'string', enum: ['nacional', 'importado'] } } } },
  gasto_compras: {
    description: 'Concentración de gasto de compras: total equivalente en USD, nacional/importado en soles y dólares, top familias y top proveedores. Filtros opcionales: año, mes, familia.',
    input_schema: { type: 'object', properties: { anio: { type: 'integer' }, mes: { type: 'integer' }, familia: { type: 'string' } } } },
  exportar_excel: {
    description: 'Genera y descarga en el equipo del usuario un Excel. tipo: requerimiento (id = N° req), proyecto (id = PRY o nombre), pendientes (persona opcional), reponer, indicadores (persona/año/mes opcionales).',
    input_schema: { type: 'object', properties: { tipo: { type: 'string', enum: ['requerimiento', 'proyecto', 'pendientes', 'reponer', 'indicadores'] },
      id: { type: 'string' }, persona: { type: 'string' }, anio: { type: 'integer' }, mes: { type: 'integer' } }, required: ['tipo'] } },
};

// Qué puede consultar cada vista. El supervisor solo ve sus requerimientos y proyectos.
function herramientasPara(vista, ger) {
  let nombres = ['consultar_requerimiento', 'consultar_proyecto', 'requerimientos_de_solicitante', 'consultar_oc', 'exportar_excel'];
  if (vista !== 'supervisor') nombres = nombres.concat(['stock_producto', 'pendientes_compras', 'productos_por_reponer']);
  if (vista !== 'supervisor' && ger) nombres = nombres.concat(['indicadores_logistica', 'gasto_compras']);
  return nombres.map(function (n) { return Object.assign({ name: n }, TOOLS[n]); });
}

function sistema(req) {
  const vistas = {
    supervisor: 'La persona es un SUPERVISOR de obra: solo le interesa cómo van sus requerimientos y proyectos. No tiene acceso a stock, pendientes de Compras, indicadores ni gasto; si lo pide, explícale amablemente que en su vista solo ves requerimientos y proyectos.',
    logistica: 'La persona es del área de LOGÍSTICA / COMPRAS: le interesan stock, kardex, qué comprar para stock, pendientes de compra y requerimientos.',
    gerencia: 'La persona es de GERENCIA: le interesan indicadores, gasto y el estado general.',
    completa: 'La persona usa la plataforma completa desde la PC.',
  };
  return [
    'Eres Leo, el asistente de la plataforma de reportes de Grupo Friopacking (cadena de frío, Perú).',
    vistas[req.vista] || vistas.completa,
    req.ger ? 'El módulo Gerencial está desbloqueado.' : 'El módulo Gerencial está bloqueado: indicadores y gasto solo se ven con la clave; si los piden, di que ingresen la clave en Gerencia.',
    'Hoy es ' + (req.hoy || '') + (req.corte ? '. Los datos están al ' + req.corte : '') + '.',
    'Reglas:',
    '- Responde SIEMPRE en español, corto y directo, como para leer en el celular. Usa **negritas** para lo importante y listas con "- " cuando haya varios elementos.',
    '- Nunca inventes datos: usa las herramientas. Si una herramienta no encuentra algo, dilo.',
    '- Un número suelto de 3 a 7 dígitos casi siempre es un requerimiento (0001-00xxxxx). "PRY…" es un proyecto.',
    '- Al hablar de un requerimiento, di el estado de cada ítem con palabras simples: en revisión de almacén, en Compras sin OC, con OC por llegar (con fecha si la hay), en almacén, enviado a obra, recibido en obra. Señala lo que falta.',
    '- Si piden un Excel, usa exportar_excel y confirma el nombre del archivo descargado.',
    '- "d.h." significa días hábiles. Las importaciones no cuentan en el indicador de entrega en Lurín.',
    '- Solo temas de la plataforma y la operación de Friopacking; para otra cosa, redirige con amabilidad.',
  ].join('\n');
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const props = PropertiesService.getScriptProperties();
    const key = props.getProperty('ANTHROPIC_API_KEY');
    if (!key) return json({ error: 'Falta ANTHROPIC_API_KEY en las propiedades del script.' });

    const hoy = Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd');
    const usadas = Number(props.getProperty('uso_' + hoy) || 0);
    // Una "consulta" es la primera vuelta de cada pregunta (las vueltas de herramientas no suman).
    const esNueva = Array.isArray(req.messages) && req.messages.length && typeof req.messages[req.messages.length - 1].content === 'string';
    if (esNueva && usadas >= LIMITE_DIARIO) return json({ error: 'Se alcanzó el límite de consultas de hoy.' });
    if (esNueva) props.setProperty('uso_' + hoy, String(usadas + 1));

    const cuerpo = {
      model: MODELO, max_tokens: MAX_TOKENS,
      system: sistema(req),
      tools: herramientasPara(req.vista, !!req.ger),
      messages: (req.messages || []).slice(-30),
    };
    const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(cuerpo),
    });
    return ContentService.createTextOutput(r.getContentText()).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return json({ error: String(err && err.message || err) });
  }
}

function doGet() { return json({ ok: true, servicio: 'Polar IA', modelo: MODELO }); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
