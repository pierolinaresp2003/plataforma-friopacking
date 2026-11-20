/**
 * ============================================================================
 * CLASIFICAR — ordena solos los PDF que Power Automate deja sueltos (20.11.2026)
 * ============================================================================
 * Archivo NUEVO del mismo proyecto de Apps Script (no reemplaza a Código.gs).
 *
 * Cuando el proveedor responde, Power Automate deja cada adjunto suelto en
 * "Certificados GFP" (o en "Detalle Certificados"), con el asunto pegado al nombre.
 * Este archivo lo lleva a su carpeta, igual que lo que se sube desde la app:
 *
 *   Detalle Certificados / PRY — CLIENTE / Req. / Código — PRODUCTO /
 *       CERTIFICADO — <nombre original>.pdf
 *       FICHA TECNICA — <nombre original>.pdf
 *
 * Cómo sabe a dónde va cada archivo:
 *   · Correo de UN producto: el asunto trae  [ref PRY|Req|Código].
 *   · Correo de VARIOS productos: el asunto trae  [ref PRY|SOL XXXXXX]  y la app,
 *     al generar el correo, le manda a este script la lista de lo pedido
 *     (N°, OC, Req., Código, Producto). Con eso:
 *       - Certificado → por el N° de OC (en el nombre y dentro del PDF).
 *       - Ficha técnica → por el modelo del producto (TCW063, 5SL4210-7RC, CAT6A…).
 *       - Si el proveedor nombró el archivo con el N° de la tabla ("3 - Certificado"), manda eso.
 *   · Si el asunto no trae SOL pero sí el PRY, usa todo lo pedido en ese proyecto.
 *   · Si sirve para varios productos, queda una copia en la carpeta de cada uno.
 *   · Si el certificado dice por dentro otra OC que la del nombre, NO se archiva:
 *     queda en la raíz como "REVISAR — …" para pedirlo de nuevo al proveedor.
 *   · Lo que no se pueda ubicar se queda en la raíz (la app lo mostrará "Por clasificar").
 *
 * INSTALACIÓN (una sola vez)
 *   1. En el editor: + (Archivo) → Script → nombre "Clasificar" → pega TODO esto.
 *   2. En Código.gs, dentro de doPost, debajo de la línea de 'mover', agrega:
 *        if(p.accion === 'solicitud') return guardarSolicitud(p);
 *   3. Guarda (Ctrl+S). Elige la función  instalarDisparador  y pulsa Ejecutar (acepta permisos).
 *   4. Implementar → Administrar implementaciones → ✏ → Versión: Nueva → Implementar (la URL no cambia).
 * ============================================================================
 */

const CL_ETQ = { cert:'CERTIFICADO', ficha:'FICHA TECNICA' };
const CL_MIN_IMG = 30 * 1024;          // imágenes más chicas son logos de firma: se ignoran
const CL_ESPERA_INDICE_MIN = 120;      // minutos que se espera a que Drive lea el texto del PDF

/** Ejecútala una vez: deja a organizarTodo corriendo cada 10 minutos (y quita el disparador viejo). */
function instalarDisparador(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    const fn = t.getHandlerFunction();
    if(fn === 'organizar' || fn === 'organizarTodo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('organizarTodo').timeBased().everyMinutes(10).create();
  Logger.log('Listo: organizarTodo corre cada 10 minutos.');
}

// ---------------------------------------------------------------------------
// Registro de lo pedido en cada correo (lo manda la app al generar el correo)
// ---------------------------------------------------------------------------
/** Cuerpo: { clave, accion:'solicitud', id, pry, nomPry, items:[{ n, oc, req, codigo, producto }] } */
function guardarSolicitud(p){
  const out = function(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); };
  if(!p.id || !p.pry || !p.items || !p.items.length) return out({ ok:false, error:'solicitud incompleta' });
  const prev = clLeerSol(p.id), items = (prev ? prev.items : []).slice();
  p.items.forEach(function(x){
    const k = (x.req || '') + '|' + (x.codigo || '');
    const i = items.findIndex(function(y){ return (y.req || '') + '|' + (y.codigo || '') === k; });
    const nuevo = { n:+x.n || 0, oc:String(x.oc || ''), req:String(x.req || ''), codigo:String(x.codigo || ''), producto:String(x.producto || '').slice(0, 90) };
    if(i === -1) items.push(nuevo); else items[i] = nuevo;
  });
  clGuardarSol({ id:String(p.id), pry:String(p.pry), nomPry:String(p.nomPry || ''), t:Date.now(), items:items });
  return out({ ok:true, id:p.id, items:items.length });
}
function clGuardarSol(s){
  const P = PropertiesService.getScriptProperties(), txt = JSON.stringify(s), TROZO = 8000, base = 'sol|' + s.id;
  const viejos = P.getKeys().filter(function(k){ return k === base || k.indexOf(base + '|') === 0; });
  viejos.forEach(function(k){ P.deleteProperty(k); });
  const partes = Math.ceil(txt.length / TROZO);
  for(let i = 0; i < partes; i++) P.setProperty(base + '|' + i, txt.substr(i * TROZO, TROZO));
  P.setProperty(base, String(partes));
  clPodar(P);
}
function clLeerSol(id){
  const P = PropertiesService.getScriptProperties(), n = +P.getProperty('sol|' + id) || 0;
  if(!n) return null;
  let txt = ''; for(let i = 0; i < n; i++) txt += P.getProperty('sol|' + id + '|' + i) || '';
  try{ return JSON.parse(txt); }catch(e){ return null; }
}
function clSolsDePry(pry){
  const P = PropertiesService.getScriptProperties();
  return P.getKeys().filter(function(k){ return /^sol\|[^|]+$/.test(k); })
    .map(function(k){ return clLeerSol(k.split('|')[1]); }).filter(function(s){ return s && s.pry === pry; });
}
/** Borra los registros de más de 240 días para no llenar el almacén del script. */
function clPodar(P){
  const lim = Date.now() - 240 * 864e5;
  P.getKeys().filter(function(k){ return /^sol\|[^|]+$/.test(k); }).forEach(function(k){
    const s = clLeerSol(k.split('|')[1]); if(s && s.t && s.t < lim){
      P.getKeys().filter(function(x){ return x === k || x.indexOf(k + '|') === 0; }).forEach(function(x){ P.deleteProperty(x); }); } });
}

// ---------------------------------------------------------------------------
// organizarTodo — reemplaza a "organizar" (lo hace todo: un producto y varios)
// ---------------------------------------------------------------------------
function organizarTodo(){
  const candado = LockService.getScriptLock();
  if(!candado.tryLock(30000)) return;
  try{
    const raiz = DriveApp.getFolderById(CONFIG.CARPETA_RAIZ), cache = {};
    const carpetas = [raiz], pa = raiz.getParents(); while(pa.hasNext()) carpetas.push(pa.next());   // "Detalle Certificados" y "Certificados GFP"
    let movidos = 0, revisar = 0, quedan = 0;
    carpetas.forEach(function(c){
      const sueltos = [], it = c.getFiles();
      while(it.hasNext()) sueltos.push(it.next());
      sueltos.forEach(function(f){
        try{
          const r = clArchivar(f, c, raiz, cache);
          if(r === 'ok') movidos++; else if(r === 'revisar') revisar++; else if(r === 'queda') quedan++;
        }catch(err){ Logger.log('Error con ' + f.getName() + ': ' + err); }
      });
    });
    Logger.log('Archivados: ' + movidos + ' · para revisar: ' + revisar + ' · sin ubicar: ' + quedan);
  } finally { candado.releaseLock(); }
}

/** Devuelve 'ok' (archivado), 'revisar', 'queda' (no se pudo ubicar), 'esperar' o 'ignorar'. */
function clArchivar(f, carpeta, raiz, cache){
  const nombre = f.getName(), mime = f.getMimeType();
  if(/^REVISAR/i.test(nombre)) return 'ignorar';
  const esPdf = /pdf/i.test(mime) || /\.pdf$/i.test(nombre), esImg = /^image\//i.test(mime);
  if(!esPdf && !(esImg && f.getSize() >= CL_MIN_IMG)) return 'ignorar';

  // 1) Un producto: [ref PRY|Req|Código]
  const r1 = nombre.match(/\[ref\s+([^|\]]+)\|([^|\]]*)\|([^|\]]+)\]/i);
  // 2) Varios productos: [ref PRY|SOL XXXXXX]
  const r2 = nombre.match(/\[ref\s+([^|\]]+)\|\s*SOL\s+([A-Z0-9]+)\s*\]/i);
  // 3) Sin referencia pero con el PRY en el asunto (correos enviados antes de este cambio)
  const r3 = nombre.match(/\b(PR[YD]\d{4}-\d{7})\b/i);
  if(!r1 && !r2 && !r3) return 'queda';
  const pry = (r1 || r2 || r3)[1].trim().toUpperCase();

  let sols = r2 ? [clLeerSol(r2[2].toUpperCase())].filter(Boolean) : [];
  if(!sols.length) sols = clSolsDePry(pry);
  const items = clUnir(sols), nomPry = (sols[0] && sols[0].nomPry) || '';
  const orig = clNombreOriginal(nombre, nomPry);
  let tipo = clTipo(orig);

  let destino = [];
  if(r1){
    const req = r1[2].trim(), cod = r1[3].trim();
    destino = items.filter(function(x){ return x.codigo === cod && (!req || x.req === req); });
    if(!destino.length) destino = [{ req:req, codigo:cod, producto:'', oc:'' }];
    if(!tipo) tipo = 'cert';
  }else{
    if(!items.length) return 'queda';
    const m = clEmparejar(orig, tipo, items);
    tipo = m.tipo;
    // Certificado: lo que diga el PDF por dentro manda sobre el nombre.
    if(tipo === 'cert'){
      const ocs = [...new Set(items.map(function(x){ return x.oc; }).filter(Boolean))];
      const enTexto = ocs.filter(function(oc){ return clTextoTiene(f, carpeta, oc); });
      if(enTexto.length){
        const ocNombre = m.items.map(function(x){ return x.oc; });
        if(m.porOc && ocNombre.length && !ocNombre.some(function(oc){ return enTexto.indexOf(oc) !== -1; })){
          f.setName(clLimpiar('REVISAR — dice OC ' + enTexto.join(', ') + ' — ' + orig, 200));
          return 'revisar';
        }
        if(!m.porNumero) m.items = items.filter(function(x){ return enTexto.indexOf(x.oc) !== -1; });
      }else if(m.porOc && (Date.now() - f.getDateCreated().getTime()) < CL_ESPERA_INDICE_MIN * 60000){
        return 'esperar';   // Drive todavía no leyó el texto del PDF: se revisa en la próxima pasada
      }
    }
    destino = m.items;
  }
  if(!destino.length) return 'queda';

  const final = clLimpiar(CL_ETQ[tipo] + ' — ' + orig, 200);
  const carpetasDestino = [], vistos = {};
  destino.forEach(function(d){
    const c = clCarpetaProducto(raiz, cache, pry, nomPry, d);
    if(!vistos[c.getId()]){ vistos[c.getId()] = 1; carpetasDestino.push(c); }
  });
  f.setName(final);
  const faltan = carpetasDestino.filter(function(c){ return !clYaEsta(c, final, f.getSize()); });   // donde ya está igual, no se repite
  if(!faltan.length){ f.setTrashed(true); return 'ok'; }
  faltan.slice(1).forEach(function(c){ f.makeCopy(final, c); });
  f.moveTo(faltan[0]);
  return 'ok';
}

// ---------------------------------------------------------------------------
// Lógica de emparejar (sin Drive: se puede probar sola)
// ---------------------------------------------------------------------------
function clUnir(sols){
  const out = [], k = {};
  sols.forEach(function(s){ (s.items || []).forEach(function(x){ const c = x.req + '|' + x.codigo; if(!k[c]){ k[c] = 1; out.push(x); } }); });
  return out;
}
/** Quita el asunto que Power Automate pega delante del nombre del adjunto. */
function clNombreOriginal(nombre, nomPry){
  let s = String(nombre);
  const cierre = s.lastIndexOf(']');
  if(cierre !== -1 && cierre < s.length - 1) s = s.substring(cierre + 1);
  else if(s.indexOf(' ~ ') !== -1) s = s.split(' ~ ').pop();
  else if(nomPry && clNorm(s).indexOf(clNorm(nomPry)) !== -1){
    const N = clNorm(s), i = N.indexOf(clNorm(nomPry)) + clNorm(nomPry).length; s = s.substring(i);
  }else{
    const m = s.match(/^(?:(?:RE|RV|FW|FWD|RES)\s*:\s*)*Certificados de calidad.*?PR[YD]\d{4}-\d{7}\s*[—-]\s*/i);
    if(m) s = s.substring(m[0].length);
  }
  s = s.replace(/^[\s~|.\-—]+/, '').trim();
  return s || 'documento.pdf';
}
function clTipo(orig){
  const n = clNorm(orig);
  if(/FICHA|\bFT\b|DATASHEET|DATA SHEET|HOJA DE DATOS|HOJA TECNICA|ESPECIFICACION/.test(n)) return 'ficha';
  if(/CERTIFICAD|CALIDAD|\bCERT\b|\bCOC\b|CONFORMIDAD/.test(n)) return 'cert';
  return '';
}
/**
 * Elige a qué productos va un archivo.
 *   orig: nombre original del adjunto · tipo: 'cert' | 'ficha' | '' · items: lo pedido.
 * Devuelve { tipo, items, porNumero, porOc }.
 */
function clEmparejar(orig, tipo, items){
  const N = clNorm(orig);
  // a) "3 - Certificado.pdf" / "1-3 Certificado.pdf" / "2,4 Ficha.pdf": el N° de la tabla
  const mn = N.match(/^\s*(\d{1,3}(?:\s*(?:[-,Y&]|AL)\s*\d{1,3})*)\s*[-_.)\s](?=[\s\-_]*[A-Z])/);
  if(mn){
    const nums = [], partes = mn[1].split(/\s*(?:,|Y|&)\s*/);
    partes.forEach(function(p){ const r = p.split(/\s*(?:-|AL)\s*/).map(Number);
      if(r.length === 2 && r[1] > r[0] && r[1] - r[0] < 60){ for(let i = r[0]; i <= r[1]; i++) nums.push(i); } else r.forEach(function(x){ nums.push(x); }); });
    const it = items.filter(function(x){ return nums.indexOf(+x.n) !== -1; });
    if(it.length && nums.every(function(x){ return x >= 1 && x <= 999; })) return { tipo:tipo || (clTieneOc(N, items).length ? 'cert' : 'ficha'), items:it, porNumero:true, porOc:false };
  }
  // b) N° de OC en el nombre
  const porOc = clTieneOc(N, items);
  if(porOc.length && tipo !== 'ficha') return { tipo:'cert', items:porOc, porNumero:false, porOc:true };
  if(!tipo) tipo = porOc.length ? 'cert' : 'ficha';
  // c) Modelo / descripción del producto
  const A = clAlnum(N.replace(/(\d)[.,](\d)/g, '$1$2'));
  const tokNom = clTokens(N);
  let mejor = 0, elegidos = [];
  items.forEach(function(x){
    const D = clNorm(x.producto), DA = clAlnum(D.replace(/(\d)[.,](\d)/g, '$1$2'));
    let s = 0;
    tokNom.codigos.forEach(function(t){ if(DA.indexOf(t) !== -1) s += 3; });
    clTokens(D).codigos.forEach(function(t){ if(t.length >= 5 && A.indexOf(t) !== -1 && tokNom.codigos.indexOf(t) === -1) s += 3; });
    tokNom.palabras.forEach(function(w){ if(D.indexOf(w) !== -1) s += 1; });
    if(s > mejor){ mejor = s; elegidos = [x]; } else if(s === mejor && s > 0) elegidos.push(x);
  });
  if(mejor >= 4) return { tipo:tipo, items:elegidos, porNumero:false, porOc:false };
  // Una ficha que no se pudo ubicar por modelo pero trae la OC en el nombre
  if(porOc.length) return { tipo:tipo, items:porOc, porNumero:false, porOc:true };
  return { tipo:tipo, items:[], porNumero:false, porOc:false };
}
function clTieneOc(N, items){
  const nums = (N.match(/\d{4,}/g) || []).map(function(x){ return String(+x); });
  const full = N.match(/\d{4}-\d{7}/g) || [];
  return items.filter(function(x){ if(!x.oc) return false; const oc = String(x.oc), tail = String(+(oc.split('-').pop() || 0));
    return full.indexOf(oc) !== -1 || (tail.length >= 4 && nums.indexOf(tail) !== -1); });
}
function clTokens(N){
  const crudo = N.replace(/(\d)[.,](\d)/g, '$1$2').split(/[^A-Z0-9]+/).filter(Boolean);
  const codigos = [], palabras = [];
  crudo.forEach(function(t){
    if(/[A-Z]/.test(t) && /\d/.test(t) && t.length >= 4) codigos.push(t);
    else if(/^[A-Z]+$/.test(t) && t.length >= 5 && CL_COMUNES.indexOf(t) === -1) palabras.push(t);
  });
  return { codigos:codigos, palabras:palabras };
}
const CL_COMUNES = ['CERTIFICADO', 'CALIDAD', 'FICHA', 'TECNICA', 'FRIOPACKING', 'FRIOPAKING', 'PRODUCTO', 'PARA', 'CON', 'DATOS', 'MATERIAL', 'DOCUMENTO', 'GENERAL'];
function clNorm(s){ return clSinAcentos(String(s || '').toUpperCase()).replace(/\.PDF$|\.JPE?G$|\.PNG$/, ''); }
function clAlnum(s){ return String(s).replace(/[^A-Z0-9]/g, ''); }
function clSinAcentos(s){ return String(s).replace(/[ÁÀÄÂ]/g, 'A').replace(/[ÉÈËÊ]/g, 'E').replace(/[ÍÌÏÎ]/g, 'I').replace(/[ÓÒÖÔ]/g, 'O').replace(/[ÚÙÜÛ]/g, 'U').replace(/Ñ/g, 'N'); }
function clLimpiar(t, largo){
  let s = String(t || '').replace(/[\/\\:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();
  if(largo && s.length > largo) s = s.substring(0, largo).trim();
  return s;
}

// ---------------------------------------------------------------------------
// Drive
// ---------------------------------------------------------------------------
/** ¿El texto del PDF (lo que Drive lee) contiene esa OC? */
function clTextoTiene(f, carpeta, oc){
  try{
    const it = DriveApp.searchFiles("'" + carpeta.getId() + "' in parents and fullText contains '" + String(oc).replace(/'/g, '') + "'");
    while(it.hasNext()) if(it.next().getId() === f.getId()) return true;
  }catch(err){}
  return false;
}
/** PRY — CLIENTE / Req. / Código — PRODUCTO, reusando las carpetas que ya existen (no duplica). */
function clCarpetaProducto(raiz, cache, pry, nomPry, d){
  const cPry = clHija(raiz, cache, pry, clLimpiar(pry + (nomPry ? CONFIG.SEP + String(nomPry).toUpperCase() : ''), 120));
  const req = clLimpiar(d.req || 'SIN REQ', 60);
  const cReq = clHija(cPry, cache, req, req);
  const desc = String(d.producto || '').toUpperCase(), corta = desc.length > (CONFIG.MAX_DESC || 60) ? desc.substring(0, CONFIG.MAX_DESC || 60).trim() + '…' : desc;
  return clHija(cReq, cache, d.codigo, clLimpiar(d.codigo + (desc ? CONFIG.SEP + corta : ''), 120));
}
/** Carpeta hija cuyo nombre es "clave" o empieza con "clave — "; si no hay, la crea con "nombreNuevo". */
function clHija(padre, cache, clave, nombreNuevo){
  const k = padre.getId() + '|' + clave;
  if(cache[k]) return cache[k];
  const it = padre.getFolders();
  while(it.hasNext()){
    const c = it.next(), n = c.getName();
    if(n === clave || n.indexOf(clave + CONFIG.SEP) === 0 || n.indexOf(clave + ' - ') === 0){ cache[k] = c; return c; }
  }
  cache[k] = padre.createFolder(nombreNuevo);
  return cache[k];
}
function clYaEsta(carpeta, nombre, peso){
  const it = carpeta.getFilesByName(nombre);
  while(it.hasNext()) if(it.next().getSize() === peso) return true;
  return false;
}
