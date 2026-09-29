/** Mi Gestor Personal — Google Apps Script vinculado a una hoja de cálculo.
 * Ejecutar instalarGestor() una vez y desplegar como aplicación web.
 * Registra gastos e ingresos escritos en lenguaje simple; el panel es privado.
 */
const MG = {
  HOJA: 'Movimientos',
  CATS: ['Supermercado','Comida','Transporte','Casa','Servicios','Salud','Ropa','Ocio','Viajes','Educación','Sueldo','Otros']
};

function instalarGestor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Abrí Apps Script desde una hoja de cálculo.');
  let sh = ss.getSheetByName(MG.HOJA);
  if (!sh) sh = ss.insertSheet(MG.HOJA);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['ID','Fecha','Tipo','Monto','Moneda','Categoría','Descripción','Cuenta','Origen','Creado']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,10).setFontWeight('bold').setBackground('#e9f3ee');
    sh.getRange('B:B').setNumberFormat('dd/mm/yyyy');
    sh.getRange('D:D').setNumberFormat('#,##0.00');
    sh.autoResizeColumns(1,10);
  }
  PropertiesService.getScriptProperties().setProperty('MG_SHEET_ID', ss.getId());
  asegurarColumnas_(sh);
  return 'Listo. Desplegá la aplicación web con acceso limitado a tu cuenta.';
}

function hoja_() {
  const id = PropertiesService.getScriptProperties().getProperty('MG_SHEET_ID');
  if (!id) throw new Error('Primero ejecutá instalarGestor().');
  return SpreadsheetApp.openById(id).getSheetByName(MG.HOJA);
}

function asegurarColumnas_(sh) {
  const extras = ['Monto MXN','Tipo de cambio MXN','Fecha cotización','Fuente cotización','Texto ticket'];
  if (sh.getMaxColumns() < 15) sh.insertColumnsAfter(sh.getMaxColumns(), 15-sh.getMaxColumns());
  extras.forEach((h,i) => sh.getRange(1,11+i).setValue(h));
  sh.getRange('K:L').setNumberFormat('#,##0.00');
}

/** Cotización de referencia del día del movimiento; puede ser el último día hábil publicado. */
function cotizacionMXN_(moneda, fecha) {
  if (moneda === 'MXN') return {rate:1,date:fecha,source:'MXN'};
  const cache = CacheService.getScriptCache();
  const key = 'fx_'+moneda+'_'+fecha, cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  const url = 'https://api.frankfurter.dev/v2/rate/'+moneda+'/MXN?date='+fecha;
  let response;
  try { response = UrlFetchApp.fetch(url,{muteHttpExceptions:true}); }
  catch(e) { throw new Error('No pude consultar el tipo de cambio. Probá de nuevo más tarde.'); }
  if (response.getResponseCode() !== 200) throw new Error('No hay cotización '+moneda+'/MXN para esta fecha. No se guardó el movimiento.');
  const d = JSON.parse(response.getContentText());
  if (!(Number(d.rate) > 0) || !d.date) throw new Error('La respuesta del tipo de cambio es inválida. No se guardó el movimiento.');
  const result={rate:Number(d.rate),date:String(d.date),source:'Frankfurter (referencia)'};
  cache.put(key,JSON.stringify(result),21600);
  return result;
}

/** Un problema temporal del proveedor no debe impedir registrar el gasto. */
function cotizacionDisponible_(moneda, fecha) {
  try { return cotizacionMXN_(moneda, fecha); }
  catch (e) { return null; }
}

/** Ejecutar manualmente una vez para convertir los registros anteriores, hasta 100 por ejecución. */
function actualizarConversionesExistentes() {
  const sh=hoja_(); asegurarColumnas_(sh);
  const rows=sh.getDataRange().getValues(), zona=Session.getScriptTimeZone();
  let actualizados=0;
  for(let i=1;i<rows.length && actualizados<100;i++) {
    const r=rows[i];
    if (r[10] !== '' || !(r[1] instanceof Date) || !(Number(r[3]) > 0) || !['MXN','ARS','UYU','USD'].includes(r[4])) continue;
    const fecha=Utilities.formatDate(r[1],zona,'yyyy-MM-dd');
    const c=cotizacionDisponible_(r[4],fecha);
    if (!c) continue;
    sh.getRange(i+1,11,1,4).setValues([[Math.round(Number(r[3])*c.rate*100)/100,c.rate,c.date,c.source]]);
    actualizados++;
  }
  return actualizados+' movimientos anteriores convertidos. Si faltan, ejecutá de nuevo.';
}

function parsearMovimiento(texto) {
  const s = String(texto || '').trim();
  if (!s) throw new Error('Escribí el movimiento. Ejemplo: gasté 250 pesos en supermercado');
  const m = s.match(/(?:\$|USD\s*|MXN\s*|ARS\s*|UYU\s*)?\s*(\d[\d.,]*)(?:\s*(?:pesos|dólares|dolares|usd|mxn|ars|uyu))?/i);
  if (!m) throw new Error('No encontré un monto. Ejemplo: gasté 250 pesos en supermercado');
  let numero = m[1];
  if (numero.includes(',') && numero.includes('.')) {
    numero = numero.lastIndexOf(',') > numero.lastIndexOf('.') ? numero.replace(/\./g,'').replace(',','.') : numero.replace(/,/g,'');
  } else if (/^\d{1,3}([.,]\d{3})+$/.test(numero)) {
    numero = numero.replace(/[.,]/g,'');
  } else numero = numero.replace(',','.');
  const monto = Number(numero);
  if (!(monto > 0) || !isFinite(monto)) throw new Error('Monto inválido.');
  const tipo = /(?:^|\s)(?:cobr[eé]|ingres[oé]|recib[ií]|me pagaron|sueldo|venta|reembolso)(?=\s|$)/i.test(s) ? 'Ingreso' : 'Gasto';
  const moneda = /\b(usd|d[oó]lares?)\b/i.test(s) ? 'USD' : /\b(ars|argentinos?)\b/i.test(s) ? 'ARS' : /\b(uyu|uruguayos?)\b/i.test(s) ? 'UYU' : 'MXN';
  const reglas = [
    [/super|mercado|despensa|verdura|leche/i,'Supermercado'],
    [/restaurante|caf[eé]|almuerzo|cena|delivery/i,'Comida'],
    [/uber|taxi|nafta|gasolina|transporte|estacionamiento/i,'Transporte'],
    [/renta|alquiler|mueble|ferreter/i,'Casa'],
    [/luz|agua|gas|internet|tel[eé]fono/i,'Servicios'],
    [/m[eé]dic|farmacia|consulta|medicamento/i,'Salud'],
    [/ropa|zapatos/i,'Ropa'],[/cine|juego|concierto/i,'Ocio'],
    [/vuelo|hotel|viaje/i,'Viajes'],[/curso|libro|escuela/i,'Educación'],
    [/sueldo|salario|n[oó]mina/i,'Sueldo']
  ];
  const categoria = (reglas.find(([re]) => re.test(s)) || [null,'Otros'])[1];
  return {tipo, monto, moneda, categoria, descripcion:s, cuenta:'Efectivo'};
}

function guardarMovimiento(datos) {
  if (!datos || typeof datos !== 'object') throw new Error('Faltan datos.');
  const tipo = String(datos.tipo || 'Gasto');
  if (!['Gasto','Ingreso'].includes(tipo)) throw new Error('Tipo inválido.');
  const monto = Number(datos.monto);
  if (!(monto > 0) || !isFinite(monto)) throw new Error('Monto inválido.');
  const moneda = String(datos.moneda || 'MXN').toUpperCase();
  if (!['MXN','ARS','UYU','USD'].includes(moneda)) throw new Error('Moneda inválida.');
  const fechaTexto = datos.fecha || Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fechaTexto)) throw new Error('Fecha inválida.');
  const fecha = new Date(fechaTexto + 'T12:00:00');
  if (isNaN(fecha.getTime()) || Utilities.formatDate(fecha,Session.getScriptTimeZone(),'yyyy-MM-dd') !== fechaTexto) throw new Error('Fecha inválida.');
  if (fechaTexto > Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd')) throw new Error('La fecha no puede ser futura.');
  const cotizacion = cotizacionDisponible_(moneda, fechaTexto);
  const montoMXN = cotizacion ? Math.round(monto*cotizacion.rate*100)/100 : null;
  const categoria = String(datos.categoria || 'Otros').slice(0,80);
  const descripcion = String(datos.descripcion || '').slice(0,500);
  const cuenta = String(datos.cuenta || 'Efectivo').slice(0,80);
  const textoTicket = String(datos.textoTicket || '').slice(0,30000);
  // Evita que texto ingresado por usuarios se interprete como fórmula de Sheets.
  const seguro = v => /^[=+@\-]/.test(v) ? "'" + v : v;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh=hoja_(); asegurarColumnas_(sh);
    sh.appendRow([Utilities.getUuid(), fecha, tipo, monto, moneda,
      seguro(categoria), seguro(descripcion), seguro(cuenta), 'Web', new Date(),
      montoMXN === null ? '' : montoMXN,
      cotizacion ? cotizacion.rate : '', cotizacion ? cotizacion.date : '', cotizacion ? cotizacion.source : 'Pendiente',seguro(textoTicket)]);
  } finally { lock.releaseLock(); }
  return {mensaje:montoMXN === null ? 'Movimiento guardado. Conversión a MXN pendiente.' : 'Movimiento guardado',tipo,categoria,moneda,monto,montoMXN,fecha:fechaTexto,cotizacion,cuenta};
}

function guardarTexto(texto, fecha, monedaElegida, cuenta) { const datos = parsearMovimiento(texto); datos.fecha = fecha; if (monedaElegida && monedaElegida !== 'Auto') datos.moneda = monedaElegida; if (cuenta) datos.cuenta = cuenta; return guardarMovimiento(datos); }

/** Convierte temporalmente la foto a Google Docs para extraer texto por OCR. */
function leerTicket(formulario) {
  const foto = formulario && formulario.foto;
  if (!foto || typeof foto.getBytes !== 'function') throw new Error('Elegí una foto del ticket.');
  if (!['image/jpeg','image/png','image/gif','image/bmp'].includes(foto.getContentType())) throw new Error('Usá una imagen JPG o PNG.');
  if (foto.getBytes().length > 10*1024*1024) throw new Error('La imagen debe pesar menos de 10 MB.');
  if (typeof Drive === 'undefined') throw new Error('Falta activar el servicio avanzado Drive API en Apps Script (Servicios → + → Drive API).');
  let archivo;
  try {
    archivo = Drive.Files.create({name:'Ticket temporal',mimeType:'application/vnd.google-apps.document'},foto,{ocrLanguage:'es',fields:'id'});
    const texto = DocumentApp.openById(archivo.id).getBody().getText().trim();
    if (!texto) throw new Error('No encontré texto legible. Probá con otra foto.');
    return interpretarTicket_(texto);
  } finally { if (archivo && archivo.id) Drive.Files.remove(archivo.id); }
}

function interpretarTicket_(texto) {
  const lineas = texto.split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
  const importe = s => { const m=s.match(/(?:\$|UYU|MXN|ARS|USD)?\s*(\d[\d.,]*\d|\d)(?:\s*(?:UYU|MXN|ARS|USD))?\s*$/i); if(!m)return null; let n=m[1];if(n.includes(',')&&n.includes('.')) n=n.lastIndexOf(',')>n.lastIndexOf('.')?n.replace(/\./g,'').replace(',','.'):n.replace(/,/g,'');else if(/^\d{1,3}([.,]\d{3})+$/.test(n))n=n.replace(/[.,]/g,'');else n=n.replace(',','.');return Number(n); };
  const total = lineas.filter(x=>/\b(total(?:\s+a\s+pagar)?|importe\s+total|a\s+pagar)\b/i.test(x)).map(importe).filter(n=>n>0).pop();
  const fechas = texto.match(/\b(\d{1,2})[\/.-](\d{1,2})[\/.-](20\d{2})\b/);
  const fecha = fechas ? fechas[3]+'-'+fechas[2].padStart(2,'0')+'-'+fechas[1].padStart(2,'0') : '';
  const moneda = /\b(UYU|pesos\s+uruguayos)\b/i.test(texto)?'UYU':/\b(ARS|pesos\s+argentinos)\b/i.test(texto)?'ARS':/\b(USD|d[oó]lares)\b/i.test(texto)?'USD':'MXN';
  return {comercio:lineas[0]||'',fecha,monto:total||'',moneda,texto};
}

function guardarTicket(datos) {
  if (!datos || !datos.textoTicket) throw new Error('Primero leé la foto del ticket.');
  return guardarMovimiento({tipo:'Gasto',monto:datos.monto,moneda:datos.moneda,fecha:datos.fecha,categoria:datos.categoria,descripcion:datos.descripcion,cuenta:datos.cuenta,textoTicket:datos.textoTicket});
}

function eliminarMovimiento(id) {
  if (typeof id !== 'string' || !/^[0-9a-f-]{36}$/i.test(id)) throw new Error('Movimiento inválido.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = hoja_();
    const ultimo = sh.getLastRow();
    if (ultimo < 2) throw new Error('No encontré el movimiento.');
    const ids = sh.getRange(2,1,ultimo-1,1).getValues();
    const posicion = ids.findIndex(r => String(r[0]) === id);
    if (posicion < 0) throw new Error('No encontré el movimiento.');
    sh.deleteRow(posicion + 2);
  } finally { lock.releaseLock(); }
  return 'Movimiento eliminado.';
}

function datosPanel(modo, referencia) {
  const sh = hoja_(), values = sh.getDataRange().getValues().slice(1);
  const zona = Session.getScriptTimeZone();
  const fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(referencia)) ? referencia : Utilities.formatDate(new Date(), zona, 'yyyy-MM-dd');
  modo = ['Todos','Día','Mes','Año'].includes(modo) ? modo : 'Todos';
  const periodo = modo === 'Todos' ? '' : modo === 'Día' ? fecha : modo === 'Mes' ? fecha.slice(0,7) : fecha.slice(0,4);
  const movimientos = values.filter(r => r[1] instanceof Date && (!periodo || Utilities.formatDate(r[1], zona, 'yyyy-MM-dd').startsWith(periodo)))
    .map(r => ({id:String(r[0]),fecha:Utilities.formatDate(r[1],zona,'dd/MM/yyyy'), tipo:r[2],monto:Number(r[3]),moneda:r[4],categoria:r[5],descripcion:r[6],cuenta:r[7],montoMXN:typeof r[10]==='number'?r[10]:null,tasa:r[11],cotizacion:r[12],textoTicket:r[14]||''}))
    .reverse();
  const resumen = {}; const categoriasMXN = {}; let ingresosMXN=0,gastosMXN=0,pendientes=0;
  movimientos.forEach(r => {
    const k = r.moneda;
    if (!resumen[k]) resumen[k] = {ingresos:0,gastos:0,categorias:{}};
    const a = resumen[k];
    if (r.tipo === 'Ingreso') { a.ingresos += r.monto; if(r.montoMXN!==null) ingresosMXN+=r.montoMXN; }
    else { a.gastos += r.monto; a.categorias[r.categoria] = (a.categorias[r.categoria] || 0) + r.monto; if(r.montoMXN!==null){gastosMXN+=r.montoMXN;categoriasMXN[r.categoria]=(categoriasMXN[r.categoria]||0)+r.montoMXN;} }
    if(r.montoMXN===null) pendientes++;
  });
  return {periodo,modo,resumen,movimientos,totalMXN:{ingresos:ingresosMXN,gastos:gastosMXN,categorias:categoriasMXN,pendientes}};
}

function doGet() {
  return HtmlService.createHtmlOutput(PANEL_HTML_()).setTitle('Mi Gestor Personal')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function PANEL_HTML_() { return `<!doctype html><html lang="es"><head><meta charset="utf-8"><style>
body{font:16px system-ui;margin:0;background:#f7f8f5;color:#24332d}main{max-width:1000px;margin:auto;padding:20px}h1{color:#32624c}section{background:white;border-radius:15px;padding:18px;margin:16px 0;box-shadow:0 2px 10px #0001}input,select,button{font:inherit;padding:10px;border:1px solid #ccd5ce;border-radius:8px}input[type=text]{width:min(95%,540px)}button{background:#357458;color:white;border:0;cursor:pointer}.grid{display:flex;gap:12px;flex-wrap:wrap}.card{background:#eaf3ec;padding:14px;border-radius:12px;min-width:180px}table{width:100%;border-collapse:collapse}td,th{padding:8px;border-bottom:1px solid #eee;text-align:left}.scroll{overflow:auto}.bar{height:18px;background:#6b9e79;border-radius:6px}.chart-row{display:grid;grid-template-columns:minmax(90px,150px) 1fr;gap:10px;align-items:center;margin:12px 0}.chart-track{background:#edf3ee;border-radius:6px}.chart-label{overflow-wrap:anywhere}small{color:#666}
</style></head><body><main><h1>Mi Gestor Personal</h1><section><h2>Registrar movimiento</h2><p>Por ejemplo: “gasté 250 pesos en supermercado” o “cobré 800 USD por consulta”.</p><input id="texto" type="text" placeholder="Escribí tu gasto o ingreso" onkeydown="if(event.key==='Enter')guardar()"><button id="boton" onclick="guardar()">Guardar</button><p><label>Moneda <select id="moneda"><option value="Auto">Detectar del texto</option><option value="UYU">Pesos uruguayos</option><option value="MXN">Pesos mexicanos</option><option value="ARS">Pesos argentinos</option><option value="USD">Dólares</option></select></label></p><p><label>Fecha del movimiento <input id="fecha" type="date"></label></p><p><label>Cómo pagué <select id="cuenta"><option>AMEX AM MCA</option><option>AMEX AM DRC</option><option>AMEX AM ERCA</option><option>AMEX PLAT MCA</option><option>AMEX PLAT DRC</option><option>Tarjeta BBva</option><option>VISA INB</option><option>TRANS BBVa</option><option>TRANS Galicia</option><option selected>Efectivo</option></select></label></p><p id="estado" role="status"></p><div id="ultimo" class="card" style="display:none"></div><small>Revisá la categoría y la moneda en la tabla; la interpretación del texto es básica.</small></section><section><h2>Leer un ticket</h2><p>Sacá una foto o elegí una imagen. Revisá los datos antes de guardar.</p><form id="formTicket" onsubmit="event.preventDefault();leerFoto()"><input id="foto" name="foto" type="file" accept="image/jpeg,image/png" capture="environment" required><button id="botonLeer" type="submit">Leer ticket</button></form><p id="estadoTicket" role="status"></p><div id="borradorTicket" style="display:none"><p><label>Comercio o descripción <input id="ticketDescripcion" type="text"></label></p><p><label>Fecha <input id="ticketFecha" type="date"></label> <label>Total <input id="ticketMonto" type="number" min="0.01" step="0.01"></label></p><p><label>Moneda <select id="ticketMoneda"><option>UYU</option><option>MXN</option><option>ARS</option><option>USD</option></select></label> <label>Rubro <select id="ticketCategoria"><option>Supermercado</option><option>Comida</option><option>Transporte</option><option>Casa</option><option>Servicios</option><option>Salud</option><option>Ropa</option><option>Ocio</option><option>Viajes</option><option>Educación</option><option>Otros</option></select></label></p><p><label>Cómo pagué <select id="ticketCuenta"><option>AMEX AM MCA</option><option>AMEX AM DRC</option><option>AMEX AM ERCA</option><option>AMEX PLAT MCA</option><option>AMEX PLAT DRC</option><option>Tarjeta BBva</option><option>VISA INB</option><option>TRANS BBVa</option><option>TRANS Galicia</option><option selected>Efectivo</option></select></label></p><p><label>Texto leído del ticket<br><textarea id="ticketTexto" rows="9" style="width:95%"></textarea></label></p><button id="botonTicket" type="button" onclick="guardarFoto()">Guardar gasto del ticket</button></div></section><section><label>Ver <select id="modo" onchange="cambiarModo()"><option>Todos</option><option>Día</option><option>Mes</option><option>Año</option></select></label> <label id="etiquetaFiltro" style="display:none">Fecha <input id="filtro" type="date" onchange="cargar()"></label></section><section><h2>Mis movimientos</h2><p id="cantidad">Cargando movimientos…</p><div class="scroll"><table><thead><tr><th>Fecha</th><th>Tipo</th><th>Monto original</th><th>Equivalente MXN</th><th>Categoría</th><th>Cómo pagué</th><th>Descripción</th><th></th></tr></thead><tbody id="filas"></tbody></table></div><p><button onclick="csv()">Descargar CSV</button></p></section><section><h2>Gastos por rubro</h2><p>En pesos mexicanos para los movimientos mostrados, de mayor a menor.</p><div id="grafico"></div><div class="scroll"><table><thead><tr><th>Rubro</th><th>Gasto (MXN)</th><th>Porcentaje</th></tr></thead><tbody id="rubros"></tbody></table></div><p id="pendientes" style="display:none"><small></small></p></section></main><script>
let datos=null;const $=id=>document.getElementById(id);const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const hoy=new Date();const fechaLocal=[hoy.getFullYear(),String(hoy.getMonth()+1).padStart(2,'0'),String(hoy.getDate()).padStart(2,'0')].join('-');
$('fecha').value=fechaLocal;$('filtro').value=fechaLocal;cargar();
let guardando=false;
function guardar(){if(guardando)return;let t=$('texto').value;if(!t.trim()){$('estado').textContent='Escribí un movimiento.';return}guardando=true;$('boton').disabled=true;$('estado').textContent='Guardando…';google.script.run.withSuccessHandler(m=>{guardando=false;$('boton').disabled=false;$('estado').textContent=m.mensaje;$('ultimo').style.display='block';$('ultimo').textContent=m.tipo+': '+m.monto.toFixed(2)+' '+m.moneda+' · '+m.categoria+' · '+m.cuenta+(m.montoMXN===null?' · Conversión a MXN pendiente':' · '+m.montoMXN.toFixed(2)+' MXN (tasa '+m.cotizacion.rate+', fecha '+m.cotizacion.date+')');$('texto').value='';cargar();$('texto').focus()}).withFailureHandler(e=>{guardando=false;$('boton').disabled=false;$('estado').textContent=e.message}).guardarTexto(t,$('fecha').value,$('moneda').value,$('cuenta').value)}
function leerFoto(){let f=$('foto').files[0];if(!f)return;if(f.size>10*1024*1024){$('estadoTicket').textContent='La foto debe pesar menos de 10 MB.';return}$('botonLeer').disabled=true;$('estadoTicket').textContent='Leyendo ticket…';google.script.run.withSuccessHandler(d=>{$('botonLeer').disabled=false;$('borradorTicket').style.display='block';$('ticketDescripcion').value=d.comercio;$('ticketFecha').value=d.fecha||$('fecha').value;$('ticketMonto').value=d.monto;$('ticketMoneda').value=d.moneda;$('ticketTexto').value=d.texto;$('estadoTicket').textContent='Revisá el total, fecha y rubro antes de guardar.'}).withFailureHandler(e=>{$('botonLeer').disabled=false;$('estadoTicket').textContent=e.message}).leerTicket($('formTicket'))}
function guardarFoto(){const d={descripcion:$('ticketDescripcion').value,fecha:$('ticketFecha').value,monto:$('ticketMonto').value,moneda:$('ticketMoneda').value,categoria:$('ticketCategoria').value,cuenta:$('ticketCuenta').value,textoTicket:$('ticketTexto').value};if(!d.descripcion||!d.fecha||!(Number(d.monto)>0)){$('estadoTicket').textContent='Completá descripción, fecha y total.';return}$('botonTicket').disabled=true;$('estadoTicket').textContent='Guardando gasto…';google.script.run.withSuccessHandler(m=>{$('botonTicket').disabled=false;$('estadoTicket').textContent=m.mensaje;$('borradorTicket').style.display='none';$('formTicket').reset();cargar()}).withFailureHandler(e=>{$('botonTicket').disabled=false;$('estadoTicket').textContent=e.message}).guardarTicket(d)}
function cambiarModo(){$('etiquetaFiltro').style.display=$('modo').value==='Todos'?'none':'inline';cargar()}
function cargar(){google.script.run.withSuccessHandler(p=>{datos=p;try{render()}catch(e){$('filas').innerHTML='<tr><td colspan="8">No pude mostrar los movimientos: '+esc(e.message)+'</td></tr>'}}).withFailureHandler(e=>$('filas').innerHTML='<tr><td colspan="8">No pude cargar los movimientos: '+esc(e.message)+'</td></tr>').datosPanel($('modo').value,$('filtro').value)}
function render(){
 let p=datos, mx=p.totalMXN;
 $('cantidad').textContent=(p.movimientos||[]).length+' movimiento(s) guardado(s)';
 $('filas').innerHTML=(p.movimientos||[]).map(r=>'<tr><td>'+esc(r.fecha)+'</td><td>'+esc(r.tipo)+'</td><td>'+Number(r.monto).toFixed(2)+' '+esc(r.moneda)+'</td><td>'+(r.montoMXN==null||r.montoMXN===''?'Pendiente':Number(r.montoMXN).toFixed(2)+' MXN')+'</td><td>'+esc(r.categoria)+'</td><td>'+esc(r.cuenta||'Efectivo')+'</td><td>'+esc(r.descripcion)+(r.textoTicket?'<details><summary>Ver ticket leído</summary><pre style="white-space:pre-wrap">'+esc(r.textoTicket)+'</pre></details>':'')+'</td><td><button type="button" data-id="'+esc(r.id)+'" onclick="eliminar(this.dataset.id)">Eliminar</button></td></tr>').join('')||'<tr><td colspan="8">Todavía no hay movimientos para esta selección.</td></tr>';
 let rubros=Object.entries(mx.categorias).filter(([,n])=>n>0).sort((x,y)=>y[1]-x[1]);
 let mayor=rubros.length?rubros[0][1]:0;
 $('grafico').innerHTML=rubros.map(([c,n])=>'<div class="chart-row"><span class="chart-label">'+esc(c)+'</span><div class="chart-track"><div class="bar" role="img" aria-label="'+esc(c)+': '+n.toFixed(2)+' MXN" style="width:'+Math.max(1,Math.round(100*n/mayor))+'%"></div></div></div>').join('')||'Sin gastos en este período';
 $('rubros').innerHTML=rubros.map(([c,n])=>'<tr><td>'+esc(c)+'</td><td>'+n.toFixed(2)+'</td><td>'+(100*n/mx.gastos).toFixed(1)+'%</td></tr>').join('');
 $('pendientes').style.display=mx.pendientes?'block':'none';$('pendientes').firstElementChild.textContent=mx.pendientes+' movimiento(s) anterior(es) sin convertir; no están sumados.';
}
function eliminar(id){const r=(datos?.movimientos||[]).find(m=>m.id===id);if(!r)return;if(!confirm('¿Eliminar este movimiento? '+r.fecha+' · '+r.monto+' '+r.moneda+' · '+r.descripcion))return;$('estado').textContent='Eliminando…';google.script.run.withSuccessHandler(m=>{$('estado').textContent=m;$('ultimo').style.display='none';cargar()}).withFailureHandler(e=>$('estado').textContent=e.message).eliminarMovimiento(id)}
function csv(){if(!datos)return;let rows=[['Fecha','Tipo','Monto','Moneda','Equivalente MXN','Tasa MXN','Fecha cotización','Categoría','Descripción','Cuenta'],...datos.movimientos.map(r=>[r.fecha,r.tipo,r.monto,r.moneda,r.montoMXN??'',r.tasa??'',r.cotizacion??'',r.categoria,r.descripcion,r.cuenta])];let s=String.fromCharCode(65279)+rows.map(row=>row.map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(',')).join(String.fromCharCode(13,10));let a=document.createElement('a');a.href=URL.createObjectURL(new Blob([s],{type:'text/csv;charset=utf-8'}));a.download='movimientos-'+(datos.periodo||'todos')+'.csv';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
</script></body></html>`; }
