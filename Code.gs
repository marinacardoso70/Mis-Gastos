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
  const extras = ['Monto MXN','Tipo de cambio MXN','Fecha cotización','Fuente cotización'];
  if (sh.getMaxColumns() < 14) sh.insertColumnsAfter(sh.getMaxColumns(), 14-sh.getMaxColumns());
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

/** Ejecutar manualmente una vez para convertir los registros anteriores, hasta 100 por ejecución. */
function actualizarConversionesExistentes() {
  const sh=hoja_(); asegurarColumnas_(sh);
  const rows=sh.getDataRange().getValues(), zona=Session.getScriptTimeZone();
  let actualizados=0;
  for(let i=1;i<rows.length && actualizados<100;i++) {
    const r=rows[i];
    if (r[10] !== '' || !(r[1] instanceof Date) || !(Number(r[3]) > 0) || !['MXN','ARS','UYU','USD'].includes(r[4])) continue;
    const fecha=Utilities.formatDate(r[1],zona,'yyyy-MM-dd');
    const c=cotizacionMXN_(r[4],fecha);
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
  const cotizacion = cotizacionMXN_(moneda, fechaTexto);
  const montoMXN = Math.round(monto*cotizacion.rate*100)/100;
  const categoria = String(datos.categoria || 'Otros').slice(0,80);
  const descripcion = String(datos.descripcion || '').slice(0,500);
  const cuenta = String(datos.cuenta || 'Efectivo').slice(0,80);
  // Evita que texto ingresado por usuarios se interprete como fórmula de Sheets.
  const seguro = v => /^[=+@\-]/.test(v) ? "'" + v : v;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh=hoja_(); asegurarColumnas_(sh);
    sh.appendRow([Utilities.getUuid(), fecha, tipo, monto, moneda,
      seguro(categoria), seguro(descripcion), seguro(cuenta), 'Web', new Date(),
      montoMXN, cotizacion.rate, cotizacion.date, cotizacion.source]);
  } finally { lock.releaseLock(); }
  return {mensaje:'Movimiento guardado',tipo,categoria,moneda,monto,montoMXN,fecha:fechaTexto,cotizacion};
}

function guardarTexto(texto, fecha, monedaElegida) { const datos = parsearMovimiento(texto); datos.fecha = fecha; if (monedaElegida && monedaElegida !== 'Auto') datos.moneda = monedaElegida; return guardarMovimiento(datos); }

function datosPanel(modo, referencia) {
  const sh = hoja_(), values = sh.getDataRange().getValues().slice(1);
  const zona = Session.getScriptTimeZone();
  const fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(referencia)) ? referencia : Utilities.formatDate(new Date(), zona, 'yyyy-MM-dd');
  modo = ['Día','Mes','Año'].includes(modo) ? modo : 'Día';
  const periodo = modo === 'Día' ? fecha : modo === 'Mes' ? fecha.slice(0,7) : fecha.slice(0,4);
  const movimientos = values.filter(r => r[1] instanceof Date && Utilities.formatDate(r[1], zona, 'yyyy-MM-dd').startsWith(periodo))
    .map(r => ({fecha:Utilities.formatDate(r[1],zona,'dd/MM/yyyy'), tipo:r[2],monto:Number(r[3]),moneda:r[4],categoria:r[5],descripcion:r[6],cuenta:r[7],montoMXN:typeof r[10]==='number'?r[10]:null,tasa:r[11],cotizacion:r[12]}))
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
body{font:16px system-ui;margin:0;background:#f7f8f5;color:#24332d}main{max-width:1000px;margin:auto;padding:20px}h1{color:#32624c}section{background:white;border-radius:15px;padding:18px;margin:16px 0;box-shadow:0 2px 10px #0001}input,select,button{font:inherit;padding:10px;border:1px solid #ccd5ce;border-radius:8px}input[type=text]{width:min(95%,540px)}button{background:#357458;color:white;border:0;cursor:pointer}.grid{display:flex;gap:12px;flex-wrap:wrap}.card{background:#eaf3ec;padding:14px;border-radius:12px;min-width:180px}table{width:100%;border-collapse:collapse}td,th{padding:8px;border-bottom:1px solid #eee;text-align:left}.scroll{overflow:auto}.bar{height:12px;background:#6b9e79;border-radius:6px}small{color:#666}
</style></head><body><main><h1>Mi Gestor Personal</h1><section><h2>Registrar movimiento</h2><p>Por ejemplo: “gasté 250 pesos en supermercado” o “cobré 800 USD por consulta”.</p><input id="texto" type="text" placeholder="Escribí tu gasto o ingreso" onkeydown="if(event.key==='Enter')guardar()"><button id="boton" onclick="guardar()">Guardar</button><p><label>Moneda <select id="moneda"><option value="Auto">Detectar del texto</option><option value="UYU">Pesos uruguayos</option><option value="MXN">Pesos mexicanos</option><option value="ARS">Pesos argentinos</option><option value="USD">Dólares</option></select></label></p><p><label>Fecha del movimiento <input id="fecha" type="date"></label></p><p id="estado" role="status"></p><div id="ultimo" class="card" style="display:none"></div><small>Revisá la categoría y la moneda en la tabla; la interpretación del texto es básica.</small></section><section><label>Ver <select id="modo" onchange="cargar()"><option>Día</option><option>Mes</option><option>Año</option></select></label> <label>Fecha <input id="filtro" type="date" onchange="cargar()"></label><div id="totalMXN" class="card"></div><div id="resumen" class="grid"></div></section><section><h2>En qué gasté (MXN)</h2><div id="categorias"></div></section><section><h2>Movimientos</h2><div class="scroll"><table><thead><tr><th>Fecha</th><th>Tipo</th><th>Monto original</th><th>Equivalente MXN</th><th>Categoría</th><th>Descripción</th></tr></thead><tbody id="filas"></tbody></table></div><p><button onclick="csv()">Descargar CSV</button></p></section></main><script>
let datos=null;const $=id=>document.getElementById(id);const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const hoy=new Date();const fechaLocal=[hoy.getFullYear(),String(hoy.getMonth()+1).padStart(2,'0'),String(hoy.getDate()).padStart(2,'0')].join('-');
$('fecha').value=fechaLocal;$('filtro').value=fechaLocal;cargar();
let guardando=false;
function guardar(){if(guardando)return;let t=$('texto').value;if(!t.trim()){$('estado').textContent='Escribí un movimiento.';return}guardando=true;$('boton').disabled=true;$('estado').textContent='Guardando…';google.script.run.withSuccessHandler(m=>{guardando=false;$('boton').disabled=false;$('estado').textContent=m.mensaje;$('ultimo').style.display='block';$('ultimo').textContent=m.tipo+': '+m.monto.toFixed(2)+' '+m.moneda+' · '+m.categoria+' · '+m.montoMXN.toFixed(2)+' MXN (tasa '+m.cotizacion.rate+', fecha '+m.cotizacion.date+')';$('texto').value='';$('filtro').value=$('fecha').value;cargar();$('texto').focus()}).withFailureHandler(e=>{guardando=false;$('boton').disabled=false;$('estado').textContent=e.message}).guardarTexto(t,$('fecha').value,$('moneda').value)}
function cargar(){google.script.run.withSuccessHandler(p=>{datos=p;render()}).withFailureHandler(e=>$('resumen').textContent=e.message).datosPanel($('modo').value,$('filtro').value)}
function render(){
 let p=datos, currencies=Object.keys(p.resumen), mx=p.totalMXN;
 $('totalMXN').innerHTML='<b>Total en pesos mexicanos</b><p>Ingresos: '+mx.ingresos.toFixed(2)+' MXN · Gastos: '+mx.gastos.toFixed(2)+' MXN · Saldo: '+(mx.ingresos-mx.gastos).toFixed(2)+' MXN</p>'+(mx.pendientes?'<small>'+mx.pendientes+' movimiento(s) anterior(es) sin convertir; no están sumados.</small>':'');
 $('resumen').innerHTML=currencies.length?currencies.map(k=>{let a=p.resumen[k];return '<div class="card"><b>'+esc(k)+'</b><p>Ingresos: '+a.ingresos.toFixed(2)+'</p><p>Gastos: '+a.gastos.toFixed(2)+'</p><b>Saldo: '+(a.ingresos-a.gastos).toFixed(2)+'</b></div>'}).join(''):'Sin movimientos en el período';
 $('categorias').innerHTML=Object.entries(mx.categorias).sort((x,y)=>y[1]-x[1]).map(([c,n])=>'<p>'+esc(c)+' · '+n.toFixed(2)+' MXN</p><div class="bar" style="width:'+Math.max(1,Math.round(100*n/mx.gastos))+'%"></div>').join('')||'Sin gastos convertidos en este período';
 $('filas').innerHTML=p.movimientos.map(r=>'<tr><td>'+esc(r.fecha)+'</td><td>'+esc(r.tipo)+'</td><td>'+r.monto.toFixed(2)+' '+esc(r.moneda)+'</td><td>'+(r.montoMXN===null?'Pendiente':r.montoMXN.toFixed(2)+' MXN')+'</td><td>'+esc(r.categoria)+'</td><td>'+esc(r.descripcion)+'</td></tr>').join('');
}
function csv(){if(!datos)return;let rows=[['Fecha','Tipo','Monto','Moneda','Equivalente MXN','Tasa MXN','Fecha cotización','Categoría','Descripción','Cuenta'],...datos.movimientos.map(r=>[r.fecha,r.tipo,r.monto,r.moneda,r.montoMXN??'',r.tasa??'',r.cotizacion??'',r.categoria,r.descripcion,r.cuenta])];let s='\ufeff'+rows.map(row=>row.map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(',')).join('\r\n');let a=document.createElement('a');a.href=URL.createObjectURL(new Blob([s],{type:'text/csv;charset=utf-8'}));a.download='movimientos-'+datos.periodo+'.csv';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
</script></body></html>`; }
