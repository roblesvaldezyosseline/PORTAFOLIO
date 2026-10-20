/**
 * Joyería – Automatización con Google Apps Script
 * Autora: Yosseline Robles Valdez
 *
 * Qué hace:
 *  1. Menú "Joyería" dentro de Google Sheets.
 *  2. generarReporte(): consolida KPIs de la hoja Pedidos en la hoja "Reporte".
 *  3. pedidosEnRiesgo(): lista los pedidos en producción que llegarán tarde.
 *  4. actualizarTipoCambio(): consume una API REST (JSON) y guarda el tipo de cambio USD→PEN.
 *  5. enviarReporte(): envía el resumen por correo.
 *  6. programarEnvioDiario(): crea un activador para que todo corra solo cada mañana.
 *
 * Requisito: importar datos_joyeria.xlsx a Google Sheets (Archivo > Importar).
 */

const CORREO_DESTINO = Session.getActiveUser().getEmail(); // cámbialo si quieres
const DIAS_ENVIO = 3; // envío promedio histórico: 2,5 días

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Joyería')
    .addItem('1. Generar reporte de KPIs', 'generarReporte')
    .addItem('2. Ver pedidos en riesgo', 'pedidosEnRiesgo')
    .addItem('3. Actualizar tipo de cambio (API)', 'actualizarTipoCambio')
    .addItem('4. Enviar reporte por correo', 'enviarReporte')
    .addSeparator()
    .addItem('Programar envío diario 8:00 a. m.', 'programarEnvioDiario')
    .addToUi();
}

/** Lee una hoja y devuelve una lista de objetos {columna: valor}. */
function leerHoja(nombre) {
  const datos = SpreadsheetApp.getActive().getSheetByName(nombre).getDataRange().getValues();
  const cab = datos.shift();
  return datos.map(fila => Object.fromEntries(cab.map((c, i) => [c, fila[i]])));
}

/** Escribe una tabla (con cabecera) en una hoja, creándola si no existe. */
function escribirHoja(nombre, filas) {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  hoja.clear();
  hoja.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
  hoja.getRange(1, 1, 1, filas[0].length).setFontWeight('bold').setBackground('#FBE0E9');
  hoja.autoResizeColumns(1, filas[0].length);
  return hoja;
}

function calcularKpis() {
  const pedidos = leerHoja('Pedidos');
  const validos = pedidos.filter(p => ['Entregado', 'En producción'].includes(p.estado));
  const entregados = pedidos.filter(p => p.estado === 'Entregado');
  const aTiempo = entregados.filter(p => new Date(p.fecha_entrega) <= new Date(p.fecha_prometida));
  const ventas = validos.reduce((s, p) => s + Number(p.total_soles), 0);
  const pct = (a, b) => b ? Math.round(a / b * 1000) / 10 : 0;

  const porCanal = {};
  validos.forEach(p => porCanal[p.canal_venta] = (porCanal[p.canal_venta] || 0) + Number(p.total_soles));

  return {
    ventas: Math.round(ventas),
    pedidos: validos.length,
    ticket: Math.round(ventas / validos.length),
    otif: pct(aTiempo.length, entregados.length),
    errorGrabado: pct(pedidos.filter(p => p.error_grabado === 'Sí').length, pedidos.length),
    cancelados: pct(pedidos.filter(p => p.estado === 'Cancelado').length, pedidos.length),
    porCanal: porCanal
  };
}

function generarReporte() {
  const k = calcularKpis();
  const filas = [
    ['Indicador', 'Valor'],
    ['Ventas netas (S/)', k.ventas],
    ['Pedidos válidos', k.pedidos],
    ['Ticket promedio (S/)', k.ticket],
    ['Entregas a tiempo (%)', k.otif],
    ['Error de grabado (%)', k.errorGrabado],
    ['Cancelación (%)', k.cancelados],
    ['', ''],
    ['Canal', 'Ventas (S/)']
  ].concat(Object.entries(k.porCanal)
    .sort((a, b) => b[1] - a[1])
    .map(([canal, v]) => [canal, Math.round(v)]));
  escribirHoja('Reporte', filas);
  SpreadsheetApp.getActive().toast('Reporte actualizado', 'Joyería', 4);
  return k;
}

function pedidosEnRiesgo() {
  const hoy = new Date();
  const dia = 24 * 60 * 60 * 1000;
  const riesgo = leerHoja('Pedidos')
    .filter(p => p.estado === 'En producción')
    .map(p => {
      const finTaller = new Date(new Date(p.fecha_pedido).getTime() + p.dias_produccion * dia);
      const llegada = new Date(Math.max(finTaller, hoy) + DIAS_ENVIO * dia);
      const retraso = Math.ceil((llegada - new Date(p.fecha_prometida)) / dia);
      return [p.id_pedido, p.id_cliente, p.fecha_prometida, llegada, retraso];
    })
    .filter(f => f[4] > 0);

  escribirHoja('Pedidos en riesgo',
    [['Pedido', 'Cliente', 'Fecha prometida', 'Llegada estimada', 'Días de retraso']].concat(riesgo));
  return riesgo;
}

/** Integración con API REST: tipo de cambio USD → PEN (respuesta JSON). */
function actualizarTipoCambio() {
  const resp = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/USD', { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('La API respondió ' + resp.getResponseCode());
  const json = JSON.parse(resp.getContentText());
  const tc = json.rates.PEN;
  const k = calcularKpis();
  escribirHoja('Tipo de cambio', [
    ['Fecha', 'USD → PEN', 'Ventas en S/', 'Ventas en USD'],
    [new Date(), tc, k.ventas, Math.round(k.ventas / tc)]
  ]);
  return tc;
}

function enviarReporte() {
  const k = generarReporte();
  const riesgo = pedidosEnRiesgo();
  const cuerpo =
    'Reporte diario – Joyería\n\n' +
    'Ventas netas: S/ ' + k.ventas + '\n' +
    'Ticket promedio: S/ ' + k.ticket + '\n' +
    'Entregas a tiempo: ' + k.otif + ' %\n' +
    'Error de grabado: ' + k.errorGrabado + ' %\n\n' +
    'Pedidos en riesgo de llegar tarde: ' + riesgo.length + '\n' +
    riesgo.map(r => '• ' + r[0] + ' (+' + r[4] + ' días)').join('\n') +
    '\n\nDetalle: ' + SpreadsheetApp.getActive().getUrl();
  MailApp.sendEmail(CORREO_DESTINO, 'Reporte diario – Joyería', cuerpo);
  SpreadsheetApp.getActive().toast('Correo enviado a ' + CORREO_DESTINO, 'Joyería', 5);
}

function programarEnvioDiario() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'enviarReporte')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('enviarReporte').timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getActive().toast('Listo: el reporte se enviará cada día a las 8:00 a. m.', 'Joyería', 5);
}
