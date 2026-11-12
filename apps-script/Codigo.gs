/**
 * CONTROL DE GASTOS — Backend (Google Apps Script)
 * Versión 4: evita filas duplicadas.
 *
 * Cambios contra la v3:
 *  - agregarGasto: si ya existe una fila con el mismo timestamp, NO la vuelve a agregar
 *    (pasa cuando la app reintenta un guardado que sí había llegado).
 *  - LockService en agregar/editar/eliminar: dos guardados al mismo tiempo no se pisan.
 *  - Se quitaron cargarHistorialCompleto / cargarTodo / resetearTodo: BORRABAN toda la
 *    hoja y la reescribían solo con Ene–May. Si alguien las corría hoy se perdían
 *    los gastos de junio en adelante.
 *  - buscarDuplicados(): función manual que solo LISTA posibles duplicados en el registro.
 */

const SHEET_ID  = '1nvNcr7RIr9anDuQ6ysr0XgT48yldwXKFm1oIMsMGT8s';
const FOLDER_ID = '10QwbJizkIdYgh_vp3oTG_6gBJRPAJkuC';

const HOJA_GASTOS  = 'Gastos';
const HOJA_DOLARES = 'Dolares';

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    let result;
    if      (action === 'list')          result = listarGastos();
    else if (action === 'add')           result = conLock(() => agregarGasto(body));
    else if (action === 'update')        result = conLock(() => editarGasto(body));
    else if (action === 'delete')        result = conLock(() => eliminarGasto(body));
    else if (action === 'updateDollars') result = conLock(() => actualizarDolares(body));
    else result = { ok: false, error: 'Acción desconocida: ' + action };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, mensaje: 'API de Control de Gastos activa v4' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function conLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function buscarFilaPorTimestamp(hoja, ts) {
  const ultima = hoja.getLastRow();
  if (ultima < 2) return -1;
  const ids = hoja.getRange(2, 1, ultima - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(ts)) return i + 2;
  }
  return -1;
}

function listarGastos() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  inicializarHojas(ss);

  const hojaG = ss.getSheetByName(HOJA_GASTOS);
  const data = hojaG.getDataRange().getValues();

  const gastos = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    gastos.push({
      timestamp:   r[0],
      fecha:       formatearFecha(r[1]),
      categoria:   r[2],
      descripcion: r[3],
      monto:       Number(r[4]) || 0,
      usuario:     r[5] || '',
      foto:        r[6] || ''
    });
  }

  const hojaD = ss.getSheetByName(HOJA_DOLARES);
  const dData = hojaD.getDataRange().getValues();
  const dollars = {
    efectivo: Number(dData[1] && dData[1][1]) || 0,
    banco:    Number(dData[2] && dData[2][1]) || 0
  };

  return { ok: true, gastos: gastos, dollars: dollars };
}

function agregarGasto(body) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  inicializarHojas(ss);
  const hoja = ss.getSheetByName(HOJA_GASTOS);

  const ts = body.timestamp || Date.now();

  // Reintento de un gasto que ya se guardó: no lo duplicamos.
  if (buscarFilaPorTimestamp(hoja, ts) !== -1) {
    return { ok: true, duplicado: true };
  }

  let fotoUrl = '';
  if (body.fotoBase64) {
    fotoUrl = subirFoto(body.fotoBase64, body.fotoNombre || 'ticket.jpg', body.fecha, body.descripcion);
  }

  hoja.appendRow([
    ts,
    body.fecha,
    body.categoria,
    body.descripcion,
    body.monto,
    body.usuario || '',
    fotoUrl
  ]);

  return { ok: true, fotoUrl: fotoUrl };
}

function editarGasto(body) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const hoja = ss.getSheetByName(HOJA_GASTOS);
  const fila = buscarFilaPorTimestamp(hoja, body.timestamp);
  if (fila === -1) return { ok: false, error: 'Gasto no encontrado' };

  hoja.getRange(fila, 2, 1, 4).setValues([[body.fecha, body.categoria, body.descripcion, body.monto]]);
  if (body.fotoBase64) {
    const url = subirFoto(body.fotoBase64, body.fotoNombre || 'ticket.jpg', body.fecha, body.descripcion);
    hoja.getRange(fila, 7).setValue(url);
  }
  return { ok: true };
}

function eliminarGasto(body) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const hoja = ss.getSheetByName(HOJA_GASTOS);
  const fila = buscarFilaPorTimestamp(hoja, body.timestamp);
  if (fila === -1) return { ok: false, error: 'Gasto no encontrado' };
  hoja.deleteRow(fila);
  return { ok: true };
}

function subirFoto(base64, nombreOriginal, fecha, descripcion) {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const decoded = Utilities.base64Decode(base64);
  const ext = (nombreOriginal.split('.').pop() || 'jpg').toLowerCase();
  const safeDesc = String(descripcion || 'gasto').replace(/[^a-zA-Z0-9-_]/g, '_').substring(0, 30);
  const filename = `${fecha}_${safeDesc}_${Date.now()}.${ext}`;
  const blob = Utilities.newBlob(decoded, 'image/jpeg', filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w1200`;
}

function actualizarDolares(body) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  inicializarHojas(ss);
  const hoja = ss.getSheetByName(HOJA_DOLARES);
  hoja.getRange('B2').setValue(Number(body.efectivo) || 0);
  hoja.getRange('B3').setValue(Number(body.banco) || 0);
  hoja.getRange('B4').setValue(new Date());
  return { ok: true };
}

function inicializarHojas(ss) {
  let hojaG = ss.getSheetByName(HOJA_GASTOS);
  if (!hojaG) {
    hojaG = ss.insertSheet(HOJA_GASTOS);
    hojaG.appendRow(['Timestamp', 'Fecha', 'Categoría', 'Descripción', 'Monto', 'Usuario', 'Foto']);
    hojaG.getRange('A1:G1').setFontWeight('bold').setBackground('#1a2d80').setFontColor('#ffffff');
    hojaG.setFrozenRows(1);
    hojaG.setColumnWidth(1, 130);
    hojaG.setColumnWidth(2, 100);
    hojaG.setColumnWidth(3, 130);
    hojaG.setColumnWidth(4, 220);
    hojaG.setColumnWidth(5, 100);
    hojaG.setColumnWidth(6, 100);
    hojaG.setColumnWidth(7, 250);
  }

  let hojaD = ss.getSheetByName(HOJA_DOLARES);
  if (!hojaD) {
    hojaD = ss.insertSheet(HOJA_DOLARES);
    hojaD.getRange('A1:B1').setValues([['Concepto', 'Monto USD']]).setFontWeight('bold').setBackground('#1a2d80').setFontColor('#ffffff');
    hojaD.getRange('A2').setValue('Efectivo');
    hojaD.getRange('A3').setValue('Banco');
    hojaD.getRange('A4').setValue('Última actualización');
    hojaD.getRange('B2:B3').setValue(0);
    hojaD.setColumnWidth(1, 180);
    hojaD.setColumnWidth(2, 150);
  }
}

function formatearFecha(v) {
  if (!v) return '';
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(v);
}

// =====================================================
// FUNCIONES MANUALES (se corren desde el editor de Apps Script)
// =====================================================

function probarConexion() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('✓ Sheet conectado: ' + ss.getName());
    inicializarHojas(ss);
    Logger.log('✓ Hojas inicializadas correctamente');
    const folder = DriveApp.getFolderById(FOLDER_ID);
    Logger.log('✓ Carpeta de Drive conectada: ' + folder.getName());
    Logger.log('🎉 TODO LISTO. Puedes implementar el script.');
  } catch (err) {
    Logger.log('❌ ERROR: ' + err.toString());
  }
}

/**
 * 🔎 Solo REVISA (no borra nada): lista filas con la misma fecha, categoría,
 * descripción y monto. Revísalas a mano antes de eliminar.
 */
function buscarDuplicados() {
  const hoja = SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA_GASTOS);
  const data = hoja.getDataRange().getValues();
  const vistos = {};
  let n = 0;
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    const clave = [formatearFecha(r[1]), r[2], String(r[3]).trim().toLowerCase(), Number(r[4])].join('|');
    if (vistos[clave]) {
      Logger.log(`Posible duplicado: fila ${vistos[clave]} y fila ${i + 1} → ${clave}`);
      n++;
    } else {
      vistos[clave] = i + 1;
    }
  }
  Logger.log(n ? `${n} posibles duplicados.` : 'No hay filas duplicadas.');
}
