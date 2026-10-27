

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PLD-FT')
    .addItem('Actualizar seguimiento de examen', 'actualizarSeguimientoExamen')
    .addToUi();
}


function actualizarSeguimientoExamen() {
  const config = obtenerConfig();
  const hojaRoster = SpreadsheetApp.getActive().getSheetByName(SHEET_COLABORADORES);
  const roster = hojaRoster.getDataRange().getValues();
  if (roster.length < 2) return;

  const hojaIntentos = SpreadsheetApp.getActive().getSheetByName(SHEET_INTENTOS);
  const intentos = hojaIntentos.getDataRange().getValues();

  // Índice por correo, solo de intentos de la campaña vigente: cuántos
  // lleva, si alguno ya fue aprobado, su calificación, y la fecha del
  // intento más reciente.
  const porCorreo = {};
  for (let i = 1; i < intentos.length; i++) {
    const [fecha, correoFila, campanaId, , , , calificacion, aprobadoFila] = intentos[i];
    if (campanaId !== config.campanaId) continue;
    const correo = String(correoFila).trim().toLowerCase();
    if (!porCorreo[correo]) porCorreo[correo] = { intentos: 0, aprobado: false, calificacion: '', fecha: null };
    const registro = porCorreo[correo];
    registro.intentos++;
    registro.fecha = fecha;
    const esteAprobado = aprobadoFila === true || String(aprobadoFila).toUpperCase() === 'TRUE';
    if (esteAprobado) {
      registro.aprobado = true;
      registro.calificacion = calificacion;
    }
  }

  const COLOR = {
    'Aprobado': '#d9ead3',
    'En proceso': '#fff2cc',
    'Bloqueado (agotó intentos)': '#f4cccc',
    'Pendiente': '#f3f3f3',
    'Inactivo': '#ffffff'
  };

  const valores = [];
  const colores = [];
  for (let i = 1; i < roster.length; i++) {
    const [correo, , , activo] = roster[i];
    const estaActivo = activo === true || String(activo).toUpperCase() === 'TRUE';
    let fila;
    if (!correo || !estaActivo) {
      fila = ['Inactivo', '', '', ''];
    } else {
      const registro = porCorreo[String(correo).trim().toLowerCase()];
      if (!registro) {
        fila = ['Pendiente', 0, '', ''];
      } else if (registro.aprobado) {
        fila = ['Aprobado', registro.intentos, registro.calificacion, registro.fecha];
      } else if (registro.intentos >= config.maxIntentos) {
        fila = ['Bloqueado (agotó intentos)', registro.intentos, '', registro.fecha];
      } else {
        fila = ['En proceso', registro.intentos, '', registro.fecha];
      }
    }
    valores.push(fila);
    colores.push([COLOR[fila[0]] || '#ffffff', null, null, null]);
  }

  hojaRoster.getRange(1, 6, 1, 4).setValues([['Estado', 'Intentos', 'Calificación', 'Última fecha']]);
  hojaRoster.getRange(2, 6, valores.length, 4).setValues(valores);
  hojaRoster.getRange(2, 6, colores.length, 4).setBackgrounds(colores);

  const resumen = {};
  valores.forEach(f => { resumen[f[0]] = (resumen[f[0]] || 0) + 1; });
  const mensaje = `Campaña: ${config.campanaId}\n\nAprobados: ${resumen['Aprobado'] || 0}\nEn proceso: ${resumen['En proceso'] || 0}\nBloqueados: ${resumen['Bloqueado (agotó intentos)'] || 0}\nPendientes: ${resumen['Pendiente'] || 0}`;
  try {
    SpreadsheetApp.getUi().alert('Seguimiento actualizado', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log(mensaje);
  }
}

function regenerarConstanciaFaltante() {
  const correo = 'diegovillafranco@weiji.mx'; // <-- cámbialo antes de ejecutar

  const config = obtenerConfig();
  const hojaIntentos = SpreadsheetApp.getActive().getSheetByName(SHEET_INTENTOS);
  const datos = hojaIntentos.getDataRange().getValues();
  let calificacion = null;
  let total = null;
  let fechaOriginal = null;

  for (let i = 1; i < datos.length; i++) {
    const [fecha, correoFila, campanaId, , preguntaIds, , calif, aprobadoFila] = datos[i];
    const esteAprobado = aprobadoFila === true || String(aprobadoFila).toUpperCase() === 'TRUE';
    if (String(correoFila).toLowerCase() === correo.toLowerCase() && campanaId === config.campanaId && esteAprobado) {
      calificacion = calif;
      total = String(preguntaIds).split(',').length;
      fechaOriginal = fecha;
    }
  }

  if (calificacion === null) {
    throw new Error('No se encontró un intento aprobado para ese correo en la campaña activa (' + config.campanaId + ').');
  }

  const hojaRoster = SpreadsheetApp.getActive().getSheetByName(SHEET_COLABORADORES);
  const roster = hojaRoster.getDataRange().getValues();
  let nombre = '';
  let puesto = '';
  for (let i = 1; i < roster.length; i++) {
    if (String(roster[i][0]).toLowerCase() === correo.toLowerCase()) {
      nombre = roster[i][1];
      puesto = roster[i][2];
    }
  }
  if (!nombre) {
    throw new Error('Ese correo no está en ColaboradoresAutorizados, no tengo el nombre para la constancia.');
  }

  const anterior = obtenerConstanciaMasRecienteSilenciosa(correo);
  const resultado = generarConstancia(correo, nombre, puesto, calificacion, total, fechaOriginal);
  if (anterior && anterior.pdfFileId && anterior.pdfFileId !== resultado.pdfFileId) {
    try { DriveApp.getFileById(anterior.pdfFileId).setTrashed(true); } catch (e) { /* si ya no existe, no pasa nada */ }
  }

  const mensaje = 'Constancia generada para ' + correo + ' — folio ' + resultado.folio
    + (anterior ? ' (se reemplazó la anterior, folio ' + anterior.folio + ')' : '');
  try {
    SpreadsheetApp.getUi().alert(mensaje);
  } catch (e) {
    Logger.log(mensaje);
  }
}
