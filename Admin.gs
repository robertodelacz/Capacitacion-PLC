/**
 * Admin.gs
 * Herramientas de administrador — no las usan los colaboradores, son para
 * correrlas manualmente tú desde el editor de Apps Script cuando haga falta.
 */

/**
 * Genera la constancia de alguien que ya quedó como aprobado en Intentos
 * pero que, por algún error en ese momento (ej. plantilla mal configurada),
 * se quedó sin el PDF. Cambia CORREO_A_REPARAR por el correo real, elige
 * esta función en el selector de arriba del editor, y dale Ejecutar.
 */
function regenerarConstanciaFaltante() {
  const correo = 'CORREO_A_REPARAR@cualli.mx'; // <-- cámbialo antes de ejecutar

  const config = obtenerConfig();
  const hojaIntentos = SpreadsheetApp.getActive().getSheetByName(SHEET_INTENTOS);
  const datos = hojaIntentos.getDataRange().getValues();
  let calificacion = null;
  let total = null;

  for (let i = 1; i < datos.length; i++) {
    const [, correoFila, campanaId, , preguntaIds, , calif, aprobadoFila] = datos[i];
    const esteAprobado = aprobadoFila === true || String(aprobadoFila).toUpperCase() === 'TRUE';
    if (String(correoFila).toLowerCase() === correo.toLowerCase() && campanaId === config.campanaId && esteAprobado) {
      calificacion = calif;
      total = String(preguntaIds).split(',').length;
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

  const resultado = generarConstancia(correo, nombre, puesto, calificacion, total);
  const mensaje = 'Constancia generada para ' + correo + ' — folio ' + resultado.folio;
  try {
    SpreadsheetApp.getUi().alert(mensaje);
  } catch (e) {
    Logger.log(mensaje);
  }
}
