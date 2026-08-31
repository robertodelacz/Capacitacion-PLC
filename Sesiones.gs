/**
 * Sesiones.gs
 * Controla el examen "en curso": qué preguntas le tocaron a este intento,
 * en qué pregunta va, qué ha contestado, y desde cuándo corre el reloj.
 * El límite de tiempo se valida siempre aquí, en el servidor — el contador
 * que ve el colaborador en pantalla es solo referencia visual.
 */

function obtenerSesionActiva(correo, campanaId) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_SESIONES);
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (String(fila[0]).toLowerCase() === correo.toLowerCase() && fila[1] === campanaId) {
      return {
        correo: fila[0],
        campanaId: fila[1],
        preguntaIds: String(fila[2]).split(','),
        respuestas: JSON.parse(fila[3] || '{}'),
        indiceActual: Number(fila[4]),
        inicio: new Date(fila[5]),
        cambiosFoco: Number(fila[6]) || 0
      };
    }
  }
  return null;
}

function crearSesion(correo, campanaId, preguntaIds) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (obtenerSesionActiva(correo, campanaId)) {
      throw new Error('Ya tienes un examen en curso — recarga la página para continuarlo.');
    }
    const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_SESIONES);
    hoja.appendRow([correo, campanaId, preguntaIds.join(','), '{}', 0, new Date(), 0]);
  } finally {
    lock.releaseLock();
  }
}

function actualizarSesion(correo, campanaId, respuestas, indiceActual, cambiosFoco) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_SESIONES);
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).toLowerCase() === correo.toLowerCase() && datos[i][1] === campanaId) {
        hoja.getRange(i + 1, 4, 1, 2).setValues([[JSON.stringify(respuestas), indiceActual]]);
        hoja.getRange(i + 1, 7, 1, 1).setValue(cambiosFoco || 0);
        return;
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function eliminarSesion(correo, campanaId) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_SESIONES);
  const datos = hoja.getDataRange().getValues();
  for (let i = datos.length - 1; i >= 1; i--) {
    if (String(datos[i][0]).toLowerCase() === correo.toLowerCase() && datos[i][1] === campanaId) {
      hoja.deleteRow(i + 1);
    }
  }
}

function segundosRestantes(sesion, duracionMinutos) {
  const transcurridoMs = new Date().getTime() - sesion.inicio.getTime();
  const limiteMs = duracionMinutos * 60 * 1000;
  return Math.max(0, Math.round((limiteMs - transcurridoMs) / 1000));
}
