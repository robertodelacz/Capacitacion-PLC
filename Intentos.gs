/**
 * Intentos.gs
 * Determina si un colaborador puede presentar, ya está bloqueado (por
 * intentos agotados o por ya haber aprobado), o si la campaña no está
 * vigente. Todo el conteo vive en el servidor — nunca se confía en un
 * contador que pueda manipularse desde el navegador.
 */

function obtenerEstadoUsuario(correo) {
  const config = obtenerConfig();
  const ahora = new Date();

  if (ahora < config.ventanaInicio || ahora > config.ventanaFin) {
    return { estado: 'FUERA_DE_VENTANA', config };
  }

  const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_INTENTOS);
  const datos = hoja.getDataRange().getValues();
  let intentos = 0;
  let aprobado = false;
  let mejorCalificacion = '';
  let respuestasAprobado = null;

  for (let i = 1; i < datos.length; i++) {
    const [, correoFila, campanaId, , , respuestasJSON, calificacion, aprobadoFila] = datos[i];
    if (String(correoFila).toLowerCase() === correo.toLowerCase() && campanaId === config.campanaId) {
      intentos++;
      const esteAprobado = aprobadoFila === true || String(aprobadoFila).toUpperCase() === 'TRUE';
      if (esteAprobado) {
        aprobado = true;
        mejorCalificacion = calificacion;
        respuestasAprobado = respuestasJSON;
      }
    }
  }

  if (aprobado) {
    let errores = [];
    try {
      if (respuestasAprobado) {
        errores = calificarExamen(JSON.parse(respuestasAprobado), correo).errores;
      }
    } catch (e) {
      errores = []; // si algo del formato guardado no cuadra, mejor no mostrar nada a mostrar un dato erróneo
    }
    return { estado: 'APROBADO', config, mejorCalificacion, errores };
  }
  if (intentos >= config.maxIntentos) {
    return { estado: 'BLOQUEADO_INTENTOS', config, intentos };
  }
  return {
    estado: 'PUEDE_PRESENTAR',
    config,
    intentos,
    intentosRestantes: config.maxIntentos - intentos
  };
}

function registrarIntento(correo, preguntaIds, respuestas, calificacion, aprobado, cambiosFoco, aciertosCrudos, puntoExtraAplicado) {
  const config = obtenerConfig();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_INTENTOS);
    const numIntentoActual = contarIntentosSinLock(hoja, correo, config.campanaId) + 1;
    hoja.appendRow([
      new Date(),
      correo,
      config.campanaId,
      numIntentoActual,
      preguntaIds.join(','),
      JSON.stringify(respuestas),
      calificacion,
      aprobado,
      cambiosFoco || 0,
      aciertosCrudos != null ? aciertosCrudos : calificacion,
      puntoExtraAplicado || 0
    ]);
    return numIntentoActual;
  } finally {
    lock.releaseLock();
  }
}

function contarIntentosSinLock(hoja, correo, campanaId) {
  const datos = hoja.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).toLowerCase() === correo.toLowerCase() && datos[i][2] === campanaId) n++;
  }
  return n;
}
