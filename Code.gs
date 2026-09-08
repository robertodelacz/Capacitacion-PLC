/**
 * Code.gs
 * Punto de entrada del WebApp y funciones expuestas al cliente vía
 * google.script.run. El cliente nunca decide el estado del usuario, el
 * tiempo restante, ni la calificación — solo pide y recibe, una pregunta
 * a la vez.
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Interfaz')
    .evaluate()
    .setTitle('Capacitación PLD-FT · Financiera Cualli')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}

/** Se llama al cargar la app: decide qué pantalla mostrar, incluido si hay un examen en curso. */
function iniciar() {
  const usuario = obtenerUsuarioActual();
  if (!usuario.autorizado) {
    return { pantalla: 'NO_AUTORIZADO', motivo: usuario.motivo };
  }

  const estado = obtenerEstadoUsuario(usuario.correo);
  const config = estado.config;

  if (estado.estado === 'FUERA_DE_VENTANA') {
    return {
      pantalla: 'FUERA_DE_VENTANA',
      ventanaInicioTexto: formatearFechaLegible(config.ventanaInicio),
      ventanaFinTexto: formatearFechaLegible(config.ventanaFin)
    };
  }
  if (estado.estado === 'APROBADO') {
    return { pantalla: 'APROBADO', nombre: usuario.nombre, calificacion: estado.mejorCalificacion, total: config.numPreguntasExamen };
  }
  if (estado.estado === 'BLOQUEADO_INTENTOS') {
    return { pantalla: 'BLOQUEADO', correoOficial: config.correoOficialCumplimiento };
  }

  // PUEDE_PRESENTAR: si ya había un examen en curso sin terminar, lo retoma.
  const sesion = obtenerSesionActiva(usuario.correo, config.campanaId);
  if (sesion) {
    const restantes = segundosRestantes(sesion, config.duracionExamenMinutos);
    if (restantes <= 0) {
      finalizarSesion(usuario, sesion.preguntaIds, sesion.respuestas, config);
      return iniciar(); // ya se consumió el intento por tiempo; reevalúa el estado
    }
    const preguntaActual = obtenerPreguntaPorId(sesion.preguntaIds[sesion.indiceActual], usuario.correo);
    return {
      pantalla: 'EXAMEN_EN_CURSO',
      nombre: usuario.nombre,
      puesto: usuario.puesto,
      pregunta: preguntaActual,
      progreso: { actual: sesion.indiceActual + 1, total: sesion.preguntaIds.length },
      segundosRestantes: restantes
    };
  }

  return {
    pantalla: 'SELECCION_EXAMEN',
    nombre: usuario.nombre,
    puesto: usuario.puesto,
    intentosRestantes: estado.intentosRestantes,
    maxIntentos: config.maxIntentos,
    numPreguntas: config.numPreguntasExamen,
    duracionMinutos: config.duracionExamenMinutos,
    puntajeMinimo: config.puntajeMinimo
  };
}

/** El colaborador ya confirmó "sí, soy yo" — aquí arranca el reloj y se arma su examen. */
function confirmarIdentidadEIniciar() {
  const usuario = obtenerUsuarioActual();
  if (!usuario.autorizado) throw new Error('No autorizado.');

  const estado = obtenerEstadoUsuario(usuario.correo);
  if (estado.estado !== 'PUEDE_PRESENTAR') throw new Error('No es posible iniciar un examen en este momento.');
  if (obtenerSesionActiva(usuario.correo, estado.config.campanaId)) {
    throw new Error('Ya tienes un examen en curso.');
  }

  const ids = elegirPreguntasParaSesion(estado.config.numPreguntasExamen);
  crearSesion(usuario.correo, estado.config.campanaId, ids);

  return {
    pregunta: obtenerPreguntaPorId(ids[0], usuario.correo),
    progreso: { actual: 1, total: ids.length },
    segundosRestantes: estado.config.duracionExamenMinutos * 60
  };
}

/** Registra la respuesta a la pregunta actual y entrega la siguiente, o el resultado si ya era la última. */
function responderPreguntaActual(preguntaId, opcionElegida, cambiosFoco) {
  const usuario = obtenerUsuarioActual();
  if (!usuario.autorizado) throw new Error('No autorizado.');

  const config = obtenerConfig();
  const sesion = obtenerSesionActiva(usuario.correo, config.campanaId);
  if (!sesion) throw new Error('No tienes un examen en curso.');

  const restantes = segundosRestantes(sesion, config.duracionExamenMinutos);
  if (restantes <= 0) {
    return finalizarSesion(usuario, sesion.preguntaIds, sesion.respuestas, config, cambiosFoco);
  }

  const idEsperado = sesion.preguntaIds[sesion.indiceActual];
  if (String(idEsperado) !== String(preguntaId)) {
    throw new Error('Esa no es la pregunta actual de tu examen.');
  }

  sesion.respuestas[preguntaId] = opcionElegida;
  const siguienteIndice = sesion.indiceActual + 1;

  if (siguienteIndice >= sesion.preguntaIds.length) {
    return finalizarSesion(usuario, sesion.preguntaIds, sesion.respuestas, config, cambiosFoco);
  }

  actualizarSesion(usuario.correo, config.campanaId, sesion.respuestas, siguienteIndice, cambiosFoco);
  return {
    terminado: false,
    pregunta: obtenerPreguntaPorId(sesion.preguntaIds[siguienteIndice], usuario.correo),
    progreso: { actual: siguienteIndice + 1, total: sesion.preguntaIds.length },
    segundosRestantes: segundosRestantes(sesion, config.duracionExamenMinutos)
  };
}

/** El cliente llama esto cada ~20s mientras hay examen en curso, para resincronizar el reloj y detectar si el tiempo ya se agotó del lado del servidor. */
function verificarSesion() {
  const usuario = obtenerUsuarioActual();
  if (!usuario.autorizado) throw new Error('No autorizado.');

  const config = obtenerConfig();
  const sesion = obtenerSesionActiva(usuario.correo, config.campanaId);
  if (!sesion) return { sesionActiva: false };

  const restantes = segundosRestantes(sesion, config.duracionExamenMinutos);
  if (restantes <= 0) {
    const resultado = finalizarSesion(usuario, sesion.preguntaIds, sesion.respuestas, config, sesion.cambiosFoco);
    return Object.assign({ sesionActiva: false, seAgotoElTiempo: true }, resultado);
  }
  return { sesionActiva: true, segundosRestantes: restantes };
}

/** Califica, registra el intento, cierra la sesión y genera la constancia si aplica. Se usa tanto al terminar normal como al agotarse el tiempo. */
function finalizarSesion(usuario, preguntaIds, respuestasObj, config, cambiosFoco) {
  const respuestas = preguntaIds.map(id => ({ id, opcionElegida: respuestasObj[id] || null }));
  const resultado = calificarExamen(respuestas, usuario.correo);
  const aprobado = resultado.aciertos >= config.puntajeMinimo;

  const numIntento = registrarIntento(usuario.correo, preguntaIds, respuestas, resultado.aciertos, aprobado, cambiosFoco || 0);
  eliminarSesion(usuario.correo, config.campanaId);

  const salida = {
    terminado: true,
    aciertos: resultado.aciertos,
    total: resultado.total,
    aprobado,
    numIntento,
    intentosRestantes: Math.max(0, config.maxIntentos - numIntento)
  };

  if (aprobado) {
    try {
      const constancia = generarConstancia(usuario.correo, usuario.nombre, usuario.puesto, resultado.aciertos, resultado.total);
      salida.pdfBase64 = constancia.pdfBase64;
      salida.folio = constancia.folio;
    } catch (e) {
      salida.constanciaPendiente = true;
      salida.errorConstancia = e.message;
    }
  } else if (numIntento >= config.maxIntentos) {
    salida.bloqueado = true;
    salida.correoOficial = config.correoOficialCumplimiento;
  }

  return salida;
}

function descargarConstanciaPropia() {
  const usuario = obtenerUsuarioActual();
  if (!usuario.autorizado) throw new Error('No autorizado.');
  return obtenerConstanciaMasReciente(usuario.correo);
}

function formatearFechaLegible(fecha) {
  const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const tz = Session.getScriptTimeZone();
  const dia = Utilities.formatDate(fecha, tz, 'd');
  const mesIndex = Number(Utilities.formatDate(fecha, tz, 'M')) - 1;
  const anio = Utilities.formatDate(fecha, tz, 'yyyy');
  return `${dia} de ${meses[mesIndex]} de ${anio}`;
}
