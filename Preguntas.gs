/**
 * Preguntas.gs
 * Selección aleatoria y calificación del examen. La respuesta correcta
 * NUNCA se envía al cliente antes de calificar — todo ocurre en el servidor
 * para que no pueda leerse desde el código fuente del navegador.
 *
 * Además, el orden de las opciones (A/B/C/D) se mezcla de forma distinta
 * para cada colaborador en cada pregunta — así, si dos personas reciben por
 * coincidencia la misma pregunta, la respuesta correcta no está en la misma
 * letra para ambas, y un "¡es la C!" en voz alta deja de servir. La mezcla
 * es determinista (misma persona + misma pregunta = mismo orden siempre),
 * así que si alguien recarga la página no le cambian las opciones a media
 * respuesta.
 */

function elegirPreguntasParaSesion(numPreguntas) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_BANCO_PREGUNTAS);
  const datos = hoja.getDataRange().getValues();
  const ids = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0]) ids.push(datos[i][0]);
  }
  return mezclarArreglo(ids.slice()).slice(0, Math.min(numPreguntas, ids.length));
}

function obtenerPreguntaPorId(id, correo) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_BANCO_PREGUNTAS);
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    const [filaId, tema, pregunta, a, b, c, d] = datos[i];
    if (String(filaId) === String(id)) {
      const semilla = hashSemilla(correo + '|' + id);
      const mezcladas = mezclarConSemilla([a, b, c, d], semilla);
      return {
        id: filaId,
        tema,
        pregunta,
        opciones: { A: mezcladas[0], B: mezcladas[1], C: mezcladas[2], D: mezcladas[3] }
      };
    }
  }
  return null;
}

function mezclarArreglo(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

/** Hash simple y estable de un texto, usado como semilla del PRNG. */
function hashSemilla(texto) {
  let hash = 0;
  for (let i = 0; i < texto.length; i++) {
    hash = ((hash << 5) - hash + texto.charCodeAt(i)) | 0;
  }
  return Math.abs(hash) || 1;
}

/** Generador pseudoaleatorio determinista (mismo arreglo + misma semilla = mismo resultado siempre). */
function mezclarConSemilla(arr, semilla) {
  const copia = arr.slice();
  let s = semilla;
  for (let i = copia.length - 1; i > 0; i--) {
    s = (s * 9301 + 49297) % 233280;
    const j = Math.floor((s / 233280) * (i + 1));
    [copia[i], copia[j]] = [copia[j], copia[i]];
  }
  return copia;
}

function calificarExamen(respuestas, correo) {
  // respuestas: [{ id, opcionElegida }] — opcionElegida es la letra tal como
  // se le mostró a ESTA persona, así que la calificación tiene que rehacer
  // la misma mezcla (con su correo) para saber qué letra representaba la
  // respuesta correcta en su pantalla — nunca se confía en nada calculado
  // desde el cliente.
  const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_BANCO_PREGUNTAS);
  const datos = hoja.getDataRange().getValues();
  const filasPorId = {};
  for (let i = 1; i < datos.length; i++) {
    filasPorId[String(datos[i][0])] = datos[i];
  }

  let aciertos = 0;
  respuestas.forEach(r => {
    const fila = filasPorId[String(r.id)];
    if (!fila) return;
    const [, , , a, b, c, d, respuestaCorrecta] = fila;
    const letras = ['A', 'B', 'C', 'D'];
    const semilla = hashSemilla(correo + '|' + r.id);
    const ordenMezclado = mezclarConSemilla(letras, semilla);
    const posicionCorrecta = ordenMezclado.indexOf(String(respuestaCorrecta).trim().toUpperCase());
    const letraCorrectaMostrada = letras[posicionCorrecta];

    if (letraCorrectaMostrada && letraCorrectaMostrada === String(r.opcionElegida).toUpperCase()) {
      aciertos++;
    }
  });

  return { aciertos, total: respuestas.length };
}
