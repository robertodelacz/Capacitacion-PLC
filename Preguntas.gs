/**
 * Preguntas.gs
 * Selección aleatoria y calificación del examen. La respuesta correcta
 * NUNCA se envía al cliente antes de calificar — todo ocurre en el servidor
 * para que no pueda leerse desde el código fuente del navegador.
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

function obtenerPreguntaPorId(id) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_BANCO_PREGUNTAS);
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    const [filaId, tema, pregunta, a, b, c, d] = datos[i];
    if (String(filaId) === String(id)) {
      return { id: filaId, tema, pregunta, opciones: { A: a, B: b, C: c, D: d } };
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

function calificarExamen(respuestas) {
  // respuestas: [{ id, opcionElegida }] — nunca se confía en una calificación
  // que venga calculada desde el cliente.
  const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_BANCO_PREGUNTAS);
  const datos = hoja.getDataRange().getValues();
  const correctas = {};

  for (let i = 1; i < datos.length; i++) {
    const [id, , , , , , , respuestaCorrecta] = datos[i];
    correctas[id] = String(respuestaCorrecta).trim().toUpperCase();
  }

  let aciertos = 0;
  respuestas.forEach(r => {
    if (correctas[r.id] && correctas[r.id] === String(r.opcionElegida).toUpperCase()) {
      aciertos++;
    }
  });

  return { aciertos, total: respuestas.length };
}
