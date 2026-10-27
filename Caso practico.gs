

const SHEET_CASO_PRACTICO = 'CasoPractico_Respuestas';
const SHEET_CASO_PRACTICO_BORRADOR = 'CasoPractico_Borradores';

const ENCABEZADOS_CASO_PRACTICO = [
  'Fecha', 'Folio', 'Nombre completo', 'Puesto',
  'V1 Pagaré', 'V2 Contrato de crédito',
  'C1 Identificación', 'C2a Accionistas iniciales', 'C2c Rep. legal inicial', 'C2d Facultades',
  'C3a Accionistas actuales', 'C3b Rep. legal actual', 'C4 Congruencia',
  'Rúbrica — % de coincidencia por campo (orientativo, no es la calificación final)',
  'Rúbrica — promedio general (%)',
  'Correo', 'Liga hoja de ejercicios'
];


const ENCABEZADOS_BORRADOR_CP = [
  'Correo', 'Última actualización', 'Nombre completo', 'Puesto',
  'V1 Pagaré', 'V2 Contrato de crédito',
  'C1 Identificación', 'C2a Accionistas iniciales', 'C2c Rep. legal inicial', 'C2d Facultades',
  'C3a Accionistas actuales', 'C3b Rep. legal actual', 'C4 Congruencia',
  'Liga hoja de ejercicios'
];


function _rubricaCasoPractico() {
  return {
    v1: { tipo: 'texto', claves: ['deudor', 'monto', 'letra', 'fecha', 'tasa', 'aval', 'firma'] },
    v2: { tipo: 'texto', claves: ['partes', 'monto', 'plazo', 'tasa', 'comision', 'garantia', 'vencimiento', 'domicilio', 'jurisdiccion', 'firma', 'anexo'] },

    c1: { tipo: 'texto', claves: ['vanjupe', '21', 'noviembre', '2005'] },
    c2a: { tipo: 'texto', claves: ['juan pablo ramirez mendoza', 'jose de jesus ramirez mendoza', '50'] },
    c2c: { tipo: 'texto', claves: ['oralia mendoza fernandez', 'apoderada'] },
    c2d: { tipo: 'texto', claves: ['pleitos y cobranzas', 'actos de administracion', 'actos de dominio'] },
    c3a: { tipo: 'texto', claves: ['vaneira ramirez mendoza', '99', 'jose de jesus ramirez mendoza', '0.98'] },
    c3b: { tipo: 'texto', claves: ['vaneira', 'administradora unica', 'oralia', 'apoderada'] },
    c4: { tipo: 'texto', claves: ['posterior', 'donacion', 'herencia', 'aumento', 'capital'] }
  };
}

function _normalizarCP(texto) {
  return String(texto || '')
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .trim();
}

function _coincidenciaTextoCP(respuesta, claves) {
  const texto = _normalizarCP(respuesta);
  if (!texto) return 0;
  const encontradas = claves.filter(c => texto.includes(_normalizarCP(c)));
  return claves.length ? Math.round((encontradas.length / claves.length) * 100) : 0;
}

function _coincidenciaNumeroCP(respuesta, esperado) {
  const limpio = String(respuesta || '').replace(/[^0-9.\-]/g, '');
  if (limpio === '') return 0;
  const numero = Number(limpio);
  if (isNaN(numero)) return 0;
  if (numero === esperado) return 100;
  const diferencia = Math.abs(numero - esperado);
  const base = Math.max(Math.abs(esperado), 1);
  return Math.max(0, Math.round(100 - (diferencia / base) * 100));
}

function _calificarRubricaCP(respuestas) {
  const rubrica = _rubricaCasoPractico();
  const porcentajes = {};
  let suma = 0;
  let n = 0;

  Object.keys(rubrica).forEach(campo => {
    const def = rubrica[campo];
    const valor = respuestas[campo] || '';
    const pct = def.tipo === 'numero'
      ? _coincidenciaNumeroCP(valor, def.esperado)
      : _coincidenciaTextoCP(valor, def.claves);
    porcentajes[campo] = pct;
    suma += pct;
    n++;
  });

  const promedio = n ? Math.round(suma / n) : 0;
  return { porcentajes, promedio };
}

function _asegurarHojaCasoPractico() {
  const libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName(SHEET_CASO_PRACTICO);
  if (!hoja) {
    hoja = libro.insertSheet(SHEET_CASO_PRACTICO);
    hoja.appendRow(ENCABEZADOS_CASO_PRACTICO);
    hoja.getRange(1, 1, 1, ENCABEZADOS_CASO_PRACTICO.length).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function _asegurarHojaBorradorCP() {
  const libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName(SHEET_CASO_PRACTICO_BORRADOR);
  if (!hoja) {
    hoja = libro.insertSheet(SHEET_CASO_PRACTICO_BORRADOR);
    hoja.appendRow(ENCABEZADOS_BORRADOR_CP);
    hoja.getRange(1, 1, 1, ENCABEZADOS_BORRADOR_CP.length).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

/** Busca la fila de un correo (ya normalizado por el token) en la columna A. -1 si no existe todavía. */
function _filaBorradorCP(hoja, correo) {
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return -1;
  const correos = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues();
  for (let i = 0; i < correos.length; i++) {
    if (correos[i][0] === correo) return i + 2; // +2: encabezado (fila 1) + índice base 0
  }
  return -1;
}

/**
 * Llamada en silencio desde el cliente cada ~25s mientras hay cambios sin
 * respaldar. El correo del token (identidad ya verificada por el login)
 * siempre sobrescribe su propia fila — nunca se acumulan filas repetidas
 * por cada autoguardado. No califica nada: solo guarda crudo.
 */
function guardarBorradorCasoPractico(candidato, respuestas, token) {
  const usuario = _requerirSesionCP(token);
  respuestas = respuestas || {};
  const nombre = String((candidato && candidato.nombre) || '').trim();
  const puesto = String((candidato && candidato.puesto) || '').trim();
  const correo = usuario.correo; // nunca lo que mande el cliente — siempre el autenticado
  const hojaUrl = String((candidato && candidato.hojaUrl) || '').trim();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const hoja = _asegurarHojaBorradorCP();
    const campo = id => respuestas[id] || '';
    const fila = [
      correo, new Date(), nombre, puesto,
      campo('v1'), campo('v2'),
      campo('c1'), campo('c2a'), campo('c2c'), campo('c2d'),
      campo('c3a'), campo('c3b'), campo('c4'),
      hojaUrl
    ];

    const filaExistente = _filaBorradorCP(hoja, correo);
    if (filaExistente === -1) {
      hoja.appendRow(fila);
    } else {
      hoja.getRange(filaExistente, 1, 1, fila.length).setValues([fila]);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** Borra el borrador de un correo ya enviado con éxito — mantiene la hoja de respaldos solo con casos en curso. */
function _eliminarBorradorCP(correo) {
  correo = String(correo || '').trim();
  if (!correo) return;
  const hoja = _asegurarHojaBorradorCP();
  const fila = _filaBorradorCP(hoja, correo);
  if (fila !== -1) hoja.deleteRow(fila);
}

function _generarFolioCP(nombre) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const hash = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, (nombre || '') + timestamp)
  ).substring(0, 6).toUpperCase();
  return `CP-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd')}-${hash}`;
}

/**
 * Llamada desde el cliente al dar clic en "Enviar respuestas". Nunca lanza
 * el detalle de la rúbrica de vuelta al candidato — solo folio y éxito.
 * El correo viene siempre del token ya verificado, nunca de lo que el
 * cliente mande suelto en "candidato".
 */
function registrarCasoPractico(candidato, respuestas, token) {
  const usuario = _requerirSesionCP(token);
  const nombre = String((candidato && candidato.nombre) || '').trim();
  const puesto = String((candidato && candidato.puesto) || '').trim();
  const correo = usuario.correo;
  if (!nombre || !puesto) {
    throw new Error('Falta el nombre completo o el puesto del candidato.');
  }
  respuestas = respuestas || {};

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Recupera la liga de su hoja de ejercicios (si llegó a generarse)
    // ANTES de borrar el borrador — la necesitamos para dejarla registrada
    // en la fila final y para revocarle el acceso de edición: ya que
    // envió, no debe poder seguir modificándola.
    let ligaHoja = '';
    try {
      const hojaBorrador = _asegurarHojaBorradorCP();
      const filaBorrador = _filaBorradorCP(hojaBorrador, correo);
      if (filaBorrador !== -1) {
        const colLiga = ENCABEZADOS_BORRADOR_CP.indexOf('Liga hoja de ejercicios') + 1;
        ligaHoja = String(hojaBorrador.getRange(filaBorrador, colLiga).getValue() || '').trim();
      }
    } catch (e) { /* no crítico para el envío final */ }

    const hoja = _asegurarHojaCasoPractico();
    const folio = _generarFolioCP(nombre);
    const { porcentajes, promedio } = _calificarRubricaCP(respuestas);

    const campo = id => respuestas[id] || '';
    hoja.appendRow([
      new Date(), folio, nombre, puesto,
      campo('v1'), campo('v2'),
      campo('c1'), campo('c2a'), campo('c2c'), campo('c2d'),
      campo('c3a'), campo('c3b'), campo('c4'),
      JSON.stringify(porcentajes),
      promedio,
      correo, ligaHoja
    ]);

    if (ligaHoja) {
      try { _revocarAccesoHojaCP(ligaHoja, correo); } catch (e) { /* no crítico — Roberto puede revocarlo manualmente desde Drive */ }
    }

    // Ya quedó registrado de forma definitiva — el borrador de este correo
    // sobra. Si falla (p. ej. no había borrador), no debe tronar el envío,
    // que ya se completó.
    try { _eliminarBorradorCP(correo); } catch (e) { /* no crítico */ }

    return { ok: true, folio };
  } finally {
    lock.releaseLock();
  }
}

// ================= HOJA DE EJERCICIOS (Sheets) =================
// Dispersiones + Ejercicio A + Ejercicio B ya no se responden en el
// formulario: viven en una copia de Sheets por candidato, generada y
// compartida sola en cuanto captura un correo válido (crearHojaEjerciciosCP,
// llamada desde CasoPracticoJS.html al salir del campo de correo). Así
// Roberto ve no solo si el resultado es correcto, sino cómo llegó ahí
// (fórmulas, orden, método) — y conserva control porque la hoja se
// comparte solo a ese correo, nunca "cualquiera con el link".

/** Carpeta raíz que Roberto crea a mano en su Drive — nunca la creamos nosotros, solo la buscamos. */
function _carpetaCasoPracticoCP() {
  const iter = DriveApp.getFoldersByName('Caso Práctico');
  if (!iter.hasNext()) {
    throw new Error('No se encontró la carpeta "Caso Práctico" en tu Drive. Créala primero y vuelve a intentar.');
  }
  return iter.next();
}

/** Subcarpeta con el nombre del candidato dentro de la carpeta raíz — se crea la primera vez que se necesita. */
function _carpetaCandidatoCP(carpetaRaiz, nombreCandidato) {
  const nombre = String(nombreCandidato || '').trim() || 'Candidato sin nombre';
  const iter = carpetaRaiz.getFoldersByName(nombre);
  if (iter.hasNext()) return iter.next();
  return carpetaRaiz.createFolder(nombre);
}

/** Extrae el ID de archivo de Drive de una liga de Sheets (.../d/<ID>/edit...). */
function _idDesdeLigaSheetCP(liga) {
  const m = String(liga || '').match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
}

/** true si la liga apunta a un archivo que sigue existiendo y no está en la papelera. */
function _ligaVivaCP(liga) {
  const id = _idDesdeLigaSheetCP(liga);
  if (!id) return false;
  try {
    const archivo = DriveApp.getFileById(id);
    return !archivo.isTrashed();
  } catch (e) {
    return false; // borrado de forma permanente, o ya no accesible
  }
}

/** Le quita el acceso de edición al candidato sobre su hoja — Roberto conserva el control porque es el dueño de la carpeta. */
function _revocarAccesoHojaCP(liga, correo) {
  if (!liga || !correo) return;
  const id = _idDesdeLigaSheetCP(liga);
  if (!id) return;
  DriveApp.getFileById(id).removeEditor(correo);
}

/** Llena las 3 pestañas con los mismos datos y escenarios que antes vivían en el HTML — protegidas, salvo las celdas donde el candidato escribe su respuesta. */
function _llenarHojaEjerciciosCP(ss) {
  // ---- Pestaña 1: Dispersiones ----
  const hDisp = ss.getSheets()[0];
  hDisp.setName('Dispersiones');
  hDisp.getRange(1, 1).setValue('Revisión de una programación').setFontWeight('bold').setFontSize(13);
  hDisp.getRange(2, 1, 1, 9).merge().setValue(
    'Esta tabla consolida 9 dispersiones distintas programadas para el 16 de octubre de 2026, cruzando la ' +
    'información del correo del Área Jurídica con la programación bancaria de Cuentas por Pagar. Confirma que ' +
    'coincidan exactamente el monto, la cuenta CLABE y el estatus de restricciones del cliente, asegura que cada ' +
    'registro sea único, e identifica cualquier otro dato que consideres erróneo. Anota tus observaciones en la ' +
    'columna "Observaciones" — usa fórmulas donde te ayude a comparar.'
  ).setWrap(true).setFontStyle('italic').setFontColor('#4C525E');
  hDisp.setRowHeight(2, 60);

  const encabezadosDisp = ['N.°', 'Concepto', 'Fecha', 'Monto (correo)', 'Monto (programado)', 'CLABE (correo)', 'CLABE (programado)', 'Restricción activa', 'Observaciones'];
  hDisp.getRange(4, 1, 1, encabezadosDisp.length).setValues([encabezadosDisp])
    .setFontWeight('bold').setBackground('#15171C').setFontColor('#fff');

  const filasDisp = [
    [1, 'Dispersión A', '16/10/2026', 85000, 85000, '646180157012345678', '646180157012345678', 'No', ''],
    [2, 'Dispersión B', '16/10/2026', 42000, 42000, '646180157012845676', '646180157012345676', 'No', ''],
    [3, 'Dispersión C', '16/10/2026', 67500, 67500, '012180045098765432', '012180045098765432', 'Sí, retener monto', ''],
    [4, 'Dispersión D', '16/10/2026', 120000, 120000, '646180157055554444', '646180157055554444', 'No', ''],
    [5, 'Dispersión E', '14/10/2026', 15000, 150000, '646180157066667777', '646180157066667777', 'No', ''],
    [6, 'Dispersión F', '16/10/2026', 53200, 53200, '646180157078889999', '646180157087889999', 'No', ''],
    [7, 'Dispersión G', '16/10/2026', 38000, 38000, '646180157099990000', '646180157099990000', 'No', ''],
    [8, 'Dispersión G', '16/10/2026', 38000, 38000, '646180157099990000', '646180157099990000', 'No', ''],
    [9, 'Dispersión H', '16/09/2026', 65000, 56000, '646180157013334444', '646180157013334444', 'No', '']
  ];
  // Las columnas de CLABE se fijan como texto plano ANTES de escribir los
  // valores — si se deja en automático, Sheets las reinterpreta como
  // número en cuanto alguien retipea la celda (los 18 dígitos exceden lo
  // que un número puede mostrar sin notación científica, p. ej. 6.4618E+17).
  hDisp.getRange(5, 6, filasDisp.length, 2).setNumberFormat('@');
  hDisp.getRange(5, 1, filasDisp.length, encabezadosDisp.length).setValues(filasDisp);
  hDisp.getRange(5, 4, filasDisp.length, 2).setNumberFormat('$#,##0.00');
  hDisp.getRange(5, 9, filasDisp.length, 1).setBackground('#FFFBEF');
  hDisp.setFrozenRows(4);

  // Anchos fijos en vez de autoResizeColumns: con la celda de instrucciones
  // fusionada en A2:I2, el auto-resize le atribuye todo ese texto largo a
  // la columna A sola y la deja enorme — con anchos explícitos la tabla se
  // ve igual de bien sin ese efecto, y la columna de Observaciones queda
  // visible sin tener que hacer scroll horizontal.
  const anchosDisp = [40, 110, 90, 105, 120, 150, 150, 130, 220];
  anchosDisp.forEach((ancho, i) => hDisp.setColumnWidth(i + 1, ancho));

  const rangoLibreDisp = hDisp.getRange(5, 9, filasDisp.length, 1);
  const proteccionDisp = hDisp.protect().setDescription('Datos del caso — no editar');
  proteccionDisp.setUnprotectedRanges([rangoLibreDisp]);
  proteccionDisp.removeEditors(proteccionDisp.getEditors());

  // ---- Pestaña 2: Ejercicio A — Comisión por apertura ----
  const hA = ss.insertSheet('Ejercicio A', 1);
  hA.getRange(1, 1).setValue('Ejercicio A — Comisión por apertura').setFontWeight('bold').setFontSize(13);
  hA.getRange(2, 1, 1, 2).merge().setValue(
    'Se va a formalizar una operación con una línea de crédito autorizada de $1,800,000.00. Comisión por ' +
    'apertura: 2% + IVA. Gastos de formalización: $92,500.00. Calcula cuánto se retiene por comisión (con IVA) y ' +
    'cuánto se le dispersa finalmente al cliente — usa fórmulas, no captures el resultado a mano.'
  ).setWrap(true).setFontStyle('italic').setFontColor('#4C525E');
  hA.setRowHeight(2, 90);

  const datosA = [
    ['Línea autorizada', 1800000],
    ['% comisión por apertura', 0.02],
    ['Gastos de formalización', 92500],
    ['', ''],
    ['Comisión por apertura', ''],
    ['IVA', ''],
    ['Comisión total con IVA', ''],
    ['Monto a dispersar', '']
  ];
  hA.getRange(4, 1, datosA.length, 2).setValues(datosA);
  hA.getRange(4, 1, 3, 1).setFontWeight('bold');
  hA.getRange(8, 1, 4, 1).setFontWeight('bold');
  hA.getRange(4, 2, 1, 1).setNumberFormat('$#,##0.00');
  hA.getRange(5, 2, 1, 1).setNumberFormat('0.00%');
  hA.getRange(6, 2, 1, 1).setNumberFormat('$#,##0.00');
  hA.getRange(8, 2, 4, 1).setNumberFormat('$#,##0.00').setBackground('#FFFBEF');
  hA.setColumnWidth(1, 210);
  hA.setColumnWidth(2, 160);

  const rangoLibreA = hA.getRange(8, 2, 4, 1);
  const proteccionA = hA.protect().setDescription('Datos del caso — no editar');
  proteccionA.setUnprotectedRanges([rangoLibreA]);
  proteccionA.removeEditors(proteccionA.getEditors());

  // ---- Pestaña 3: Ejercicio B — Comisión por disposición ----
  const hB = ss.insertSheet('Ejercicio B', 2);
  hB.getRange(1, 1).setValue('Ejercicio B — Comisión por disposición').setFontWeight('bold').setFontSize(13);
  hB.getRange(2, 1, 1, 2).merge().setValue(
    'Un cliente con una línea de crédito autorizada de $1,000,000.00 solicita una disposición de recursos de ' +
    '$250,000.00. Comisión por disposición: 4% + IVA. Calcula el monto neto a dispersar — usa fórmulas, no ' +
    'captures el resultado a mano.'
  ).setWrap(true).setFontStyle('italic').setFontColor('#4C525E');
  hB.setRowHeight(2, 80);

  const datosB = [
    ['Línea autorizada', 1000000],
    ['Disposición solicitada', 250000],
    ['% comisión por disposición', 0.04],
    ['', ''],
    ['Comisión por disposición', ''],
    ['IVA', ''],
    ['Comisión total con IVA', ''],
    ['Neto a dispersar', '']
  ];
  hB.getRange(4, 1, datosB.length, 2).setValues(datosB);
  hB.getRange(4, 1, 3, 1).setFontWeight('bold');
  hB.getRange(8, 1, 4, 1).setFontWeight('bold');
  hB.getRange(4, 2, 2, 1).setNumberFormat('$#,##0.00');
  hB.getRange(6, 2, 1, 1).setNumberFormat('0.00%');
  hB.getRange(8, 2, 4, 1).setNumberFormat('$#,##0.00').setBackground('#FFFBEF');
  hB.setColumnWidth(1, 210);
  hB.setColumnWidth(2, 160);

  const rangoLibreB = hB.getRange(8, 2, 4, 1);
  const proteccionB = hB.protect().setDescription('Datos del caso — no editar');
  proteccionB.setUnprotectedRanges([rangoLibreB]);
  proteccionB.removeEditors(proteccionB.getEditors());

  ss.setActiveSheet(hDisp);
}

/**
 * Llamada desde el cliente en cuanto hay una sesión válida y un nombre
 * capturado. El correo SIEMPRE sale del token ya verificado por el login
 * — ya no es un parámetro que el candidato pueda escribir o cambiar, así
 * que no puede "corregirse" a medio caso; si alguien necesita cambiar de
 * correo, inicia sesión de nuevo con el correcto.
 * Idempotente por correo: si ya existe una hoja viva para este correo,
 * regresa la misma liga sin crear nada de más. Si la liga guardada ya no
 * sirve (se borró/movió a la papelera a mano, como al probar, o tú
 * borraste la fila en "CasoPractico_Borradores" para reiniciar a alguien),
 * se genera una hoja nueva en vez de regresar para siempre un link muerto.
 */
function crearHojaEjerciciosCP(nombre, token) {
  const usuario = _requerirSesionCP(token);
  nombre = String(nombre || '').trim();
  const correo = usuario.correo;
  if (!nombre) throw new Error('Falta el nombre completo del candidato.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const hoja = _asegurarHojaBorradorCP();
    const fila = _filaBorradorCP(hoja, correo);
    const colLiga = ENCABEZADOS_BORRADOR_CP.indexOf('Liga hoja de ejercicios') + 1;

    if (fila !== -1) {
      const ligaGuardada = String(hoja.getRange(fila, colLiga).getValue() || '').trim();
      if (ligaGuardada && _ligaVivaCP(ligaGuardada)) {
        return { url: ligaGuardada }; // idempotente: ya existe para este correo
      }
    }

    // No existe todavía (o la liga guardada ya no sirve): se crea desde
    // cero y se mueve a Caso Práctico/<candidato>/.
    const carpetaRaiz = _carpetaCasoPracticoCP();
    const carpetaCandidato = _carpetaCandidatoCP(carpetaRaiz, nombre);

    const ss = SpreadsheetApp.create('Caso Práctico — ' + nombre + ' — Ejercicios');
    _llenarHojaEjerciciosCP(ss);

    const archivo = DriveApp.getFileById(ss.getId());
    const padres = archivo.getParents();
    while (padres.hasNext()) { padres.next().removeFile(archivo); }
    carpetaCandidato.addFile(archivo);
    archivo.addEditor(correo);

    const liga = ss.getUrl();
    if (fila !== -1) {
      hoja.getRange(fila, colLiga).setValue(liga);
    } else {
      // Aún no había fila de borrador para este correo (el autoguardado de
      // ~25s no ha corrido todavía) — se crea una mínima con lo que ya se
      // sabe; el siguiente autoguardado la completa con el resto.
      const filaNueva = ENCABEZADOS_BORRADOR_CP.map(() => '');
      filaNueva[0] = correo;
      filaNueva[1] = new Date();
      filaNueva[2] = nombre;
      filaNueva[colLiga - 1] = liga;
      hoja.appendRow(filaNueva);
    }

    return { url: liga };
  } finally {
    lock.releaseLock();
  }
}
