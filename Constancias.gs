/**
 * Constancias.gs
 * Genera el PDF de la constancia a partir de tu plantilla de Slides,
 * sustituyendo los marcadores, y la archiva en Drive como respaldo
 * de auditoría antes de entregarla al colaborador.
 */

/** Mapa de dominio de correo -> plantilla de Slides de esa empresa.
    Las tres empresas del grupo tienen su propio diseño (logo, nombre,
    certificaciones propias), aunque comparten el mismo oficial de
    cumplimiento como firmante. */
function obtenerMapaTemplates(config) {
  return {
    'cualli.mx': { id: config.templateSlidesIdCualli, clave: 'TEMPLATE_SLIDES_ID_CUALLI' },
    'fractio.mx': { id: config.templateSlidesIdFractio, clave: 'TEMPLATE_SLIDES_ID_FRACTIO' },
    'hipoo.mx': { id: config.templateSlidesIdHipoo, clave: 'TEMPLATE_SLIDES_ID_HIPOO' }
  };
}

/** Decide qué plantilla de Slides usar. Primero resuelve el caso especial
    de los asociados de Cualli sin puesto (mismo dominio @cualli.mx, pero
    con una plantilla propia que no lleva el marcador {{PUESTO}}); ese caso
    se detecta por el campo Puesto vacío en el roster, no por el dominio.
    Para el resto, aplica el mapa dominio -> plantilla de siempre. */
function seleccionarTemplateId(correo, puesto, config) {
  const dominio = (String(correo).split('@')[1] || '').trim().toLowerCase();

  if (dominio === 'cualli.mx' && !String(puesto || '').trim()) {
    if (!config.templateSlidesIdCualliAsociados) {
      throw new Error('Falta configurar TEMPLATE_SLIDES_ID_CUALLI_ASOCIADOS en la hoja Config.');
    }
    return config.templateSlidesIdCualliAsociados;
  }

  const mapa = obtenerMapaTemplates(config);
  const entrada = mapa[dominio];
  if (!entrada) {
    throw new Error(`No hay plantilla de constancia configurada para el dominio "${dominio}".`);
  }
  if (!entrada.id) {
    throw new Error(`Falta configurar ${entrada.clave} en la hoja Config.`);
  }
  return entrada.id;
}

function generarConstancia(correo, nombre, puesto, calificacion, total) {
  const config = obtenerConfig();
  const esPrueba = String(config.campanaId).toUpperCase().includes('PRUEBA');

  const folio = generarFolio(correo);
  let pdfBlob;

  if (esPrueba) {
    // Campaña de prueba: nunca se toca la plantilla oficial de Slides.
    // Se genera un PDF genérico, obviamente marcado como prueba, para
    // validar el flujo completo (generación + descarga) sin riesgo de
    // que alguien confunda esto con una constancia real de PLD/FT.
    pdfBlob = generarPdfDePrueba(nombre, folio, calificacion, total);
  } else {
    const templateId = seleccionarTemplateId(correo, puesto, config);
    const plantilla = DriveApp.getFileById(templateId);
    // makeCopy() sin carpeta destino falla cuando el archivo de origen vive en
    // una Unidad compartida — por eso siempre se indica una carpeta explícita.
    const carpetaDestino = config.carpetaConstanciasId
      ? DriveApp.getFolderById(config.carpetaConstanciasId)
      : DriveApp.getRootFolder();
    const copia = plantilla.makeCopy(`Constancia PLD-FT - ${nombre} - ${folio}`, carpetaDestino);
    const presentacion = SlidesApp.openById(copia.getId());

    const fecha = formatearFechaCertificado(new Date());
    presentacion.getSlides().forEach(slide => {
      slide.replaceAllText('{{NOMBRE}}', nombre);
      slide.replaceAllText('{{PUESTO}}', puesto || '');
      // {{FOLIO}} solo se sustituye si decides agregarlo al diseño; si no está
      // en la plantilla, replaceAllText simplemente no encuentra nada que cambiar.
      slide.replaceAllText('{{FOLIO}}', folio);
      slide.replaceAllText('{{FECHA}}', fecha);
      slide.replaceAllText('{{CALIFICACION}}', `${calificacion}/${total}`);
    });
    presentacion.saveAndClose();

    pdfBlob = DriveApp.getFileById(copia.getId()).getAs('application/pdf');
    pdfBlob.setName(`Constancia_${folio}.pdf`);

    // La copia editable de Slides ya no se necesita una vez exportado el PDF.
    DriveApp.getFileById(copia.getId()).setTrashed(true);
  }

  let pdfFile;
  if (config.carpetaConstanciasId) {
    pdfFile = DriveApp.getFolderById(config.carpetaConstanciasId).createFile(pdfBlob);
  } else {
    pdfFile = DriveApp.createFile(pdfBlob);
  }

  registrarConstancia(correo, folio, calificacion, total, pdfFile.getId(), pdfFile.getUrl());

  return {
    folio,
    pdfBase64: Utilities.base64Encode(pdfBlob.getBytes())
  };
}

/** PDF genérico y obviamente falso, solo para probar el flujo de descarga
    durante pruebas internas. Usa Slides (no Docs) a propósito: es el mismo
    servicio que ya usa la constancia real, así que no requiere autorizar
    ningún permiso nuevo. Crea una presentación temporal, la exporta a PDF
    y la borra — nunca toca la plantilla oficial. */
function generarPdfDePrueba(nombre, folio, calificacion, total) {
  const presentacion = SlidesApp.create('temporal-prueba-' + folio);
  const slide = presentacion.getSlides()[0];
  const ancho = presentacion.getPageWidth();

  const titulo = slide.insertTextBox('CONSTANCIA DE PRUEBA', 40, 50, ancho - 80, 50);
  titulo.getText().getTextStyle().setFontSize(26).setBold(true);
  titulo.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const aviso = slide.insertTextBox(
    'Este documento se generó únicamente para validar el funcionamiento técnico del sistema. ' +
    'No acredita ninguna capacitación oficial de PLD/FT ni tiene validez alguna.',
    40, 115, ancho - 80, 70
  );
  aviso.getText().getTextStyle().setFontSize(11);
  aviso.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const nombreBox = slide.insertTextBox(nombre, 40, 210, ancho - 80, 40);
  nombreBox.getText().getTextStyle().setFontSize(18).setBold(true);
  nombreBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const detalle = slide.insertTextBox(
    `Calificación de prueba: ${calificacion}/${total}`,
    40, 255, ancho - 80, 30
  );
  detalle.getText().getTextStyle().setFontSize(12);
  detalle.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const fechaBox = slide.insertTextBox(formatearFechaCertificado(new Date()), 40, 290, ancho - 80, 30);
  fechaBox.getText().getTextStyle().setFontSize(11);
  fechaBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  presentacion.saveAndClose();

  const pdfBlob = DriveApp.getFileById(presentacion.getId()).getAs('application/pdf');
  pdfBlob.setName(`Constancia_PRUEBA_${folio}.pdf`);

  // La presentación temporal ya no se necesita una vez exportado el PDF.
  DriveApp.getFileById(presentacion.getId()).setTrashed(true);

  return pdfBlob;
}

function formatearFechaCertificado(fecha) {
  const meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
  const tz = Session.getScriptTimeZone();
  const dia = Utilities.formatDate(fecha, tz, 'dd');
  const mesIndex = Number(Utilities.formatDate(fecha, tz, 'M')) - 1;
  const anio = Utilities.formatDate(fecha, tz, 'yyyy');
  return `${dia} ${meses[mesIndex]} ${anio}`;
}

function generarFolio(correo) {
  const config = obtenerConfig();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const hash = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, correo + timestamp)
  ).substring(0, 6).toUpperCase();
  return `${config.campanaId}-${hash}`;
}

function registrarConstancia(correo, folio, calificacion, total, pdfFileId, pdfUrl) {
  const config = obtenerConfig();
  const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_CONSTANCIAS);
  hoja.appendRow([correo, config.campanaId, folio, new Date(), `${calificacion}/${total}`, pdfFileId, pdfUrl]);
}

function obtenerConstanciaMasReciente(correo) {
  const config = obtenerConfig();
  const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_CONSTANCIAS);
  const datos = hoja.getDataRange().getValues();
  let ultima = null;

  for (let i = 1; i < datos.length; i++) {
    const [correoFila, campanaId, folio, , , pdfFileId] = datos[i];
    if (String(correoFila).toLowerCase() === correo.toLowerCase() && campanaId === config.campanaId) {
      ultima = { folio, pdfFileId };
    }
  }

  if (!ultima) throw new Error('No se encontró una constancia emitida para este correo en esta campaña.');
  const pdfBlob = DriveApp.getFileById(ultima.pdfFileId).getBlob();
  return { folio: ultima.folio, pdfBase64: Utilities.base64Encode(pdfBlob.getBytes()) };
}
