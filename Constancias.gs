/**
 * Constancias.gs
 * Genera el PDF de la constancia a partir de tu plantilla de Slides,
 * sustituyendo los marcadores, y la archiva en Drive como respaldo
 * de auditoría antes de entregarla al colaborador.
 */

function generarConstancia(correo, nombre, puesto, calificacion, total) {
  const config = obtenerConfig();
  if (!config.templateSlidesId) {
    throw new Error('Falta configurar TEMPLATE_SLIDES_ID en la hoja Config.');
  }

  const folio = generarFolio(correo);
  const plantilla = DriveApp.getFileById(config.templateSlidesId);
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

  const pdfBlob = DriveApp.getFileById(copia.getId()).getAs('application/pdf');
  pdfBlob.setName(`Constancia_${folio}.pdf`);

  let pdfFile;
  if (config.carpetaConstanciasId) {
    pdfFile = DriveApp.getFolderById(config.carpetaConstanciasId).createFile(pdfBlob);
  } else {
    pdfFile = DriveApp.createFile(pdfBlob);
  }

  // La copia editable de Slides ya no se necesita una vez exportado el PDF.
  DriveApp.getFileById(copia.getId()).setTrashed(true);

  registrarConstancia(correo, folio, calificacion, total, pdfFile.getId(), pdfFile.getUrl());

  return {
    folio,
    pdfBase64: Utilities.base64Encode(pdfBlob.getBytes())
  };
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
