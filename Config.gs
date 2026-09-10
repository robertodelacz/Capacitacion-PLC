/**
 * Config.gs
 * Lee la hoja "Config" (columnas: Clave | Valor) y expone las constantes
 * del sistema. Ajusta los valores directamente en la hoja de cálculo,
 * nunca aquí en el código.
 */

const SHEET_CONFIG = 'Config';
const SHEET_COLABORADORES = 'ColaboradoresAutorizados';
const SHEET_BANCO_PREGUNTAS = 'BancoPreguntas';
const SHEET_INTENTOS = 'Intentos';
const SHEET_CONSTANCIAS = 'ConstanciasEmitidas';
const SHEET_SESIONES = 'SesionesActivas';

function obtenerConfig() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
  const datos = hoja.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < datos.length; i++) {
    const clave = String(datos[i][0]).trim();
    if (clave) config[clave] = datos[i][1];
  }
  return {
    campanaId: config['CAMPANA_ID'] || 'CAMPANA-2026-TEST',
    ventanaInicio: new Date(config['VENTANA_INICIO']),
    ventanaFin: new Date(config['VENTANA_FIN']),
    numPreguntasExamen: Number(config['NUM_PREGUNTAS_EXAMEN']) || 12,
    puntajeMinimo: Number(config['PUNTAJE_MINIMO']) || 9,
    maxIntentos: Number(config['MAX_INTENTOS']) || 2,
    duracionExamenMinutos: Number(config['DURACION_EXAMEN_MINUTOS']) || 30,
    correoOficialCumplimiento: config['CORREO_OFICIAL_CUMPLIMIENTO'] || '',
    templateSlidesIdCualli: config['TEMPLATE_SLIDES_ID_CUALLI'] || '',
    templateSlidesIdCualliAsociados: config['TEMPLATE_SLIDES_ID_CUALLI_ASOCIADOS'] || '',
    templateSlidesIdFractio: config['TEMPLATE_SLIDES_ID_FRACTIO'] || '',
    templateSlidesIdHipoo: config['TEMPLATE_SLIDES_ID_HIPOO'] || '',
    carpetaConstanciasId: config['CARPETA_CONSTANCIAS_ID'] || ''
  };
}
