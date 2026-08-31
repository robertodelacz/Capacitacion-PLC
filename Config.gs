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
    campanaId: config['CAMPANA_ID'] || 'DEMO',
    ventanaInicio: new Date(config['VENTANA_INICIO']),
    ventanaFin: new Date(config['VENTANA_FIN']),
    numPreguntasExamen: Number(config['NUM_PREGUNTAS_EXAMEN']) || 5,
    puntajeMinimo: Number(config['PUNTAJE_MINIMO']) || 4,
    maxIntentos: Number(config['MAX_INTENTOS']) || 3,
    duracionExamenMinutos: Number(config['DURACION_EXAMEN_MINUTOS']) || 40,
    correoOficialCumplimiento: config['CORREO_OFICIAL_CUMPLIMIENTO'] || '',
    templateSlidesId: config['TEMPLATE_SLIDES_ID'] || '',
    carpetaConstanciasId: config['CARPETA_CONSTANCIAS_ID'] || ''
  };
}
