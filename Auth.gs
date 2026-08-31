/**
 * Auth.gs
 * Valida que el usuario autenticado en Google (por sesión, no por dato
 * capturado a mano) esté en el roster de colaboradores autorizados.
 * Este roster es la única fuente de verdad para el control de acceso,
 * sin importar de cuál de los dominios del grupo empresarial venga el correo.
 */

function obtenerUsuarioActual() {
  const correo = Session.getActiveUser().getEmail();
  if (!correo) {
    return { autorizado: false, motivo: 'NO_SESION' };
  }

  const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_COLABORADORES);
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    const [correoFila, nombre, puesto, activo] = datos[i];
    if (String(correoFila).trim().toLowerCase() === correo.toLowerCase()) {
      const estaActivo = activo === true || String(activo).toUpperCase() === 'TRUE';
      if (!estaActivo) {
        return { autorizado: false, motivo: 'INACTIVO' };
      }
      return { autorizado: true, correo, nombre, puesto };
    }
  }

  return { autorizado: false, motivo: 'NO_EN_ROSTER' };
}
