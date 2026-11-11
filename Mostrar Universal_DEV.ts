/**
 * SCRIPT: ADMIN_MOSTRAR_HOJAS_SGC
 * OBJETIVO: Revelar la infraestructura oculta (veryHidden) exclusivamente para auditorías de QA.
 * SEGURIDAD: Protegido por validación de credenciales (Zero Trust).
 */
function main(
  workbook: ExcelScript.Workbook,
  claveAdministradorIngresada: string = "" // Parámetro a vincular con una celda input de UI
) {
  const NOMBRE_ITEM_CLAVE: string = "SISTEMA_CLAVE";
  let registroMensajeLog: string = "Inicio de rutina de revelación QA.";
  let ejecucionHabilitada: boolean = true;
  let claveSeguridadSistema: string = "";

  // --- I. VALIDACIÓN DE INFRAESTRUCTURA Y SEGURIDAD (ZERO TRUST) ---
  const rangoItemClave = workbook.getNamedItem(NOMBRE_ITEM_CLAVE)?.getRange();
  
  if (!rangoItemClave) {
    registroMensajeLog = `⛔ Error de Infraestructura: No se encontró el ítem '${NOMBRE_ITEM_CLAVE}'.`;
    ejecucionHabilitada = false;
  } else {
    claveSeguridadSistema = String(rangoItemClave.getText()).trim();
    
    // Autenticación de Negocio
    if (claveAdministradorIngresada !== claveSeguridadSistema || claveSeguridadSistema === "") {
      registroMensajeLog = "⚠️ Acceso Denegado: Credenciales de administrador inválidas o no proporcionadas.";
      ejecucionHabilitada = false;
    }
  }

  // --- II. MOTOR DE MUTACIÓN DE ESTADO ---
  if (ejecucionHabilitada) {
    try {
      const matrizHojas: ExcelScript.Worksheet[] = workbook.getWorksheets();
      let contadorReveladas: number = 0;

      matrizHojas.forEach((hojaActual: ExcelScript.Worksheet) => {
        const visibilidadActual = hojaActual.getVisibility();
        
        if (visibilidadActual !== ExcelScript.SheetVisibility.visible) {
          hojaActual.setVisibility(ExcelScript.SheetVisibility.visible);
          contadorReveladas++;
        }
      });

      registroMensajeLog = `✅ Auditoría QA Habilitada: ${contadorReveladas} hojas confidenciales han sido reveladas.`;

    } catch (errorEjecucion) {
      registroMensajeLog = `❌ Falla crítica de infraestructura al revelar hojas: ${(errorEjecucion as Error).message}`;
    } finally {
      // Punto único de salida estandarizado
      console.log(registroMensajeLog);
    }
  } else {
    // Registro de rechazo de acceso
    console.log(registroMensajeLog);
  }
}