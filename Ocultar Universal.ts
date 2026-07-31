/**
 * SCRIPT: ADMIN_OCULTAR_HOJAS_SGC
 * OBJETIVO: Aislar la base de datos y los historiales de auditoría de la interfaz de usuario.
 * REGLA DE NEGOCIO: Toda hoja sin prefijo INP_ o BD_ es clasificada como Confidencial/Backend.
 */
function main(workbook: ExcelScript.Workbook) {
  let registroMensajeLog: string = "Inicio de rutina de ocultamiento de seguridad.";
  
  try {
    const matrizHojas: ExcelScript.Worksheet[] = workbook.getWorksheets();
    let contadorOcultadas: number = 0;
    let contadorMantenidas: number = 0;

    // Iteración en RAM sobre los objetos del DOM
    matrizHojas.forEach((hojaActual: ExcelScript.Worksheet) => {
      const nombreHoja: string = hojaActual.getName().trim().toUpperCase();
      
      // Aplicación estricta de la regla de negocio
      if (!nombreHoja.startsWith("INP_") && !nombreHoja.startsWith("BD_")) {
        // Nivel veryHidden: Bloquea el botón derecho "Mostrar" en la interfaz gráfica
        hojaActual.setVisibility(ExcelScript.SheetVisibility.veryHidden);
        contadorOcultadas++;
      } else {
        hojaActual.setVisibility(ExcelScript.SheetVisibility.visible);
        contadorMantenidas++;
      }
    });

    registroMensajeLog = `✅ Bóveda asegurada. ${contadorOcultadas} hojas bloqueadas (veryHidden). ${contadorMantenidas} interfaces expuestas.`;

  } catch (errorEjecucion) {
    registroMensajeLog = `❌ Falla crítica de infraestructura al mutar visibilidad: ${(errorEjecucion as Error).message}`;
  } finally {
    // Registro de auditoría garantizado
    console.log(registroMensajeLog);
  }
}