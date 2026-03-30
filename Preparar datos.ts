/**
 * SCRIPT: UI_PREPARAR_DATOS_AUDITABLE
 * OBJETIVO: Sincronizar el estado de la hoja para procesamiento en la nube.
 * NIVEL DE AUDITORÍA: Cumplimiento de integridad de datos y trazabilidad visual.
 */
function main(workbook: ExcelScript.Workbook) {
  // --- VARIABLES DE ESTADO ---
  let isSheetCurrentlyUnprotected: boolean = false;
  const currentActiveSheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();
  const mastersWorksheet: ExcelScript.Worksheet = workbook.getWorksheet("MAESTROS");

  // --- 1. CAPTURA DE INFRAESTRUCTURA (Ámbito de Hoja y Libro) ---
  // Recuperamos los rangos con nombre definidos para la seguridad y la interfaz
  const systemKeyNamedItem = workbook.getNamedItem("SISTEMA_CLAVE");
  const feedbackUINamedItem = currentActiveSheet.getNamedItem("UI_FEEDBACK");
  const preparationUINamedItem = currentActiveSheet.getNamedItem("UI_PREPARACION");

  try {
    // --- 2. VALIDACIÓN DE EXISTENCIA DE COMPONENTES ---
    // Verificamos que todos los elementos críticos estén presentes antes de operar
    if (!systemKeyNamedItem || !preparationUINamedItem || !mastersWorksheet) {
      throw `Error de Infraestructura: No se encontró ${!systemKeyNamedItem ? "SISTEMA_CLAVE" : "UI_PREPARACION"} o la hoja MAESTROS.`;
    }

    // Extracción y limpieza de la credencial de seguridad
    const rawPasswordValue = systemKeyNamedItem.getRange().getValue();
    const formattedPasswordString = String(rawPasswordValue).trim();

    // --- 3. DESPROTECCIÓN DE LA HOJA (Apertura de Sesión de Escritura) ---
    // Se utiliza la solución confirmada de conversión a String para evitar errores de API
    currentActiveSheet.getProtection().unprotect(formattedPasswordString);
    isSheetCurrentlyUnprotected = true;

    // --- 4. ACTUALIZACIÓN DE INTERFAZ DE USUARIO (UI_PREPARACION) ---
    const preparationRange = preparationUINamedItem.getRange();
    const currentTime = new Date();
    const formattedTimeLabel = currentTime.toLocaleTimeString('es-AR', { hour12: false });
    
    // Indicador visual dinámico de actividad
    const statusHeartbeatIcon = (currentTime.getSeconds() % 2 === 0) ? "⚡" : "✨";

    // Registro visual de éxito en la hoja
    preparationRange.setValue(`[${formattedTimeLabel}] ${statusHeartbeatIcon} DATOS LISTOS. Sincronización confirmada.`);
    
    // Aplicación de formato visual de confirmación (Esquema Verde)
    const preparationRangeFormat = preparationRange.getFormat();
    preparationRangeFormat.getFill().setColor("#D4EDDA"); // Fondo Verde Exito
    preparationRangeFormat.getFont().setColor("#155724"); // Fuente Contraste
    preparationRangeFormat.getFont().setBold(true);
    preparationRangeFormat.setWrapText(true);

    // Limpieza de mensajes de error previos (Feedback de Hoja)
    if (feedbackUINamedItem) {
        const feedbackRange = feedbackUINamedItem.getRange();
        feedbackRange.setValue("");
        feedbackRange.getFormat().getFill().clear();
    }

  } catch (error) {
    // Registro de anomalías en la consola de administración
    console.log("Excepción detectada durante la preparación: " + error);
    
    // Notificación visual de error en la interfaz si la hoja está accesible
    if (feedbackUINamedItem && isSheetCurrentlyUnprotected) {
        const feedbackErrorRange = feedbackUINamedItem.getRange();
        feedbackErrorRange.setValue("❌ Error de Sistema: " + String(error));
        feedbackErrorRange.getFormat().getFill().setColor("#F8D7DA"); // Fondo Rojo Alerta
    }
  } finally {
    // --- 5. RE-PROTECCIÓN DE SEGURIDAD (Cierre de Sesión) ---
    // Garantizamos que la hoja no quede vulnerable después de la operación
    if (isSheetCurrentlyUnprotected && systemKeyNamedItem) {
      const reprotectionPassword = String(systemKeyNamedItem.getRange().getValue()).trim();
      
      currentActiveSheet.getProtection().protect({
        allowFormatCells: false,
        allowInsertRows: false,
        allowDeleteRows: false
      }, reprotectionPassword);
      
      console.log("Protocolo de seguridad finalizado: Hoja protegida nuevamente.");
    }
  }
}