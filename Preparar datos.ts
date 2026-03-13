/**
 * Script: UI_Preparar_Datos_V5_Oficial
 * Objetivo: Sincronización confirmada con manejo de Scope de Hoja.
 */
function main(workbook: ExcelScript.Workbook) {
  let is_unprotected: boolean = false;
  const active_sheet = workbook.getActiveWorksheet();
  const masters_sheet = workbook.getWorksheet("MAESTROS");

  // 1. CAPTURA DE INFRAESTRUCTURA (Respetando Scopes)
  // SISTEMA_CLAVE: Ámbito Libro
  const key_item = workbook.getNamedItem("SISTEMA_CLAVE");
  
  // UI_FEEDBACK y UI_PREPARACION: Ámbito Hoja (active_sheet)
  const ui_feedback_item = active_sheet.getNamedItem("UI_FEEDBACK");
  const ui_prepare_item = active_sheet.getNamedItem("UI_PREPARACION");

  try {
    // 2. VALIDACIÓN DE EXISTENCIA
    if (!key_item || !ui_prepare_item || !masters_sheet) {
      throw `Error: No se encontró ${!key_item ? "SISTEMA_CLAVE" : "UI_PREPARACION"}.`;
    }

    // EXTRAEMOS LA CLAVE
    const pass_val = key_item.getRange().getValue();
    const pass_str = String(pass_val).trim();

    // 3. DESPROTECCIÓN (Usando tu solución confirmada)
    active_sheet.getProtection().unprotect(pass_str);
    is_unprotected = true;

    // 4. ACTUALIZACIÓN DE INTERFAZ (UI_PREPARACION)
    const prep_range = ui_prepare_item.getRange();
    const now = new Date();
    const time_str = now.toLocaleTimeString('es-AR', { hour12: false });
    const heart = (now.getSeconds() % 2 === 0) ? "⚡" : "✨";

    prep_range.setValue(`[${time_str}] ${heart} DATOS LISTOS. Sincronización confirmada.`);
    
    // Formato de éxito
    prep_range.getFormat().getFill().setColor("#D4EDDA"); // Verde suave
    prep_range.getFormat().getFont().setColor("#155724"); // Verde oscuro
    prep_range.getFormat().getFont().setBold(true);
    prep_range.getFormat().setWrapText(true);

    // Limpiamos feedback anterior si existe
    if (ui_feedback_item) {
        ui_feedback_item.getRange().setValue("");
        ui_feedback_item.getRange().getFormat().getFill().clear();
    }

  } catch (error) {
    console.log("Error detectado: " + error);
    
    // Si falla y tenemos el rango de feedback de hoja, avisamos al usuario
    if (ui_feedback_item && is_unprotected) {
        const f_range = ui_feedback_item.getRange();
        f_range.setValue("❌ Error: " + String(error));
        f_range.getFormat().getFill().setColor("#F8D7DA");
    }
  } finally {
    // 5. RE-PROTECCIÓN SEGURA
    if (is_unprotected && key_item) {
      const final_pass = String(key_item.getRange().getValue()).trim();
      
      active_sheet.getProtection().protect({
        allowFormatCells: false,
        allowInsertRows: false,
        allowDeleteRows: false
      }, final_pass);
      
      console.log("Hoja protegida nuevamente.");
    }
  }
}