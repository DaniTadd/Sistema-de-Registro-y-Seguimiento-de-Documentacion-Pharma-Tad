function main(workbook: ExcelScript.Workbook) {
  // 1. CONSTANTES
  const SHEET_INPUT = "INPUT_DESVIOS";
  const SHEET_CONFIG = "MAESTROS";
  const CELL_CLAVE = "XFD1"; 
  const RANGO_INPUTS = "C1:C30"; // Estado + Formulario

  const wsInput = workbook.getWorksheet(SHEET_INPUT);
  if (!wsInput) throw new Error(`⛔ No se encontró la hoja ${SHEET_INPUT}`);

  // 2. OBTENER CLAVE (MODO ESTRICTO)
  const wsMaestros = workbook.getWorksheet(SHEET_CONFIG);
  
  // Validación 1: ¿Existe la hoja de configuración?
  if (!wsMaestros) {
    throw new Error(`⛔ ERROR CRÍTICO DE SEGURIDAD: Falta la hoja '${SHEET_CONFIG}'. No se puede configurar el sistema.`);
  }

  // Validación 2: ¿Hay clave configurada?
  // Usamos .getText() para leerla tal cual es
  const clave = wsMaestros.getRange(CELL_CLAVE).getText();

  if (!clave || clave === "") {
     throw new Error("⛔ ERROR CRÍTICO: La celda de contraseña (XFD1) está vacía en MAESTROS.");
  }

  // 3. APLICAR CONFIGURACIÓN
  // A) Intentar desproteger (con la clave REAL)
  try {
    wsInput.getProtection().unprotect(String(clave));
  } catch (e) {
    console.log("Aviso: La hoja ya estaba desprotegida o la clave cambió.");
  }

  // B) Bloquear TODO (Reset)
  wsInput.getRange().getFormat().getProtection().setLocked(true);

  // C) Desbloquear SOLO los inputs (C3 a C28)
  const inputs = wsInput.getRange(RANGO_INPUTS);
  inputs.getFormat().getProtection().setLocked(false);
  
  // D) Proteger la hoja (Sin opciones extra, usando la clave REAL)
  wsInput.getProtection().protect(undefined, String(clave));

  console.log(`✅ SEGURIDAD APLICADA. Inputs desbloqueados en: ${RANGO_INPUTS}`);
}