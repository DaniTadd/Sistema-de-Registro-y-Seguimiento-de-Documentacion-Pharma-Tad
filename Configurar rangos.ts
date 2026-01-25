function main(workbook: ExcelScript.Workbook) {
  // ==========================================
  // 1. CONSTANTES Y CONFIGURACIÓN
  // ==========================================
  const SHEET_INPUT = "INPUT_DESVIOS";
  const SHEET_CONFIG = "MAESTROS";
  const NOMBRE_RANGO_CLAVE = "SISTEMA_CLAVE"; 

  const wsInput = workbook.getWorksheet(SHEET_INPUT);
  const wsMaestros = workbook.getWorksheet(SHEET_CONFIG);
  
  let clave = "";
  let mensajeLog = "";

  // ==========================================
  // 2. VALIDACIÓN DE ENTORNO
  // ==========================================
  if (!wsInput || !wsMaestros) {
    mensajeLog = `⛔ Error: Faltan hojas críticas (${SHEET_INPUT} o ${SHEET_CONFIG}).`;
  } else {
    const rangoClave = workbook.getNamedItem(NOMBRE_RANGO_CLAVE)?.getRange();
    if (rangoClave) {
        clave = rangoClave.getText();
        if (clave === "") mensajeLog = "⚠️ Advertencia: La clave de sistema está vacía.";
    } else {
        mensajeLog = `⛔ Error: No se encontró la configuración '${NOMBRE_RANGO_CLAVE}'.`;
    }
  }
  
  // ==========================================
  // 3. EJECUCIÓN PRINCIPAL
  // ==========================================
  if (wsInput && clave !== "") {
    try {
        // A) Reset de Seguridad
        wsInput.getProtection().unprotect(clave);
        wsInput.getRange().getFormat().getProtection().setLocked(true); 

        // B) Lógica de Desbloqueo Selectivo
        let total = procesarColumna("B", 2, 1); 
        total += procesarColumna("E", 5, 4);    

        mensajeLog = total > 0 
            ? `✅ Configuración exitosa. ${total} campos habilitados.` 
            : "⚠️ Proceso finalizado sin campos habilitados.";

    } catch (e) {
        mensajeLog = `❌ Error de Ejecución: ${(e as Error).message}`;
    } finally {
        // Cierre estandarizado
        safeProtect(wsInput, "Input");
    }
  }

  console.log(mensajeLog);

  // ==========================================
  // ZONA DE HELPERS (Al final)
  // ==========================================

  function procesarColumna(letraEtiqueta: string, colInputIdx: number, filasIgnorar: number): number {
      
      let count = 0;
      // AGREGADO (!): wsInput! le dice al sistema "Te juro que existe".
      const usedRange = wsInput!.getRange(`${letraEtiqueta}:${letraEtiqueta}`).getUsedRange();
      
      if (usedRange) {
          const valores = usedRange.getValues();
          const filaInicial = usedRange.getRowIndex();

          valores.forEach((fila, i) => {
              const etiqueta = String(fila[0]).trim();
              const filaReal = filaInicial + i;

              if (etiqueta !== "" && filaReal > filasIgnorar) { 
                  // AGREGADO (!): Aquí también usamos wsInput!
                  wsInput!.getRangeByIndexes(filaReal, colInputIdx, 1, 1)
                         .getFormat().getProtection().setLocked(false);
                  count++;
              }
          });
      }
      return count;
  }

  // Helper 2: Cierre Seguro Estandarizado
  function safeProtect(ws: ExcelScript.Worksheet | undefined, name: string) {
      if (ws) {
          try {
              ws.getProtection().protect({}, clave);
          } catch (e) {
              const errString = JSON.stringify(e);
              if (!errString.includes("InvalidOperation")) {
                  console.log(`ℹ️ Aviso Cierre (${name}): ${errString}`);
              }
          }
      }
  }

}