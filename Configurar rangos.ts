function main(
  workbook: ExcelScript.Workbook,
  inputSheetName: string = "INP_DES", // Parametrizado
  mastersSheetName: string = "MAESTROS" // Parametrizado
) {
  // --- CONFIGURACIÓN DE IDENTIDAD ---
  const ENT = "desvío"; 
  const ART = "el";
  const NOMBRE_RANGO_CLAVE = "SISTEMA_CLAVE"; 

  const wsInput = workbook.getWorksheet(inputSheetName);
  const wsMaestros = workbook.getWorksheet(mastersSheetName);
  
  let clave = "";
  let mensajeLog = "";

  // I. VALIDACIÓN DE ENTORNO
  if (!wsInput || !wsMaestros) {
    mensajeLog = `⛔ Error: Faltan hojas críticas (${inputSheetName} o ${mastersSheetName}).`;
  } else {
    const rangoClave = workbook.getNamedItem(NOMBRE_RANGO_CLAVE)?.getRange();
    if (rangoClave) {
      clave = rangoClave.getText();
      if (clave === "") mensajeLog = "⚠️ Advertencia: La clave de sistema está vacía.";
    } else {
      mensajeLog = `⛔ Error: No se encontró el rango '${NOMBRE_RANGO_CLAVE}'.`;
    }
  }
  
  // II. EJECUCIÓN PRINCIPAL
  if (wsInput && clave !== "") {
    try {
      // A) Reset de Seguridad: Bloqueamos TODO primero
      wsInput.getProtection().unprotect(clave);
      wsInput.getRange().getFormat().getProtection().setLocked(true); 

      // B) Lógica de Desbloqueo Selectivo
      // procesarColumna(LetraEtiqueta, IndiceColumnaInput, FilasAIgnorar)
      let total = procesarColumna("B", 2, 1); // Desbloquea columna C
      total += procesarColumna("E", 5, 4);    // Desbloquea columna F (si hubiera)

      mensajeLog = total > 0 
        ? `✅ Configuración de ${ENT} exitosa. ${total} campos habilitados para entrada.` 
        : `⚠️ Proceso finalizado sin campos habilitados en ${inputSheetName}.`;

    } catch (e) {
      mensajeLog = `❌ Error de Ejecución en ${ENT}: ${(e as Error).message}`;
    } finally {
      // Cierre estandarizado alineado con el resto del sistema
      safeProtect(wsInput, `Input ${ENT.charAt(0).toUpperCase() + ENT.slice(1)}`);
    }
  }

  console.log(mensajeLog);

  // --- HELPERS INTERNOS ---

  function procesarColumna(letraEtiqueta: string, colInputIdx: number, filasIgnorar: number): number {
    let count = 0;
    const usedRange = wsInput!.getRange(`${letraEtiqueta}:${letraEtiqueta}`).getUsedRange();
    
    if (usedRange) {
      const valores = usedRange.getValues();
      const filaInicial = usedRange.getRowIndex();

      valores.forEach((fila, i) => {
        const etiqueta = String(fila[0]).trim();
        const filaReal = filaInicial + i;

        // Si la etiqueta no está vacía y no es una fila de encabezado a ignorar
        if (etiqueta !== "" && filaReal > filasIgnorar) { 
          wsInput!.getRangeByIndexes(filaReal, colInputIdx, 1, 1)
                   .getFormat().getProtection().setLocked(false);
          count++;
        }
      });
    }
    return count;
  }

  function safeProtect(ws: ExcelScript.Worksheet | undefined, name: string) {
    if (ws) {
      try {
        // Alineado con el resto de los scripts: permitimos Autofiltros
        ws.getProtection().protect({ allowAutoFilter: true }, clave);
      } catch (e) {
        console.log(`ℹ️ Aviso Cierre (${name}): Hoja ya protegida o error menor.`);
      }
    }
  }
}