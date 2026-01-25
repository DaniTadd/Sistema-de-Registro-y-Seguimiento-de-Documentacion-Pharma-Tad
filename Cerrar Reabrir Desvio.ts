function main(workbook: ExcelScript.Workbook) {
  // ==========================================
  // 1. CONFIGURACIÓN Y CONSTANTES
  // ==========================================
  const SHEET_INPUT = "INPUT_DESVIOS";
  const SHEET_BD = "BD_DESVIOS";
  const SHEET_HISTORIAL = "HISTORIAL_DESVIOS";
  const SHEET_MAESTROS = "MAESTROS";
  
  const TABLE_BD = "TablaDesvios";
  const TABLE_HISTORIAL = "TablaHistorialDesvios";
  
  const RANGO_MENSAJES = "D1:F3";
  const NOMBRE_RANGO_CLAVE = "SISTEMA_CLAVE";

  const UX = {
    EXITO_BG: "#D4EDDA", EXITO_TXT: "#155724",
    ERROR_BG: "#F8D7DA", ERROR_TXT: "#721C24"
  };

  let sePuedeProcesar = true;
  let mensajeSalida = "Inicio de proceso de estado.";
  let clave = ""; // Accesible por los helpers

  const wsInput = workbook.getWorksheet(SHEET_INPUT);
  const wsBD = workbook.getWorksheet(SHEET_BD);
  const wsHistorial = workbook.getWorksheet(SHEET_HISTORIAL);
  const wsMaestros = workbook.getWorksheet(SHEET_MAESTROS);

  // ==========================================
  // 2. VALIDACIÓN DE ENTORNO
  // ==========================================
  if (!wsInput || !wsBD || !wsHistorial || !wsMaestros) {
    sePuedeProcesar = false;
    mensajeSalida = "❌ Error Crítico: Faltan hojas del sistema.";
  } else {
    const rangoClave = workbook.getNamedItem(NOMBRE_RANGO_CLAVE)?.getRange();
    if (rangoClave) {
        clave = rangoClave.getText();
    } else {
        sePuedeProcesar = false;
        mensajeSalida = "❌ Error Configuración: Falta Nombre Definido 'SISTEMA_CLAVE'.";
    }
  }

  // ==========================================
  // 3. PROCESO DE CAMBIO DE ESTADO
  // ==========================================
  if (sePuedeProcesar) {
    try {
      // A. Preparación y Lectura Input
      wsInput.getProtection().unprotect(clave);
      
      const usedRangeB = wsInput.getRange("B:B").getUsedRange();
      if (!usedRangeB) throw new Error("Formulario vacío.");

      const labelsVal = usedRangeB.getValues();
      const rowOffset = usedRangeB.getRowIndex();
      
      // Mapas
      const inputData: { [key: string]: string | number | boolean } = {};
      const inputCoords: { [key: string]: number } = {};

      labelsVal.forEach((row, i) => {
        let etiqueta = String(row[0]).trim().toUpperCase();
        if (etiqueta !== "") {
          if (etiqueta.endsWith("*")) etiqueta = etiqueta.replace("*", "").trim();
          
          inputCoords[etiqueta] = i + rowOffset;
          
          let valorRaw = wsInput.getRange("C1").getOffsetRange(i + rowOffset, 0).getValue();
          inputData[etiqueta] = (valorRaw === undefined || valorRaw === null) ? "" : (valorRaw as string | number | boolean);
        }
      });

      // B. Extracción de Datos Críticos
      const idBuscado = inputData["ID"];
      const usuarioVal = inputData["USUARIO"] ? String(inputData["USUARIO"]).trim() : "";
      const motivoVal = inputData["MOTIVO"] ? String(inputData["MOTIVO"]).trim() : "";

      // C. Validaciones Previas
      let errores: string[] = [];

      if (!idBuscado) errores.push("Falta el ID del desvío.");
      if (usuarioVal === "") errores.push("El usuario es obligatorio para firmar el cambio de estado.");
      if (motivoVal === "") errores.push("El motivo es obligatorio para justificar el cambio de estado.");

      // D. Búsqueda en BD
      const tablaBD = wsBD.getTable(TABLE_BD);
      const headersBD = tablaBD.getHeaderRowRange().getValues()[0] as string[];
      const dataBD = tablaBD.getRangeBetweenHeaderAndTotal().getValues();
      const idxID = headersBD.indexOf("ID");
      const idxEstado = headersBD.indexOf("ESTADO");

      let rowIndex = -1;
      let estadoActual = "";

      if (idxID !== -1 && idBuscado) {
        for(let k=0; k<dataBD.length; k++) {
            if (dataBD[k][idxID] == idBuscado) {
                rowIndex = k;
                estadoActual = String(dataBD[k][idxEstado]); 
                break;
            }
        }
      }

      if (rowIndex === -1 && idBuscado) errores.push(`ID ${idBuscado} no existe en la Base de Datos.`);

      // E. Lógica de Negocio (Toggle de Estado)
      let estadoNuevo = "";
      
      if (rowIndex !== -1) {
          const estadoUpper = estadoActual.toUpperCase();
          
          if (estadoUpper === "ANULADO") {
              errores.push("El registro está ANULADO. No se puede Cerrar/Reabrir.");
          } else if (estadoUpper === "CERRADO") {
              estadoNuevo = "Abierto";
          } else {
              // Asumimos que si no es Cerrado ni Anulado, está Abierto y lo cerramos
              estadoNuevo = "Cerrado";
          }
      }

      // F. Decisión Final
      if (errores.length > 0) {
          reportarError(wsInput, RANGO_MENSAJES, "❌ ERROR:\n" + errores.join("\n"));
      } else {
          // G. COMMIT (Escritura)
          wsBD.getProtection().unprotect(clave);
          wsHistorial.getProtection().unprotect(clave);

          // 1. Actualizar BD (Solo columna Estado)
          tablaBD.getRangeBetweenHeaderAndTotal().getCell(rowIndex, idxEstado).setValue(estadoNuevo);

          // 2. Insertar en Historial
          const tHist = wsHistorial.getTable(TABLE_HISTORIAL);
          let nuevoIdEvento = 1;
          if (tHist.getRowCount() > 0) {
               const colIdEv = tHist.getColumnByName("ID_EVENTO").getRangeBetweenHeaderAndTotal().getValues();
               nuevoIdEvento = Math.max(...colIdEv.map(v => Number(v[0]))) + 1;
          }

          const headersHist = tHist.getHeaderRowRange().getValues()[0] as string[];
          
          const rowHistData: (string | number | boolean)[] = headersHist.map((h: string) => {
             const key = h.toUpperCase();
             if (key === "ID_EVENTO") return nuevoIdEvento;
             if (key === "ID_DESVIO") return idBuscado;
             if (key === "FECHA_CAMBIO") return new Date().toLocaleString('es-AR', { hour12: false });
             if (key === "USUARIO") return usuarioVal;
             if (key === "MOTIVO") return motivoVal;
             if (key === "CAMBIOS") return `[ESTADO: ${estadoActual} -> ${estadoNuevo}]`;
             return "";
          });

          tHist.addRow(-1, rowHistData);

          // 3. Limpieza UI (Quirúrgica)
          for (let key in inputCoords) {
             if (key !== "ID") { 
                 wsInput.getRangeByIndexes(inputCoords[key], 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
             }
          }

          // 4. Feedback
          const msj = wsInput.getRange(RANGO_MENSAJES);
          msj.setValue(`✅ Desvío ${idBuscado} cambió a: ${estadoNuevo.toUpperCase()}.`);
          msj.getFormat().getFill().setColor(UX.EXITO_BG);
          msj.getFormat().getFont().setColor(UX.EXITO_TXT);
          wsInput.getRange("A1").select();
      }

    } catch (e) {
      mensajeSalida = `❌ Excepción Técnica: ${(e as Error).message}`;
      reportarError(wsInput, RANGO_MENSAJES, mensajeSalida);
    } finally {
      // Cierre Seguro Universal (Check-Before-Act)
      safeProtect(wsBD, "BD");
      safeProtect(wsHistorial, "Historial"); 
      safeProtect(wsInput, "Input");
    }
  } else {
      console.log(mensajeSalida);
  }

  // ==========================================
  // ZONA DE HELPERS (Al final)
  // ==========================================

  function reportarError(ws: ExcelScript.Worksheet, dir: string, texto: string) {
    try {
      const rango = ws.getRange(dir);
      rango.setValue(texto);
      rango.getFormat().getFill().setColor(UX.ERROR_BG);
      rango.getFormat().getFont().setColor(UX.ERROR_TXT);
      rango.getFormat().setWrapText(true);
      rango.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
      rango.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
      ws.getRange("A1").select();
    } catch (e) {
      console.log(`❌ Error UI: ${texto}`);
    }
  }

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