function main(workbook: ExcelScript.Workbook) {
  // ==========================================
  // 1. CONFIGURACIÓN Y CONSTANTES
  // ==========================================
  const SHEET_INPUT = "INPUT_DESVIOS";
  const SHEET_BD = "BD_DESVIOS";
  const SHEET_MAESTROS = "MAESTROS";
  const TABLE_BD = "TablaDesvios";
  const RANGO_MENSAJES = "D1:F3";
  const NOMBRE_RANGO_CLAVE = "SISTEMA_CLAVE";

  const UX = {
    EXITO_BG: "#D4EDDA", EXITO_TXT: "#155724",
    ERROR_BG: "#F8D7DA", ERROR_TXT: "#721C24"
  };

  let sePuedeBuscar = true;
  let mensajeSalida = "Inicio de búsqueda.";
  let clave = ""; // Accesible por helpers

  const wsInput = workbook.getWorksheet(SHEET_INPUT);
  const wsBD = workbook.getWorksheet(SHEET_BD);
  const wsMaestros = workbook.getWorksheet(SHEET_MAESTROS);

  // ==========================================
  // 2. VALIDACIÓN DE ENTORNO
  // ==========================================
  if (!wsInput || !wsBD || !wsMaestros) {
    sePuedeBuscar = false;
    mensajeSalida = "❌ Error Crítico: Faltan hojas del sistema.";
  } else {
    const rangoClave = workbook.getNamedItem(NOMBRE_RANGO_CLAVE)?.getRange();
    if (rangoClave) {
      clave = rangoClave.getText();
    } else {
      sePuedeBuscar = false;
      mensajeSalida = "❌ Error Configuración: Falta Nombre Definido 'SISTEMA_CLAVE'.";
    }
  }

  // ==========================================
  // 3. PROCESO DE BÚSQUEDA
  // ==========================================
  if (sePuedeBuscar) {
    try {
      // A. Preparación de UI
      wsInput.getProtection().unprotect(clave);
      const msj = wsInput.getRange(RANGO_MENSAJES);
      msj.clear(ExcelScript.ClearApplyTo.contents);
      msj.getFormat().getFill().clear();

      // B. Lectura del Mapa (Input)
      const usedRangeB = wsInput.getRange("B:B").getUsedRange();
      
      if (!usedRangeB) {
        mensajeSalida = "⚠️ Formulario vacío (sin etiquetas en Columna B).";
        reportarError(wsInput, RANGO_MENSAJES, mensajeSalida);
      } else {
        const labelsVal = usedRangeB.getValues();
        const rowOffset = usedRangeB.getRowIndex();
        
        // Buscamos dónde está el ID y qué valor tiene
        let idBuscado: string | number | boolean | undefined;
        let filaID = -1;

        // Mapa de coordenadas: { "ETIQUETA": IndiceFila }
        const mapaCoordenadas: { [key: string]: number } = {};

        labelsVal.forEach((row, i) => {
          let etiqueta = String(row[0]).trim().toUpperCase();
          if (etiqueta !== "") {
            // Limpieza de asteriscos para búsqueda
            if (etiqueta.endsWith("*")) etiqueta = etiqueta.replace("*", "").trim();
            
            mapaCoordenadas[etiqueta] = i + rowOffset;

            if (etiqueta === "ID") {
              filaID = i + rowOffset;
              // Columna C corresponde al índice 2
              idBuscado = wsInput.getRangeByIndexes(filaID, 2, 1, 1).getValue() as string | number | boolean; 
            }
          }
        });

        // C. Validación de Input ID
        if (!idBuscado || idBuscado === "") {
           mensajeSalida = "⚠️ Por favor ingresa un número de ID para buscar.";
           reportarError(wsInput, RANGO_MENSAJES, mensajeSalida);
        } else {
           // D. Búsqueda en Base de Datos
           // Nota: Para buscar, no necesitamos desproteger la BD, solo leerla.
           const tablaBD = wsBD.getTable(TABLE_BD);
           const headers = tablaBD.getHeaderRowRange().getValues()[0] as string[];
           const dataBody = tablaBD.getRangeBetweenHeaderAndTotal().getValues();
           
           const indexColID = headers.indexOf("ID");
           const filaEncontrada = dataBody.find((row) => row[indexColID] == idBuscado);

           if (!filaEncontrada) {
             mensajeSalida = `⚠️ No se encontró el Desvío #${idBuscado} en la base de datos.`;
             reportarError(wsInput, RANGO_MENSAJES, mensajeSalida);
           } else {
             // E. Volcado de Datos (Mapping Inverso: BD -> Input)
             
             // 1. Limpiamos campos actuales (excepto ID)
             for (let label in mapaCoordenadas) {
               if (label !== "ID") {
                 wsInput.getRangeByIndexes(mapaCoordenadas[label], 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
               }
             }

             // 2. Llenamos con datos de BD
             let camposCargados = 0;
             headers.forEach((header, colIndex) => {
               const h = header.toUpperCase();
               if (mapaCoordenadas[h] !== undefined && h !== "ID") {
                 const valorBD = filaEncontrada[colIndex];
                 const targetRow = mapaCoordenadas[h];
                 wsInput.getRangeByIndexes(targetRow, 2, 1, 1).setValue(valorBD);
                 camposCargados++;
               }
             });

             // F. Éxito
             mensajeSalida = `✅ Desvío #${idBuscado} cargado (${camposCargados} campos).`;
             msj.setValue(mensajeSalida);
             msj.getFormat().getFill().setColor(UX.EXITO_BG);
             msj.getFormat().getFont().setColor(UX.EXITO_TXT);
             wsInput.getRange("A1").select();
           }
        }
      }

    } catch (e) {
      mensajeSalida = `❌ Error Técnico: ${(e as Error).message}`;
      reportarError(wsInput, RANGO_MENSAJES, mensajeSalida);
    } finally {
        // En Búsqueda no usamos Historial.
        safeProtect(wsBD, "BD");
        safeProtect(wsInput, "Input");
    }
  } else {
    // Si falló la validación inicial
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
      console.log(`❌ Error UI: No se pudo escribir en pantalla. Msg: ${texto}`);
    }
  }

  function safeProtect(ws: ExcelScript.Worksheet | undefined, name: string) {
    if (ws) {
      try {
        // Intentamos proteger. Si falla con "InvalidOperation", es que ya estaba protegida.
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