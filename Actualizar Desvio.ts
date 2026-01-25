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
    ERROR_BG: "#F8D7DA", ERROR_TXT: "#721C24",
    INFO_BG: "#E2E3E5", INFO_TXT: "#383D41"
  };

  let sePuedeProcesar = true;
  let mensajeSalida = "Inicio de actualización.";
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
  // 3. PROCESO DE ACTUALIZACIÓN
  // ==========================================
  if (sePuedeProcesar) {
    try {
      // A. Preparación y Mapeo
      wsInput.getProtection().unprotect(clave);
      
      const usedRangeB = wsInput.getRange("B:B").getUsedRange();
      if (!usedRangeB) throw new Error("Formulario vacío.");

      const labelsVal = usedRangeB.getValues();
      const rowOffset = usedRangeB.getRowIndex();
      
      const inputData: { [key: string]: string | number | boolean } = {}; 
      const inputCoords: { [key: string]: number } = {}; 
      let obligatorios: string[] = [];

      labelsVal.forEach((row, i) => {
        let etiqueta = String(row[0]).trim().toUpperCase();
        if (etiqueta !== "") {
          const esObligatorio = etiqueta.endsWith("*");
          const etiquetaLimpia = esObligatorio ? etiqueta.replace("*", "").trim() : etiqueta;
          
          if (esObligatorio) obligatorios.push(etiquetaLimpia);
          
          inputCoords[etiquetaLimpia] = i + rowOffset;

          let valorRaw = wsInput.getRange("C1").getOffsetRange(i + rowOffset, 0).getValue();
          inputData[etiquetaLimpia] = (valorRaw === undefined || valorRaw === null) ? "" : (valorRaw as string | number | boolean);
        }
      });

      // B. Validaciones Críticas (Bloqueantes)
      const idBuscado = inputData["ID"];
      
      // Nota: Leemos Usuario y Motivo pero NO los validamos todavía.
      // Primero validaremos que el resto de los datos valga la pena.
      const usuarioVal = inputData["USUARIO"] ? String(inputData["USUARIO"]).trim() : "";
      const motivoVal = inputData["MOTIVO"] ? String(inputData["MOTIVO"]).trim() : "";

      let errores: string[] = [];
      if (!idBuscado) errores.push("Falta el ID del desvío.");
      
      // C. Búsqueda en BD
      const tablaBD = wsBD.getTable(TABLE_BD);
      const headersBD = tablaBD.getHeaderRowRange().getValues()[0] as string[];
      const dataBD = tablaBD.getRangeBetweenHeaderAndTotal().getValues();
      const idxID = headersBD.indexOf("ID");
      
      let rowIndex = -1;
      let filaActualBD: (string | number | boolean)[] = [];
      
      if (idxID !== -1 && idBuscado) {
        for(let k=0; k<dataBD.length; k++) {
            if (dataBD[k][idxID] == idBuscado) {
                rowIndex = k;
                filaActualBD = dataBD[k];
                break;
            }
        }
      }

      if (rowIndex === -1 && idBuscado) errores.push(`ID ${idBuscado} no existe en la Base de Datos.`);

      // D. Validación de Estado (Si está cerrado/anulado, no seguimos)
      if (rowIndex !== -1) {
        const idxEstado = headersBD.indexOf("ESTADO");
        const estadoActual = String(filaActualBD[idxEstado]).toUpperCase();
        if (estadoActual === "CERRADO" || estadoActual === "ANULADO") {
            errores.push(`El desvío está ${estadoActual} y no admite edición.`);
        }
      }

      // Si hay errores de ID o Estado, cortamos acá.
      if (errores.length > 0) {
        reportarError(wsInput, RANGO_MENSAJES, "❌ NO SE PUEDE ACTUALIZAR:\n" + errores.join("\n"));
      } else {
        // --- E. DETECCIÓN DE CAMBIOS Y LÓGICA DE NEGOCIO ---
        // Aquí es donde validamos los datos del formulario ANTES de pedir la firma.
        
        let filaHipotetica = [...filaActualBD];
        let cambiosLog: string[] = [];
        let huboCambiosReales = false;

        headersBD.forEach((header, colIndex) => {
           const h = header.toUpperCase();
           if (["ID", "ESTADO", "AUDIT TRAIL"].includes(h)) return;

           if (inputData.hasOwnProperty(h)) {
               const valorViejo = filaActualBD[colIndex];
               let valorNuevo = inputData[h];

               // Validación de campos obligatorios vacíos (DATA VALIDATION)
               if (obligatorios.includes(h) && valorNuevo === "") {
                   errores.push(`El campo ${header} es obligatorio.`);
               } else if (!obligatorios.includes(h) && valorNuevo === "") {
                   valorNuevo = "N/A";
               }

               // Comparación
               if (String(valorViejo) != String(valorNuevo)) {
                   cambiosLog.push(`[${header}: ${formatearParaLog(valorViejo, header)} -> ${formatearParaLog(valorNuevo, header)}]`);
                   filaHipotetica[colIndex] = valorNuevo;
                   huboCambiosReales = true;
               }
           }
        });

        // F. Validación de Reglas de Negocio (TablaReglas)
        const mapaHipotetico: { [key: string]: string | number | boolean } = {};
        headersBD.forEach((h, i) => { mapaHipotetico[h.toUpperCase()] = filaHipotetica[i]; });

        const tablaReglas = wsMaestros.getTable("TablaReglas");
        if (tablaReglas) {
            const reglas = tablaReglas.getRangeBetweenHeaderAndTotal().getValues();
            reglas.forEach(r => {
                const [cA, op, cB, err] = r as string[];
                const vA = mapaHipotetico[cA.toUpperCase()];
                const vB = mapaHipotetico[cB.toUpperCase()];
                if (typeof vA === "number" && typeof vB === "number") {
                    if ((op === "<" && !(vA < vB)) || (op === ">" && !(vA > vB))) errores.push(err);
                }
            });
        }

        // --- G. DECISIÓN FINAL Y VALIDACIÓN DE FIRMA ---
        
        // 1. ¿Hay errores de datos? (Prioridad 1)
        if (errores.length > 0) {
             reportarError(wsInput, RANGO_MENSAJES, "❌ ERROR EN DATOS:\n" + errores.join(" - "));
        } 
        // 2. Si los datos están bien, ¿hubo cambios? (Prioridad 2)
        else if (!huboCambiosReales) {
             const msj = wsInput.getRange(RANGO_MENSAJES);
             msj.setValue("ℹ️ No se detectaron cambios respecto a la BD.");
             msj.getFormat().getFill().setColor(UX.INFO_BG);
             msj.getFormat().getFont().setColor(UX.INFO_TXT);
             wsInput.getRange("A1").select();
             msj.select()
             
        } 
        // 3. Si hay cambios válidos, ¿Tenemos firma? (Prioridad 3)
        // Recién AHORA pedimos Motivo y Usuario.
        else {
             if (usuarioVal === "") errores.push("Falta USUARIO para firmar.");
             if (motivoVal === "") errores.push("Falta MOTIVO del cambio.");

             if (errores.length > 0) {
                 reportarError(wsInput, RANGO_MENSAJES, "⚠️ CAMBIOS VÁLIDOS DETECTADOS.\nPor favor complete: " + errores.join(" y "));
             } else {
                 // H. COMMIT (Escritura)
                 wsBD.getProtection().unprotect(clave);
                 wsHistorial.getProtection().unprotect(clave);

                 // 1. Actualizar BD
                 const idxAudit = headersBD.indexOf("AUDIT TRAIL");
                 if (idxAudit !== -1) filaHipotetica[idxAudit] = new Date().toLocaleString('es-AR', { hour12: false });
                 
                 tablaBD.getRangeBetweenHeaderAndTotal().getRow(rowIndex).setValues([filaHipotetica]);

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
                     if (key === "CAMBIOS") return cambiosLog.join(" | ");
                     return "";
                 });
                 
                 tHist.addRow(-1, rowHistData);

                 // 3. Limpieza UI
                 for (let key in inputCoords) {
                     if (key !== "ID") { 
                         wsInput.getRangeByIndexes(inputCoords[key], 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
                     }
                 }

                 // 4. Feedback
                 const msj = wsInput.getRange(RANGO_MENSAJES);
                 msj.setValue(`✅ Desvío ${idBuscado} actualizado exitosamente.`);
                 msj.getFormat().getFill().setColor(UX.EXITO_BG);
                 msj.getFormat().getFont().setColor(UX.EXITO_TXT);
                 wsInput.getRange("A1").select();
                 msj.select()
             }
        }
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

  function formatearParaLog(valor: string | number | boolean, nombreCol: string): string {
    let resultado = String(valor);
    if (typeof valor === "number" && nombreCol.toLowerCase().includes("fecha")) {
      const fechaObj = new Date(Math.round((valor - 25569) * 86400 * 1000));
      const esTimestamp = nombreCol.toUpperCase().includes("AUDIT") || nombreCol.toUpperCase().includes("QA");
      const opciones: Intl.DateTimeFormatOptions = esTimestamp 
          ? { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false }
          : { year: 'numeric', month: '2-digit', day: '2-digit' };
      resultado = fechaObj.toLocaleDateString('es-AR', opciones);
    }
    return resultado;
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