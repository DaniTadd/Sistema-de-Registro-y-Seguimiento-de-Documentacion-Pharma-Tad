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

  let sePuedeRegistrar = true;
  let mensajeSalida = "Inicio de proceso.";
  let clave = ""; // Variable accesible por los helpers de abajo

  const wsInput = workbook.getWorksheet(SHEET_INPUT);
  const wsBD = workbook.getWorksheet(SHEET_BD);
  const wsMaestros = workbook.getWorksheet(SHEET_MAESTROS);

  // ==========================================
  // 2. VALIDACIÓN INICIAL
  // ==========================================
  if (!wsInput || !wsBD || !wsMaestros) {
    sePuedeRegistrar = false;
    mensajeSalida = "❌ Error Crítico: Faltan hojas del sistema.";
  } else {
    const rangoClave = workbook.getNamedItem(NOMBRE_RANGO_CLAVE)?.getRange();
    if (rangoClave) {
        clave = rangoClave.getText();
    } else {
        sePuedeRegistrar = false;
        mensajeSalida = "❌ Error Configuración: Falta Nombre Definido 'SISTEMA_CLAVE'.";
    }
  }

  // ==========================================
  // 3. PROCESO PRINCIPAL
  // ==========================================
  if (sePuedeRegistrar) {
    try {
        wsInput.getProtection().unprotect(clave);
        const msj = wsInput.getRange(RANGO_MENSAJES);
        msj.clear(ExcelScript.ClearApplyTo.contents);
        msj.getFormat().getFill().clear();

        const usedRangeB = wsInput.getRange("B:B").getUsedRange();
        if (!usedRangeB) throw new Error("Formulario vacío (sin etiquetas).");

        const labelsVal = usedRangeB.getValues();
        const rowOffset = usedRangeB.getRowIndex();
        
        const formMap: { [key: string]: string | number | boolean } = {};
        let obligatorios: string[] = [];

        labelsVal.forEach((row, i) => {
            let label = String(row[0]).trim().toUpperCase();
            if (label !== "") {
                const esObligatorio = label.endsWith("*");
                const cleanLabel = esObligatorio ? label.replace("*", "").trim() : label;
                if (esObligatorio) obligatorios.push(cleanLabel);
                
                const rawValue = wsInput.getRange("C1").getOffsetRange(i + rowOffset, 0).getValue();
                formMap[cleanLabel] = rawValue as string | number | boolean;
            }
        });

        // Validaciones de Obligatorios
        let errores: string[] = [];
        obligatorios.forEach(campo => {
            if (formMap[campo] === undefined || formMap[campo] === "") errores.push(campo);
        });

        // Validaciones de Formato Fecha
        for (let key in formMap) {
            if (key.includes("FECHA") && formMap[key] !== "") {
                 if (typeof formMap[key] !== "number") errores.push(`${key} (Formato inválido)`);
            }
        }

        // Validaciones de Reglas de Negocio
        const tablaReglas = wsMaestros.getTable("TablaReglas");
        if (tablaReglas) {
            const reglas = tablaReglas.getRangeBetweenHeaderAndTotal().getValues();
            reglas.forEach(r => {
                const [cA, op, cB, err] = r as string[];
                const vA = formMap[cA.toUpperCase()];
                const vB = formMap[cB.toUpperCase()];
                if (typeof vA === "number" && typeof vB === "number") {
                    if ((op === "<" && !(vA < vB)) || (op === ">" && !(vA > vB))) errores.push(err);
                }
            });
        }

        if (errores.length > 0) {
            reportarError(wsInput, RANGO_MENSAJES, "❌ DATOS INVÁLIDOS:\n" + errores.join(" - "), UX);
        } else {
            // Persistencia en BD
            wsBD.getProtection().unprotect(clave);
            const tablaBD = wsBD.getTable(TABLE_BD);
            const headers = tablaBD.getHeaderRowRange().getValues()[0] as string[];
            
            // Generar ID Autonumérico
            let newId = 1;
            if (tablaBD.getRowCount() > 0) {
                const ids = tablaBD.getColumnByName("ID").getRangeBetweenHeaderAndTotal().getValues();
                newId = Math.max(...ids.map(v => Number(v[0]))) + 1;
            }

            const timeStamp = new Date().toLocaleString('es-AR', { hour12: false });
            const rowData: (string | number | boolean)[] = headers.map((h: string) => {
                const key = h.toUpperCase();
                if (key === "ID") return newId;
                if (key === "ESTADO") return "Abierto";
                if (key === "AUDIT TRAIL") return timeStamp;
                const valor = formMap[key];
                return (valor !== undefined && valor !== "") ? valor : "N/A";
            });

            tablaBD.addRow(-1, rowData);

            // Éxito
            usedRangeB.getOffsetRange(0, 1).clear(ExcelScript.ClearApplyTo.contents); 
            msj.setValue(`✅ Registro #${newId} guardado con éxito.`);
            msj.getFormat().getFill().setColor(UX.EXITO_BG);
            msj.getFormat().getFont().setColor(UX.EXITO_TXT);
            msj.getFormat().getFont().setBold(true);
            wsInput.getRange("A1").select();
        }

    } catch (e) {
        const errorTxt: string = (e as Error).message || JSON.stringify(e);
        mensajeSalida = `❌ Excepción: ${errorTxt}`;
        reportarError(wsInput, RANGO_MENSAJES, mensajeSalida, UX);
    } finally {
        // Cierre Seguro Estandarizado (Sin pasar clave)
        safeProtect(wsBD, "BD");
        safeProtect(wsInput, "Input");
    }
  } else {
      // Si falló la validación inicial, reportamos si es posible
      if (wsInput) reportarError(wsInput, RANGO_MENSAJES, mensajeSalida, UX);
  }

  // ==========================================
  // ZONA DE HELPERS (Al final)
  // ==========================================
  
  function reportarError(ws: ExcelScript.Worksheet, dir: string, texto: string, colors: typeof UX) {
    try {
      const rango = ws.getRange(dir);
      rango.setValue(texto);
      rango.getFormat().getFill().setColor(UX.ERROR_BG);
      rango.getFormat().getFont().setColor(UX.ERROR_TXT);
      rango.getFormat().setWrapText(true);
      ws.getRange("A1").select();
    } catch (e) {
      console.log("Error UI: " + JSON.stringify(e));
    }
  }
  
  // Helper que toma 'clave' del scope principal
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