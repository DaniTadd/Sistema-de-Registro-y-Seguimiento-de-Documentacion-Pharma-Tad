function main(
  workbook: ExcelScript.Workbook,
  targetTableName: string = "TablaDesvios",
  historyTableName: string = "TablaDesviosHistorial",
  inputSheetName: string = "INP_DES",
  userEmail: string = "USUARIO_TEST_MANUAL"
) {
  // --- CONFIGURACIÓN DE IDENTIDAD DEL MÓDULO ---
  const ENT = "desvío";  // El nombre de la entidad en minúsculas
  const ART = "el";      // "el" para masculino, "la" para femenino
  const GEN = "o";       // "o" para masculino (cerrado), "a" para femenino (cerrada)

  type CellValue = string | number | boolean;
  interface ActionResult { success: boolean; message: string; logLevel: 'EXITO' | 'ERROR' | 'WARN' | 'INFO'; }
  interface UXMap { [key: string]: { bg: string; txt: string } }

  const actionResult: ActionResult = { success: true, message: "Inicio", logLevel: 'INFO' };
  const UX_COLORS: UXMap = {
    EXITO: { bg: "#D4EDDA", txt: "#155724" },
    ERROR: { bg: "#F8D7DA", txt: "#721C24" },
    WARN: { bg: "#FFF3CD", txt: "#856404" },
    INFO: { bg: "#E2E3E5", txt: "#383D41" }
  };

  let inputWS: ExcelScript.Worksheet | undefined,
    dbTab: ExcelScript.Table | undefined,
    histTab: ExcelScript.Table | undefined,
    sysKey: ExcelScript.NamedItem | undefined;
  let pass: string = "";

  try {
    // I. INFRAESTRUCTURA
    inputWS = workbook.getWorksheet(inputSheetName);
    sysKey = workbook.getNamedItem("SISTEMA_CLAVE");

    if (!inputWS) throw new Error(`Error: No se halló la hoja '${inputSheetName}'.`);
    if (!sysKey) throw new Error("Error: No se halló 'SISTEMA_CLAVE'.");

    pass = sysKey.getRange().getText();
    dbTab = workbook.getTable(targetTableName);
    histTab = workbook.getTable(historyTableName);

    if (!dbTab || !histTab) throw new Error("Error: Tablas de datos o historial no encontradas.");

    // II. CAPTURA DE DATOS DEL FORMULARIO
    inputWS.getProtection().unprotect(pass);
    const labelRange = inputWS.getRange("B:B").getUsedRange();

    if (labelRange) {
      const ls = labelRange.getValues() as CellValue[][];
      const off = labelRange.getRowIndex();
      const headers = dbTab.getHeaderRowRange().getValues()[0].map(h => String(h).toUpperCase().replace(/\s/g, "_"));
      const pkName = headers[0];

      let searchId: CellValue = "";
      let motivo: CellValue = "";

      ls.forEach((row, i) => {
        const clean = String(row[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
        if (clean === pkName) searchId = inputWS!.getRangeByIndexes(i + off, 2, 1, 1).getValue();
        if (clean === "MOTIVO") motivo = inputWS!.getRangeByIndexes(i + off, 2, 1, 1).getValue();
      });

      // III. VALIDACIONES INICIALES
      if (!searchId || searchId === "N/A") {
        actionResult.success = false;
        actionResult.message = `Se necesita un ID de ${ENT} para realizar esta acción.`;
        actionResult.logLevel = 'WARN';
      } else if (!motivo || motivo === "" || motivo === "N/A") {
        actionResult.success = false;
        actionResult.message = `El motivo es obligatorio para cambiar el estado de ${ART} ${ENT}.`;
        actionResult.logLevel = 'WARN';
      } else {
        // IV. BÚSQUEDA DEL REGISTRO
        const vs = dbTab.getRangeBetweenHeaderAndTotal().getValues();
        let rIdx = -1, c = 0, found = false;

        while (c < vs.length && !found) {
          if (String(vs[c][headers.indexOf(pkName)]) === String(searchId)) {
            rIdx = c;
            found = true;
          }
          c++;
        }

        if (!found) {
          actionResult.success = false;
          actionResult.message = `${ENT.charAt(0).toUpperCase() + ENT.slice(1)} #${searchId} no encontrad${GEN}.`;
          actionResult.logLevel = 'ERROR';
        } else {
          const currentState = String(vs[rIdx][headers.indexOf("ESTADO")]).toUpperCase();

          // 1. Bloqueo de Anulación (Inmutabilidad)
          if (currentState === "ANULADO") {
            actionResult.success = false;
            actionResult.message = `Un registro ANULADO es definitivo y no puede ser reabierto.`;
            actionResult.logLevel = 'WARN';
          } else {
            // 2. Determinación del nuevo estado (Toggle)
            const newState = (currentState === "ABIERTO") ? "CERRADO" : "ABIERTO";
            const labelAccion = (newState === "ABIERTO") ? "reapertura" : "cierre";

            // 3. COMMIT del Cambio en Tabla Principal
            dbTab.getWorksheet().getProtection().unprotect(pass);
            const estIdx = headers.indexOf("ESTADO");
            const audIdx = headers.indexOf("AUDIT_TRAIL");
            const usuIdx = headers.indexOf("USUARIO");

            dbTab.getRangeBetweenHeaderAndTotal().getCell(rIdx, estIdx).setValue(newState);
            
            // Registramos quién hizo el cambio y cuándo en la tabla principal
            if (usuIdx !== -1) dbTab.getRangeBetweenHeaderAndTotal().getCell(rIdx, usuIdx).setValue(userEmail);
            if (audIdx !== -1) {
              dbTab.getRangeBetweenHeaderAndTotal().getCell(rIdx, audIdx).setValue(
                new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false })
              );
            }

            // 4. Registro en el Historial (Audit Trail)
            histTab.getWorksheet().getProtection().unprotect(pass);
            const hiR = (histTab.getHeaderRowRange().getValues()[0] as string[]).map(h => {
              const head = h.toUpperCase();
              if (head === "ID_EVENTO") {
                const col = histTab!.getColumnByName("ID_EVENTO").getRangeBetweenHeaderAndTotal();
                const v = col ? col.getValues() as number[][] : [];
                return histTab!.getRowCount() === 0 ? 1 : Math.max(...v.map(x => Number(x[0]))) + 1;
              }
              if (head === pkName) return searchId;
              if (head === "USUARIO") return userEmail;
              if (head === "MOTIVO") return String(motivo);
              if (head === "CAMBIOS") return `ESTADO: [${currentState}] -> [${newState}]`;
              if (head === "FECHA_CAMBIO") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
              return "";
            });
            histTab.addRow(-1, hiR);

            actionResult.message = `✅ ${ENT.charAt(0).toUpperCase() + ENT.slice(1)} #${searchId} ${newState === "ABIERTO" ? "reabiert" + GEN : "cerrad" + GEN} con éxito.`;
            actionResult.logLevel = 'EXITO';
            clearForm(inputWS, ls, off, pkName);
          }
        }
      }
    }
  } catch (err) {
    actionResult.success = false;
    actionResult.message = `❌ Sistema: ${String(err)}`;
    actionResult.logLevel = 'ERROR';
  } finally {
    if (sysKey && inputWS) {
      const p = sysKey.getRange().getText();
      updateUI(inputWS, actionResult, UX_COLORS, p);
      protect(inputWS, p, actionResult);
      if (dbTab) protect(dbTab.getWorksheet(), p, actionResult);
      if (histTab) protect(histTab.getWorksheet(), p, actionResult);
    }
  }

  // --- FUNCIONES HELPER ---
  function updateUI(ws: ExcelScript.Worksheet, res: ActionResult, colors: UXMap, p: string) {
    const item = ws.getNamedItem("UI_FEEDBACK");
    if (item) {
      const r = item.getRange();
      const c = colors[res.logLevel];
      try {
        ws.getProtection().unprotect(p);
        r.setValue(res.message);
        r.getFormat().getFill().setColor(c.bg);
        r.getFormat().getFont().setColor(c.txt);
        r.getFormat().getFont().setBold(true);
        r.getFormat().setWrapText(true);
        ws.getRange("A1").select();
        r.select();
      } catch (e) { console.log("Error UI: " + e); }
    }
  }

  function protect(ws: ExcelScript.Worksheet | undefined, p: string, res: ActionResult) {
    if (ws) {
      try { ws.getProtection().protect({ allowAutoFilter: true }, p); }
      catch (e) { res.message += ` [⚠️ Seguridad: ${ws.getName()}]`; }
    }
  }

  function clearForm(ws: ExcelScript.Worksheet, ls: CellValue[][], o: number, pk: string) {
    ls.forEach((row, i) => {
      const c = String(row[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
      if (c !== "" && c !== pk && c !== "USUARIO") {
        ws.getRangeByIndexes(i + o, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
      }
    });
  }
}