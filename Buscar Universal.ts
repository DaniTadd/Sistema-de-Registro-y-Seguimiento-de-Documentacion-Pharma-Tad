function main(
  workbook: ExcelScript.Workbook,
  targetTableName: string = "TablaDesvios",
  inputSheetName: string = "INP_DES"
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
    const filter = dbTab.getAutoFilter();
    if (filter) {
      filter.clearCriteria();
    }

    if (!dbTab) throw new Error(`Error: La tabla '${targetTableName}' no existe.`);

    // II. CAPTURA DEL ID A BUSCAR
    inputWS.getProtection().unprotect(pass);
    const labelRange = inputWS.getRange("B:B").getUsedRange();
    const headers = dbTab.getHeaderRowRange().getValues()[0].map(h => String(h).toUpperCase().replace(/\s/g, "_"));
    const pkName = headers[0];

    if (labelRange) {
      const labels = labelRange.getValues() as CellValue[][];
      const offset = labelRange.getRowIndex();
      const coords: { [key: string]: number } = {};
      let searchId: CellValue = "";

      // Mapeamos posiciones y capturamos el ID del formulario
      labels.forEach((row, i) => {
        const clean = String(row[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
        if (clean !== "") {
          coords[clean] = i + offset;
          if (clean === pkName) {
            searchId = inputWS!.getRangeByIndexes(i + offset, 2, 1, 1).getValue();
          }
        }
      });

      if (!searchId || searchId === "N/A") {
        actionResult.success = false;
        actionResult.message = `Se necesita un ID de ${ENT} para continuar.`;
        actionResult.logLevel = 'WARN';
      } else {
        // III. PROCESO DE BÚSQUEDA
        const vals = dbTab.getRangeBetweenHeaderAndTotal().getValues();
        const texts = dbTab.getRangeBetweenHeaderAndTotal().getTexts();
        let rIdx = -1, c = 0, found = false;

        while (c < vals.length && !found) {
          if (String(vals[c][headers.indexOf(pkName)]) === String(searchId)) {
            rIdx = c;
            found = true;
          }
          c++;
        }

        if (found) {
          // 1. Limpiamos el formulario antes de cargar (excepto el ID)
          labels.forEach((row, i) => {
            const clean = String(row[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
            if (clean !== pkName && clean !== "") {
              inputWS!.getRangeByIndexes(i + offset, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
            }
          });

          // 2. Poblamos el formulario con los datos de la tabla
          headers.forEach((h, col) => {
            if (coords[h] !== undefined && h !== pkName) {
              const range = inputWS!.getRangeByIndexes(coords[h], 2, 1, 1);

              // Lógica Quirúrgica para Fechas: Mantenemos el valor numérico + formato
              if (h.includes("FECHA")) {
                range.setValue(vals[rIdx][col]); 
                range.setNumberFormatLocal("dd/mm/yyyy");
              } else {
                range.setValue(texts[rIdx][col]);
              }
            }
          });

          actionResult.message = `✅ ${ENT.charAt(0).toUpperCase() + ENT.slice(1)} #${searchId} cargad${GEN} con éxito.`;
          actionResult.logLevel = 'EXITO';

        } else {
          actionResult.success = false;
          actionResult.message = `${ENT.charAt(0).toUpperCase() + ENT.slice(1)} #${searchId} no existe en la base de datos.`;
          actionResult.logLevel = 'ERROR';
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
}