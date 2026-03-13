function main(
  workbook: ExcelScript.Workbook,
  entidad: string,      // Viene de PA: "DESVIO" o "CAPA"
  idDesdePanel: string, // El ID que el usuario escribe en el panel
  userEmail: string     // Para auditoría si fuera necesario
) {
  // --- MAPEADOR DE CONFIGURACIÓN (El Cerebro) ---
  const entidadLimpia = String(entidad).replace(/[\[\]"]/g, "").toUpperCase().trim();
  const MAPPING: { [key: string]: { tab: string, sheet: string, ent: string, art: string, gen: string } } = {
    "DESVIO": { tab: "TablaDesvios", sheet: "INP_DES", ent: "desvío", art: "el", gen: "o" },
    "CAPA": { tab: "TablaCapas", sheet: "INP_CAPAS", ent: "CAPA", art: "la", gen: "a" }
  };

  const config = MAPPING[entidadLimpia.toUpperCase()];
  if (!config) throw new Error(`La entidad '${entidad}' no está configurada en el mapeador.`);

  // --- CONFIGURACIÓN DE IDENTIDAD DINÁMICA ---
  const ENT = config.ent;
  const ART = config.art;
  const GEN = config.gen;
  const targetTableName = config.tab;
  const inputSheetName = config.sheet;

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

    if (!dbTab) throw new Error(`Error: La tabla '${targetTableName}' no existe.`);

    const filter = dbTab.getAutoFilter();
    if (filter) filter.clearCriteria();

    // II. CAPTURA DEL ID A BUSCAR (Prioridad Panel)
    inputWS.getProtection().unprotect(pass);
    const searchId = idDesdePanel.trim().toUpperCase();

    if (!searchId || searchId === "") {
      actionResult.success = false;
      actionResult.message = `Se necesita un ID de ${ENT} en el panel para continuar.`;
      actionResult.logLevel = 'WARN';
    } else {
      
      const labelRange = inputWS.getRange("B:B").getUsedRange();
      const headers = dbTab.getHeaderRowRange().getValues()[0].map(h => String(h).toUpperCase().replace(/\s/g, "_"));
      const pkName = headers[0];

      if (labelRange) {
        const labels = labelRange.getValues() as CellValue[][];
        const offset = labelRange.getRowIndex();
        const coords: { [key: string]: number } = {};

        // Mapeamos posiciones del formulario
        labels.forEach((row, i) => {
          const clean = String(row[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
          if (clean !== "") coords[clean] = i + offset;
        });

        // III. PROCESO DE BÚSQUEDA
        const vals = dbTab.getRangeBetweenHeaderAndTotal().getValues();
        const texts = dbTab.getRangeBetweenHeaderAndTotal().getTexts();
        let rIdx = -1, c = 0, found = false;

        while (c < vals.length && !found) {
          if (String(vals[c][headers.indexOf(pkName)]) === searchId) {
            rIdx = c;
            found = true;
          }
          c++;
        }

        if (found) {
          // 1. Limpiamos el formulario antes de cargar
          labels.forEach((row, i) => {
            const clean = String(row[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
            if (clean !== "") {
              inputWS!.getRangeByIndexes(i + offset, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
            }
          });

          // 2. Poblamos el formulario con los datos de la tabla
          headers.forEach((h, col) => {
            if (coords[h] !== undefined) {
              const range = inputWS!.getRangeByIndexes(coords[h], 2, 1, 1);

              if (h.includes("FECHA")) {
                range.setValue(vals[rIdx][col]); 
                range.setNumberFormatLocal("dd/mm/aaaa");
              } else {
                range.setValue(texts[rIdx][col]);
              }
            }
          });

          // 3. Forzamos que el ID del panel quede escrito en el formulario por seguridad visual
          if (coords[pkName] !== undefined) {
            inputWS!.getRangeByIndexes(coords[pkName], 2, 1, 1).setValue(searchId);
          }

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

  // --- FUNCIONES HELPER (SIN CAMBIOS) ---
  function updateUI(ws: ExcelScript.Worksheet, res: ActionResult, colors: UXMap, p: string) {
    const item = ws.getNamedItem("UI_FEEDBACK");
    if (!item) return; 
    const r = item.getRange();
    const c = colors[res.logLevel];
    const d = new Date();
    const timeStr = d.toLocaleTimeString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
    const heartbeat = (d.getSeconds() % 2 === 0) ? "⚡" : "✨";
    try {
      ws.getProtection().unprotect(p);
      r.setValue(`[${timeStr}] ${heartbeat} ${res.message}`);
      r.getFormat().getFill().setColor(c.bg);
      r.getFormat().getFont().setColor(c.txt);
      r.getFormat().getFont().setBold(true);
      r.getFormat().setWrapText(true);
    } catch (e) {
      try { r.setValue(res.message); } catch (e2) {}
    }
  }

  function protect(ws: ExcelScript.Worksheet | undefined, p: string, res: ActionResult) {
    if (ws) {
      try { ws.getProtection().protect({ allowAutoFilter: true }, p); }
      catch (e) { res.message += ` [⚠️ Seguridad: ${ws.getName()}]`; }
    }
  }
}