function main(
  workbook: ExcelScript.Workbook,
  entidad: string,          // Viene de PA: "DESVIO" o "CAPA"
  idDesdePanel: string,     // ID desde el panel de PA
  motivoDesdePanel: string, // Motivo desde el panel de PA
  userEmail: string
) {
  // --- LIMPIEZA DE PARÁMETRO (Anti-Error de PA) ---
  const entidadLimpia = String(entidad).replace(/[\[\]"]/g, "").toUpperCase().trim();

  // --- MAPEADOR DE CONFIGURACIÓN (El Cerebro) ---
  const MAPPING: { [key: string]: { tab: string, hist: string, inp: string, ent: string, art: string, gen: string } } = {
    "DESVIO": { tab: "TablaDesvios", hist: "TablaDesviosHistorial", inp: "INP_DES", ent: "desvío", art: "el", gen: "o" },
    "CAPA": { tab: "TablaCAPAs", hist: "TablaCAPAsHistorial", inp: "INP_CAPA", ent: "CAPA", art: "la", gen: "a" }
  };

  const config = MAPPING[entidadLimpia];
  if (!config) throw new Error(`La entidad '${entidad}' no está configurada.`);

  // --- CONFIGURACIÓN DE IDENTIDAD DINÁMICA ---
  const ENT = config.ent;
  const ART = config.art;
  const GEN = config.gen;
  const targetTableName = config.tab;
  const historyTableName = config.hist;
  const inputSheetName = config.inp;

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
    inputWS.getRange("A1").select()
    pass = sysKey.getRange().getText();
    dbTab = workbook.getTable(targetTableName);
    histTab = workbook.getTable(historyTableName);

    if (!dbTab || !histTab) throw new Error("Error: Tablas de datos o historial no encontradas.");

    // II. CAPTURA DE DATOS PARA ANULACIÓN
    inputWS.getProtection().unprotect(pass);
    const labelRange = inputWS.getRange("B:B").getUsedRange();
    const headers = dbTab.getHeaderRowRange().getValues()[0].map(h => String(h).toUpperCase().replace(/\s/g, "_"));
    const pkName = headers[0];

    // Prioridad a los datos del panel
    const searchId = idDesdePanel.trim().toUpperCase();
    const motivo = motivoDesdePanel.trim();

    // III. VALIDACIONES PREVIAS
    if (!searchId || searchId === "") {
      actionResult.success = false;
      actionResult.message = `Se necesita un ID de ${ENT} en el panel para anular.`;
      actionResult.logLevel = 'WARN';
    } else if (!motivo || motivo === "") {
      actionResult.success = false;
      actionResult.message = `No se puede anular, el motivo es obligatorio en el panel.`;
      actionResult.logLevel = 'WARN';
    } else {
      // IV. BÚSQUEDA DEL REGISTRO
      const vs = dbTab.getRangeBetweenHeaderAndTotal().getValues();
      let rIdx = -1, c = 0, found = false;

      while (c < vs.length && !found) {
        if (String(vs[c][headers.indexOf(pkName)]) === searchId) {
          rIdx = c;
          found = true;
        }
        c++;
      }

      if (!found) {
        actionResult.success = false;
        actionResult.message = `${ENT.charAt(0).toUpperCase() + ENT.slice(1)} #${searchId} no encontrad${GEN}.`;
        actionResult.logLevel = 'ERROR';
      } else if (String(vs[rIdx][headers.indexOf("ESTADO")]).toUpperCase() === "ANULADO") {
        actionResult.success = false;
        actionResult.message = `${ART.charAt(0).toUpperCase() + ART.slice(1)} ${ENT} #${searchId} ya se encuentra anulad${GEN}.`;
        actionResult.logLevel = 'INFO';
      } else {
        // V. EJECUCIÓN DE ANULACIÓN (COMMIT)
        
        // 1. Actualizar Tabla Principal
        dbTab.getWorksheet().getProtection().unprotect(pass);
        const estIdx = headers.indexOf("ESTADO");
        const audIdx = headers.indexOf("AUDIT_TRAIL");
        const usuIdx = headers.indexOf("USUARIO");

        const tableRange = dbTab.getRangeBetweenHeaderAndTotal();
        if (estIdx !== -1) tableRange.getCell(rIdx, estIdx).setValue("ANULADO");
        if (usuIdx !== -1) tableRange.getCell(rIdx, usuIdx).setValue(userEmail);
        if (audIdx !== -1) {
          tableRange.getCell(rIdx, audIdx).setValue(
            new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false })
          );
        }

        // 2. Registrar en Historial
        histTab.getWorksheet().getProtection().unprotect(pass);
        const hRow = (histTab.getHeaderRowRange().getValues()[0] as string[]).map(h => {
          const head = h.toUpperCase();
          if (head === "ID_EVENTO") {
            const col = histTab!.getColumnByName("ID_EVENTO").getRangeBetweenHeaderAndTotal();
            const v = col ? col.getValues() as number[][] : [];
            return histTab!.getRowCount() === 0 ? 1 : Math.max(...v.map(x => Number(x[0]))) + 1;
          }
          if (head === pkName) return searchId;
          if (head === "USUARIO") return userEmail;
          if (head === "MOTIVO") return motivo;
          if (head === "CAMBIOS") return "[REGISTRO ANULADO]";
          if (head === "FECHA_CAMBIO") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
          return "";
        });
        histTab.addRow(-1, hRow);

        actionResult.message = `✅ ${ENT.charAt(0).toUpperCase() + ENT.slice(1)} #${searchId} ANULAD${GEN.toUpperCase()} correctamente.`;
        actionResult.logLevel = 'EXITO';

        // Limpieza de formulario
        if (labelRange) {
          const ls = labelRange.getValues() as CellValue[][];
          const off = labelRange.getRowIndex();
          clearForm(inputWS, ls, off, pkName);
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

  function clearForm(ws: ExcelScript.Worksheet, ls: CellValue[][], o: number, pk: string) {
    ls.forEach((row, i) => {
      const c = String(row[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
      if (c !== "" && c !== pk && c !== "USUARIO") {
        ws.getRangeByIndexes(i + o, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
      }
    });
  }
}