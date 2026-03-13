function main(
  workbook: ExcelScript.Workbook,
  entidad: string,      // Viene de PA: "DESVIO" o "CAPA"
  userEmail: string     // Usuario que registra
) {
  // --- LIMPIEZA DE PARÁMETRO (Anti-Error de PA) ---
  const entidadLimpia = String(entidad).replace(/[\[\]"]/g, "").toUpperCase().trim();

  // --- MAPEADOR DE CONFIGURACIÓN (El Cerebro) ---
  const MAPPING: { [key: string]: { tab: string, hist: string, inp: string, pref: string, ent: string, art: string, gen: string } } = {
    "DESVIO": { tab: "TablaDesvios", hist: "TablaDesviosHistorial", inp: "INP_DES", pref: "D-", ent: "desvío", art: "el", gen: "o" },
    "CAPA": { tab: "TablaCapas", hist: "TablaCapasHistorial", inp: "INP_CAPAS", pref: "C-", ent: "CAPA", art: "la", gen: "a" },
    "AFECTACION": { tab: "TablaAfectacion", hist: "TablaAfectacionHistorial", inp: "INP_AFECT", pref: "AF-", ent: "AFECTACION", art: "la", gen: "a" }
  };

  const config = MAPPING[entidadLimpia];
  if (!config) throw new Error(`La entidad '${entidad}' no está configurada.`);

  // --- CONFIGURACIÓN DE IDENTIDAD DINÁMICA ---
  const ENT = config.ent;
  const ART = config.art;
  const GEN = config.gen;
  const PREF = config.pref;
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

  let inputWS: ExcelScript.Worksheet | undefined, dbTab: ExcelScript.Table | undefined, histTab: ExcelScript.Table | undefined, mastersWS: ExcelScript.Worksheet | undefined, sysKey: ExcelScript.NamedItem | undefined;
  let pass: string = "";
  let hRow: CellValue[] | undefined = undefined;

  try {
    // I. INFRAESTRUCTURA
    inputWS = workbook.getWorksheet(inputSheetName);
    mastersWS = workbook.getWorksheet("MAESTROS");
    sysKey = workbook.getNamedItem("SISTEMA_CLAVE");

    if (!inputWS) throw new Error(`Error de configuración: No se halló la hoja de entrada '${inputSheetName}'.`);
    if (!mastersWS) throw new Error("Error de configuración: Hoja MAESTROS no disponible.");
    if (!sysKey) throw new Error("Error de configuración: No se halló 'SISTEMA_CLAVE'.");

    pass = sysKey.getRange().getText();
    dbTab = workbook.getTable(targetTableName);
    
    const filter = dbTab.getAutoFilter();
    if (filter) filter.clearCriteria();
    
    histTab = workbook.getTable(historyTableName);

    if (!dbTab) throw new Error(`Error: La tabla de datos '${targetTableName}' no existe.`);
    if (!histTab) throw new Error(`Error: La tabla de historial '${historyTableName}' no existe.`);

    // II. NEGOCIO & CAPTURA DE DATOS
    inputWS.getProtection().unprotect(pass);
    const labelRange = inputWS.getRange("B:B").getUsedRange();

    if (labelRange) {
      const labels = labelRange.getValues() as CellValue[][];
      const offset = labelRange.getRowIndex();
      const inputData: { [key: string]: CellValue } = {};
      const required: string[] = [];

      labels.forEach((row, i) => {
        const raw = String(row[0]).trim().toUpperCase();
        if (raw !== "") {
          const clean = raw.replace("*", "").trim().replace(/\s/g, "_");
          if (raw.endsWith("*")) required.push(clean);
          const val = inputWS!.getRangeByIndexes(i + offset, 2, 1, 1).getValue();
          // Lógica de obligatorios y N/A
          inputData[clean] = (val === null || String(val).trim() === "") ? (raw.endsWith("*") ? "" : "N/A") : val;
        }
      });

      const headers = dbTab.getHeaderRowRange().getValues()[0].map(h => String(h).toUpperCase().replace(/\s/g, "_"));
      const pkName = headers[0];
      const vErrors: string[] = [];

      // --- 1. VALIDACIÓN DE FECHAS ---
      for (let key in inputData) {
        if (key.includes("FECHA") && inputData[key] !== "N/A" && inputData[key] !== "") {
          if (isNaN(parseDateToNum(inputData[key]))) {
            vErrors.push(`Formato de fecha inválido en: ${key.replace(/_/g, " ")}`);
          }
        }
      }

      // --- 2. MOTOR DE REGLAS DE NEGOCIO ---
      const valFuente = inputData;
      const vEs: string[] = [];
      const rulesTab = mastersWS.getTable("TablaReglas");
      if (rulesTab) {
        rulesTab.getRangeBetweenHeaderAndTotal().getValues().forEach(r => {
          if (targetTableName.toUpperCase().includes(String(r[0]).toUpperCase())) {
            const kA = String(r[1]).toUpperCase().replace(/\s/g, "_");
            const op = String(r[2]);
            const vB_Raw = String(r[3]);
            const msg = String(r[4]);
            const vA = valFuente[kA];

            if (op === "EXISTE_EN") {
              if (vA && vA !== "N/A" && vB_Raw.includes("[")) {
                const [tName, cPart] = vB_Raw.split("[");
                const cName = cPart.replace("]", "");
                const targetTab = workbook.getTable(tName);
                if (targetTab) {
                  const masterVals = targetTab.getColumnByName(cName).getRangeBetweenHeaderAndTotal().getValues();
                  const existe = masterVals.some(row => String(row[0]) === String(vA));
                  if (!existe) vEs.push(msg);
                }
              }
            } 
            else if (op === "ESTA_ABIERTO") {
              // vA: ID que escribió el usuario (ej: "DES-001")
              // vB_Raw: "BD_DESV[ID];[ESTADO];ABIERTO"
              if (vA && vA !== "N/A" && vB_Raw.includes(";")) {
                  const [targetPart, statusColPart, validValue] = vB_Raw.split(";"); 
                  // targetPart = "BD_DESV[ID]"
                  const [tName, cPart] = targetPart.split("[");
                  const idColName = cPart.replace("]", "");
                  const statusColName = statusColPart.replace("[", "").replace("]", "");

                  const targetTab = workbook.getTable(tName);
                  if (targetTab) {
                      // Obtenemos las columnas de ID y de ESTADO
                      const idVals = targetTab.getColumnByName(idColName).getRangeBetweenHeaderAndTotal().getValues();
                      const statusVals = targetTab.getColumnByName(statusColName).getRangeBetweenHeaderAndTotal().getValues();

                      // Buscamos el índice de la fila donde está nuestro ID
                      const rowIndex = idVals.findIndex(row => String(row[0]) === String(vA));

                      if (rowIndex !== -1) {
                          const estadoActual = String(statusVals[rowIndex][0]);
                          // Si el estado no es el esperado (ABIERTO), disparamos el error
                          if (estadoActual !== validValue) {
                              vEs.push(msg);
                          }
            }
        }
    }
            }
            else if (op === "<=" || op === ">=") {
              const kB = vB_Raw.toUpperCase().replace(/\s/g, "_");
              const vB = valFuente[kB];
              if (vA && vB && vA !== "N/A" && vB !== "N/A") {
                const dA = parseDateToNum(vA), dB = parseDateToNum(vB);
                if (isNaN(dA) || isNaN(dB)) {
                  vEs.push("Error: Formato de fecha inválido.");
                } else if (op === "<=" && !(dA <= dB)) {
                  vEs.push(msg);
                } else if (op === ">=" && !(dA >= dB)) {
                  vEs.push(msg);
                }
              }
            }
          }
        });
      }

      // --- 3. CAMPOS OBLIGATORIOS ---
      required.forEach(f => { if (inputData[f] === "") vErrors.push(`Falta: ${f.replace(/_/g, " ")}`); });

      // --- 4. GENERADOR DE ID CON PREFIJO (finalID) ---
      const idCol = dbTab.getColumnByName(pkName).getRangeBetweenHeaderAndTotal();
      const idValues = idCol ? idCol.getValues() as string[][] : [];
      let nextNum = 1;
      if (idValues.length > 0) {
        const nums = idValues.map(v => parseInt(String(v[0]).replace(/\D/g, '')) || 0);
        nextNum = Math.max(...nums) + 1;
      }
      const finalID = PREF + nextNum;

      // --- III. PROCESAMIENTO (COMMIT) ---
      if (vErrors.length > 0 || vEs.length > 0) {
        actionResult.success = false;
        actionResult.message = "⚠️ Validación:\n" + [...vErrors, ...vEs].join("\n");
        actionResult.logLevel = 'WARN';
      } else {
        try {
          // A. GRABACIÓN EN BASE DE DATOS
          dbTab.getWorksheet().getProtection().unprotect(pass);
          const nRow = headers.map(h => {
            if (h === pkName) return finalID;
            if (h === "ESTADO") return "ABIERTO";
            if (h === "USUARIO") return userEmail;
            if (h === "AUDIT_TRAIL") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
            if (inputData.hasOwnProperty(h)) {
              const valor = inputData[h];
              return (valor === "" || valor === null || valor === "N/A") ? "N/A" : valor;
            }
            return null;
          });
          dbTab.addRow(-1, nRow);

          // B. GRABACIÓN EN HISTORIAL
          histTab.getWorksheet().getProtection().unprotect(pass);
          hRow = (histTab.getHeaderRowRange().getValues()[0] as string[]).map(h => {
            const head = h.toUpperCase();
            if (head === "ID_EVENTO") {
              const col = histTab!.getColumnByName("ID_EVENTO").getRangeBetweenHeaderAndTotal();
              const v = col ? col.getValues() as number[][] : [];
              return histTab!.getRowCount() === 0 ? 1 : Math.max(...v.map(x => Number(x[0]))) + 1;
            }
            if (head === pkName) return finalID;
            if (head === "USUARIO") return userEmail;
            if (head === "MOTIVO") return "Registro inicial.";
            if (head === "CAMBIOS") return "[NUEVO REGISTRO CREADO]";
            if (head === "FECHA_CAMBIO") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
            return "";
          });
          histTab.addRow(-1, hRow);

          actionResult.message = `✅ ${ART.toUpperCase()} ${ENT.toUpperCase()} #${finalID} se ha cread${GEN} correctamente.`;
          actionResult.logLevel = 'EXITO';
          clearForm(inputWS, labels, offset, pkName);

        } catch (e) {
          actionResult.success = false;
          actionResult.logLevel = 'ERROR';
          actionResult.message = `❌ Error al grabar ${finalID}: ${String(e)}`;
        }
      }
    }
  } catch (err) {
    actionResult.success = false;
    actionResult.message = `❌ Sistema: ${String(err)}`;
    actionResult.logLevel = 'ERROR';
  } finally {
    if (inputWS) {
      const p: string = sysKey ? sysKey.getRange().getText() : "";
      updateUI(inputWS, actionResult, UX_COLORS, p);
      if (sysKey) {
        protect(inputWS, p, actionResult);
        if (dbTab) protect(dbTab.getWorksheet(), p, actionResult);
        if (histTab) protect(histTab.getWorksheet(), p, actionResult);
        if (mastersWS) protect(mastersWS, p, actionResult);
      }
    }
  }

  // --- FUNCIONES HELPER ---
  function parseDateToNum(v: CellValue): number {
    if (typeof v === "number") return v;
    const p = String(v).split("/");
    if (p.length === 3) {
      const dt = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
      return (dt.getFullYear() === parseInt(p[2])) ? dt.getTime() : NaN;
    }
    return NaN;
  }

  function updateUI(ws: ExcelScript.Worksheet, res: ActionResult, colors: UXMap, p: string) {
    const item = ws.getNamedItem("UI_FEEDBACK");
    const ui_prep = ws.getNamedItem("UI_PREPARACION");
    if (!item) return; 
    if (!ui_prep) return; 
    const r = item.getRange();
    const rangePrep = ui_prep.getRange();
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
      rangePrep.setValue("");
      rangePrep.getFormat().getFill().clear();
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
      if (c !== "" && c !== pk && c !== "USUARIO" && c !== "MOTIVO") {
        ws.getRangeByIndexes(i + o, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
      }
    });
  }
}