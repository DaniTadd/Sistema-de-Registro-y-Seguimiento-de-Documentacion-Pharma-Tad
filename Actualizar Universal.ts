function main(
  workbook: ExcelScript.Workbook,
  entidad: string,          // Viene de PA: "DESVIO" o "CAPA"
  idDesdePanel: string,     // ID escrito en el panel
  motivoDesdePanel: string, // Motivo escrito en el panel
  userEmail: string
) {
  // --- LIMPIEZA DE PARÁMETRO (Anti-Error de PA) ---
  const entidadLimpia = String(entidad).replace(/[\[\]"]/g, "").toUpperCase().trim();

  // --- MAPEADOR DE CONFIGURACIÓN (El Cerebro) ---
  const MAPPING: { [key: string]: { tab: string, hist: string, inp: string, ent: string, art: string, gen: string } } = {
    "DESVIO": { tab: "TablaDesvios", hist: "TablaDesviosHistorial", inp: "INP_DES", ent: "desvío", art: "el", gen: "o" },
    "CAPA": { tab: "TablaCapas", hist: "TablaCapasHistorial", inp: "INP_CAPAS", ent: "CAPA", art: "la", gen: "a" }
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
    mastersWS: ExcelScript.Worksheet | undefined,
    sysKey: ExcelScript.NamedItem | undefined;
  let pass: string = "";

  try {
    // I. INFRAESTRUCTURA
    inputWS = workbook.getWorksheet(inputSheetName);
    mastersWS = workbook.getWorksheet("MAESTROS");
    sysKey = workbook.getNamedItem("SISTEMA_CLAVE");

    if (!inputWS) throw new Error(`Error de configuración: No se halló la hoja '${inputSheetName}'.`);
    if (!mastersWS) throw new Error("Error de configuración: Hoja MAESTROS no disponible.");
    if (!sysKey) throw new Error("Error de configuración: No se halló 'SISTEMA_CLAVE'.");

    pass = sysKey.getRange().getText();
    dbTab = workbook.getTable(targetTableName);
    histTab = workbook.getTable(historyTableName);

    if (!dbTab) throw new Error(`Error: La tabla '${targetTableName}' no existe.`);
    if (!histTab) throw new Error(`Error: La tabla '${historyTableName}' no existe.`);

    // II. CAPTURA DE DATOS DEL FORMULARIO
    inputWS.getProtection().unprotect(pass);
    const labelRange = inputWS.getRange("B:B").getUsedRange();

    if (labelRange) {
      const ls = labelRange.getValues() as CellValue[][];
      const off = labelRange.getRowIndex();
      const data: { [key: string]: CellValue } = {};
      const req: string[] = [];

      ls.forEach((row, i) => {
        const raw = String(row[0]).trim().toUpperCase();
        if (raw !== "") {
          const c = raw.replace("*", "").trim().replace(/\s/g, "_");
          if (raw.endsWith("*")) req.push(c);
          const val = inputWS!.getRangeByIndexes(i + off, 2, 1, 1).getValue();
          data[c] = (val === null || String(val).trim() === "") ? (raw.endsWith("*") ? "" : "N/A") : val;
        }
      });

      const headers = dbTab.getHeaderRowRange().getValues()[0].map(h => String(h).toUpperCase().replace(/\s/g, "_"));
      const pkName = headers[0];
      
      // PRIORIDAD DE ID DESDE PANEL
      
      // --- SALVAGUARDA DE INTEGRIDAD ESTRICTA ---
      const idEnPanel = idDesdePanel.trim().toUpperCase();
      const idEnHoja = String(data[pkName] || "").trim().toUpperCase();

      // REGLA DE ORO: Deben existir ambos y ser idénticos para avanzar.
      if (idEnPanel === "" || idEnHoja === "" || idEnHoja === "N/A" || idEnPanel !== idEnHoja) {
          actionResult.success = false;
          actionResult.message = `⚠️ ERROR: Mismatch de ID. Panel: [${idEnPanel || "VACÍO"}] | Hoja: [${idEnHoja || "VACÍO"}]. Busque el registro en Excel antes de actualizar.`;
          actionResult.logLevel = 'ERROR';
          throw new Error(actionResult.message);
      }

      // Si pasó el filtro, el ID es 100% confiable
      const searchId = idEnPanel;

      if (!searchId || searchId === "" || searchId === "N/A") {
        actionResult.success = false;
        actionResult.message = `Se necesita un ID de ${ENT} para continuar.`;
        actionResult.logLevel = 'WARN';
      } else {
        // III. BÚSQUEDA Y VALIDACIÓN DE ESTADO
        const vs = dbTab.getRangeBetweenHeaderAndTotal().getValues();
        const ts = dbTab.getRangeBetweenHeaderAndTotal().getTexts();
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
        } else {
          const estadoActual = String(vs[rIdx][headers.indexOf("ESTADO")]).toUpperCase();

          if (["CERRADO", "ANULADO"].includes(estadoActual)) {
            actionResult.success = false;
            actionResult.message = `⚠️ No se puede modificar un${ART} ${ENT} en estado ${estadoActual}.`;
            actionResult.logLevel = 'WARN';
          } else {
            // IV. COMPARACIÓN DE CAMBIOS
            const hRow = [...vs[rIdx]];
            const log: string[] = [];
            const vEs: string[] = [];
            const cambiosABase: { col: number, val: CellValue }[] = [];

            // Validación Quirúrgica de Obligatorios
            req.forEach(f => {
              if (data[f] === "" || data[f] === null) vEs.push(`Falta campo obligatorio: ${f.replace(/_/g, " ")}`);
            });

            headers.forEach((h, colIdx) => {
              if (["AUDIT_TRAIL", "ESTADO", "USUARIO"].includes(h)) return;
              if (data.hasOwnProperty(h)) {
                const oT = ts[rIdx][colIdx]; 
                const nV = data[h];         
                let nT = "";

                if (typeof nV === "number" && h.includes("FECHA")) {
                  const ms = Math.round((nV - 25569) * 86400 * 1000) + 43200000;
                  nT = new Date(ms).toLocaleDateString('es-AR', {
                    timeZone: 'America/Argentina/Buenos_Aires',
                    day: '2-digit', month: '2-digit', year: 'numeric'
                  });
                } else {
                  nT = String(nV);
                }

                if (oT.trim() !== nT.trim()) {
                  log.push(`${h}: [${oT}] -> [${nT}]`);
                  cambiosABase.push({ col: colIdx, val: nV });
                }
              }
            });

            // V. REGLAS DE NEGOCIO
            const rulesTab = mastersWS.getTable("TablaReglas");
            const valFuente = data; 
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
                  else if (op === "<=" || op === ">=") {
                    const kB = vB_Raw.toUpperCase().replace(/\s/g, "_");
                    const vB = valFuente[kB] !== undefined ? valFuente[kB] : hRow[headers.indexOf(kB)];
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

            // VI. DECISIÓN DE COMMIT
            if (vEs.length > 0) {
              actionResult.success = false;
              actionResult.message = "⚠️ " + vEs.join(" | ");
              actionResult.logLevel = 'WARN';
            } else if (log.length === 0) {
              actionResult.message = "ℹ️ No se detectaron cambios para guardar.";
              actionResult.logLevel = 'INFO';
            } else if (!motivoDesdePanel || motivoDesdePanel.trim() === "") {
              actionResult.success = false;
              actionResult.message = `⚠️ Por favor, ingresa el motivo de la modificación en el panel.`;
              actionResult.logLevel = 'WARN';
            } else {
              // VII. EJECUCIÓN DEL COMMIT
              dbTab.getWorksheet().getProtection().unprotect(pass);
              const filaRango = dbTab.getRangeBetweenHeaderAndTotal().getRow(rIdx);
              cambiosABase.forEach(c => filaRango.getCell(0, c.col).setValue(c.val));
              
              const aI = headers.indexOf("AUDIT_TRAIL");
              if (aI !== -1) filaRango.getCell(0, aI).setValue(new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false }));
              const uI = headers.indexOf("USUARIO");
              if (uI !== -1) filaRango.getCell(0, uI).setValue(userEmail);

              // Grabación en Historial
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
                if (head === "MOTIVO") return motivoDesdePanel.trim();
                if (head === "CAMBIOS") return log.join(" | ");
                if (head === "FECHA_CAMBIO") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
                return "";
              });
              histTab.addRow(-1, hiR);

              actionResult.message = `✅ ${searchId} actualizad${GEN} con éxito.`;
              actionResult.logLevel = 'EXITO';
              clearForm(inputWS, ls, off, pkName);
            }
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

  // --- FUNCIONES HELPER (SIN CAMBIOS) ---
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
      if (c !== "" && c !== pk && c !== "MOTIVO") {
        ws.getRangeByIndexes(i + o, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
      }
    });
  }
}