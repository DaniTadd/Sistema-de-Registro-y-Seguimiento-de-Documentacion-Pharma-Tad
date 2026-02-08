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
    mastersWS: ExcelScript.Worksheet | undefined,
    sysKey: ExcelScript.NamedItem | undefined;
  let pass: string = "";

  try {
    // I. INFRAESTRUCTURA
    inputWS = workbook.getWorksheet(inputSheetName);
    mastersWS = workbook.getWorksheet("MAESTROS");
    sysKey = workbook.getNamedItem("SISTEMA_CLAVE");

    if (!inputWS) throw new Error(`Error de configuración: No se halló la hoja de entrada '${inputSheetName}'.`);
    if (!mastersWS) throw new Error("Error de configuración: Hoja MAESTROS no disponible.");
    if (!sysKey) throw new Error("Error de configuración: No se halló el rango 'SISTEMA_CLAVE'.");

    pass = sysKey.getRange().getText();
    dbTab = workbook.getTable(targetTableName);
    histTab = workbook.getTable(historyTableName);

    if (!dbTab) throw new Error(`Error de configuración: La tabla '${targetTableName}' no existe.`);
    if (!histTab) throw new Error(`Error de configuración: La tabla '${historyTableName}' no existe.`);

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
          data[c] = (val === "" || val === null) ? (raw.endsWith("*") ? "" : "N/A") : val;
        }
      });

      const headers = dbTab.getHeaderRowRange().getValues()[0].map(h => String(h).toUpperCase().replace(/\s/g, "_"));
      const pkName = headers[0];
      const searchId = data[pkName];

      if (!searchId || searchId === "N/A") {
        actionResult.success = false;
        actionResult.message = `Se necesita un ID de ${ENT} para continuar.`;
        actionResult.logLevel = 'WARN';
      } else {
        // III. BÚSQUEDA Y VALIDACIÓN DE ESTADO
        const vs = dbTab.getRangeBetweenHeaderAndTotal().getValues();
        const ts = dbTab.getRangeBetweenHeaderAndTotal().getTexts();
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
          const estadoActual = String(vs[rIdx][headers.indexOf("ESTADO")]).toUpperCase();

          if (["CERRADO", "ANULADO"].includes(estadoActual)) {
            actionResult.success = false;
            actionResult.message = `⚠️ No se puede modificar un${ART === "la" ? "a" : ""} ${ENT} que se encuentra en estado ${estadoActual}.`;
            actionResult.logLevel = 'WARN';
          } else {
            // IV. COMPARACIÓN DE CAMBIOS (LOGICA QUIRÚRGICA)
            const hRow = [...vs[rIdx]];
            const log: string[] = [];
            const vEs: string[] = [];
            const cambiosABase: { col: number, val: CellValue }[] = [];

            headers.forEach((h, colIdx) => {
              if (["AUDIT_TRAIL", "ESTADO", "USUARIO"].includes(h)) return;
              if (data.hasOwnProperty(h)) {
                const oT = ts[rIdx][colIdx]; // Texto actual en la tabla
                const nV = data[h];         // Valor nuevo en el formulario
                let nT = "";

                // Tratamiento de fechas (Búfer de 12hs + 2-digit)
                if (typeof nV === "number" && h.includes("FECHA")) {
                  const ms = Math.round((nV - 25569) * 86400 * 1000) + 43200000;
                  nT = new Date(ms).toLocaleDateString('es-AR', {
                    timeZone: 'America/Argentina/Buenos_Aires',
                    day: '2-digit',
                    month: '2-digit',
                    year: 'numeric'
                  });
                } else {
                  nT = String(nV);
                }

                if (oT.trim() !== nT.trim()) {
                  log.push(`${h}: [${oT}] -> [${nT}]`);
                  hRow[colIdx] = nV;
                  cambiosABase.push({ col: colIdx, val: nV });
                }
              }
            });

            // V. REGLAS DE NEGOCIO (Motor Universal)
            const rulesTab = mastersWS.getTable("TablaReglas");
            const valFuente = data; // 'data' es como lo llamamos en Actualizar
            // V. REGLAS DE NEGOCIO (Motor Universal)
            if (rulesTab) {
              rulesTab.getRangeBetweenHeaderAndTotal().getValues().forEach(r => {
                if (targetTableName.toUpperCase().includes(String(r[0]).toUpperCase())) {
                  const kA = String(r[1]).toUpperCase().replace(/\s/g, "_");
                  const op = String(r[2]);
                  const vB_Raw = String(r[3]);
                  const msg = String(r[4]);

                  // Usamos valFuente (el puente que definimos arriba)
                  const vA = valFuente[kA];

                  // --- CASO 1: EXISTE EN TABLA EXTERNA ---
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
                  
                  // --- CASO 2: COMPARACIÓN DE FECHAS ---
                  else if (op === "<=" || op === ">=") {
                    const kB = vB_Raw.toUpperCase().replace(/\s/g, "_");
                    // Buscamos vB en el formulario. Si no está, lo buscamos en la fila de la tabla (solo en Actualizar)
                    const vB = valFuente[kB] !== undefined ? valFuente[kB] : (typeof hRow !== 'undefined' ? hRow[headers.indexOf(kB)] : undefined);

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
            } else if (!data["MOTIVO"] || data["MOTIVO"] === "N/A" || data["MOTIVO"] === "") {
              actionResult.success = false;
              actionResult.message = `⚠️ Por favor, ingresa el motivo de la modificación de ${ART} ${ENT}.`;
              actionResult.logLevel = 'WARN';
            } else {
              // VII. EJECUCIÓN DEL COMMIT
              dbTab.getWorksheet().getProtection().unprotect(pass);
              
              // Actualizamos Auditoría y Usuario en Tabla Principal
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
                if (head === "MOTIVO") return String(data["MOTIVO"]);
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
      if (c !== "" && c !== pk && c !== "MOTIVO") {
        ws.getRangeByIndexes(i + o, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
      }
    });
  }
}