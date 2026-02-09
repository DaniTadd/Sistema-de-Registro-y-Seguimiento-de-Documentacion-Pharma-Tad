function main(
  workbook: ExcelScript.Workbook,
  targetTableName: string = "TablaDesvios",
  historyTableName: string = "TablaDesviosHistorial",
  inputSheetName: string = "INP_DES",
  entitySequenceCode: string = "DESVIO",
  userEmail: string = "USUARIO_TEST_MANUAL"
) {
  // --- CONFIGURACIÓN DE IDENTIDAD DEL MÓDULO ---
  const ENT = "desvío";  // El nombre de la entidad en minúsculas
  const ART = "el";      // "el" para masculino, "la" para femenino
  const GEN = "o";       // "o" para masculino (cerrado), "a" para femenino (cerrada)
  const PREF = "D-";

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

  let inputWS: ExcelScript.Worksheet | undefined,dbTab: ExcelScript.Table | undefined, histTab: ExcelScript.Table | undefined, mastersWS: ExcelScript.Worksheet | undefined, sysKey: ExcelScript.NamedItem | undefined;
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
    if (filter) {
      filter.clearCriteria();
    }
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
          inputData[clean] = (val === "" || val === null) ? (raw.endsWith("*") ? "" : "N/A") : val;
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
      // V. REGLAS DE NEGOCIO (Motor Universal)const rulesTab = mastersWS.getTable("TablaReglas");
      const valFuente = inputData; // 'inputData' es como lo llamamos en Registrar
      const vEs: string[] = [];
      // V. REGLAS DE NEGOCIO (Motor Universal)
      const rulesTab = mastersWS.getTable("TablaReglas");
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
              const vB = valFuente[kB] !== undefined ? valFuente[kB] : (hRow ? hRow[headers.indexOf(kB)] : undefined);

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
      if (vErrors.length > 0) {
        actionResult.success = false;
        actionResult.message = "⚠️ Validación:\n" + vErrors.join("\n");
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
                // Si está vacío o es N/A (esto ya lo calculamos arriba en la captura), devolvemos N/A
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

          // C. RESULTADO EXITOSO
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
    console.log(`[LOG] Éxito: ${actionResult.success}. Mensaje: ${actionResult.message}`);
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
        ws.getRange("A1").select(); r.select();
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
      if (c !== "" && c !== pk && c !== "USUARIO" && c !== "MOTIVO") {
        ws.getRangeByIndexes(i + o, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
      }
    });
  }
}