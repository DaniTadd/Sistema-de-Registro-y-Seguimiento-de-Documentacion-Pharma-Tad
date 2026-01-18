function main(workbook: ExcelScript.Workbook) {
  // 1. CONFIGURACIÓN Y CONSTANTES
  const SHEET_INPUT = "INPUT_DESVIOS";
  const SHEET_BD = "BD_DESVIOS";
  const SHEET_HISTORIAL = "HISTORIAL_DESVIOS";
  const SHEET_MAESTROS = "MAESTROS";
  const TABLE_BD = "TablaDesvios";
  const TABLE_HISTORIAL = "TablaHistorialDesvios";
  const CELL_CLAVE = "XFD1";
  const RANGO_MENSAJES = "D1:F3";

  const UX = {
    EXITO_BG: "#D4EDDA",
    EXITO_TXT: "#155724",
    ERROR_BG: "#F8D7DA",
    ERROR_TXT: "#721C24"
  };

  // --- FUNCIÓN DE REPORTE CONSISTENTE (Según Registrar/Buscar) ---
  function reportarError(ws: ExcelScript.Worksheet, dir: string, texto: string, colors: typeof UX) {
    try {
        const rango = ws.getRange(dir);
        rango.setValue(texto);
        rango.getFormat().getFill().setColor(colors.ERROR_BG);
        rango.getFormat().getFont().setColor(colors.ERROR_TXT);
        rango.getFormat().setWrapText(true);
        rango.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        rango.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        rango.select(); 
    } catch (writeError) {
        console.log("💥 ERROR TÉCNICO (DEBUG):", writeError);
        throw new Error("⛔ ERROR CRÍTICO DEL SISTEMA: " + texto);
    }
  }

  /**
   * Helper: Conversión de formatos para el Audit Trail (Legibilidad Humana).
   * Evita que las fechas se vean como números de serie de Excel.
   */
  function formatearParaLog(valor: string | number | boolean, nombreCol: string): string {
    if (typeof valor === "number" && nombreCol.toLowerCase().includes("fecha")) {
      const fecha = new Date(Math.round((valor - 25569) * 86400 * 1000));
      return fecha.toLocaleDateString('es-AR', { hour12: false, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' }); 
    }
    return String(valor);
  }

  const wsInput = workbook.getWorksheet(SHEET_INPUT);
  const wsBD = workbook.getWorksheet(SHEET_BD);
  const wsHistorial = workbook.getWorksheet(SHEET_HISTORIAL);
  const wsMaestros = workbook.getWorksheet(SHEET_MAESTROS);

  let clave: string = "";

  if (wsInput && wsBD && wsHistorial && wsMaestros) {
    try {
      // 2. PREPARACIÓN Y SEGURIDAD
      clave = wsMaestros.getRange(CELL_CLAVE).getText();
      wsInput.getProtection().unprotect(clave); 

      // 3. CAPTURA DINÁMICA POR ETIQUETAS (Columna B -> Columna C)
      // Eliminamos rigidez de índices: Buscamos por el nombre en la celda de al lado.
      const rangoForm = wsInput.getRange("B6:C28");
      const valoresForm = rangoForm.getValues();
      const formMap: { [key: string]: { valor: string | number | boolean, celda: ExcelScript.Range } } = {};
      
      valoresForm.forEach((fila, index) => {
        const etiqueta = String(fila[0]).trim();
        if (etiqueta) { 
          formMap[etiqueta] = { valor: fila[1], celda: rangoForm.getCell(index, 1) }; 
        }
      });

      const idBuscado = wsInput.getRange("C2").getValue() as number;
      const motivoVal = formMap["MOTIVO"] ? String(formMap["MOTIVO"].valor).trim() : "";
      const usuarioVal = formMap["USUARIO"] ? String(formMap["USUARIO"].valor).trim() : "";

      const tablaBD = wsBD.getTable(TABLE_BD);
      const dataBD = tablaBD.getRangeBetweenHeaderAndTotal().getValues();
      const headersBD = tablaBD.getHeaderRowRange().getValues()[0] as string[];

      let rowIndex = -1;
      for (let i = 0; i < dataBD.length; i++) {
        if (dataBD[i][headersBD.indexOf("ID")] == idBuscado) { rowIndex = i; }
      }

      // 4. VALIDACIONES Y BANDERAS GMP
      let msgError = "";
      if (rowIndex === -1) {
        msgError = "ID no encontrado en la base de datos.";
      } else if (dataBD[rowIndex][headersBD.indexOf("ESTADO")] === "Cerrado") {
        msgError = "Registro Cerrado. No se permiten modificaciones.";
      } else if (usuarioVal === "") {
        msgError = "Falta identificar el 'Usuario' para la firma del cambio (Atribución).";
      } else {
        // Temporalidad: fSuceso <= fRegistro <= fQA
        const fS = formMap["FECHA SUCESO"]?.valor as number;
        const fR = formMap["FECHA REGISTRO"]?.valor as number;
        const fQ = formMap["FECHA QA"]?.valor as number;
        if (fS && fR && fR < fS) msgError = "Fecha Registro no puede ser anterior a Fecha Suceso.";
        else if (fR && fQ && fQ < fR) msgError = "Fecha QA no puede ser anterior a Fecha Registro.";
      }

      // 5. DETECCIÓN DE CAMBIOS
      let cambios: string[] = [];
      let nuevaFila = rowIndex !== -1 ? [...dataBD[rowIndex]] : [];
      
      if (msgError === "" && rowIndex !== -1) {
        const columnasAComparar = headersBD.filter(h => !["ID", "ESTADO", "AUDIT TRAIL"].includes(h));
        columnasAComparar.forEach(col => {
          if (formMap[col]) {
            const bdIdx = headersBD.indexOf(col);
            const vViejo = dataBD[rowIndex][bdIdx];
            const vNuevo = formMap[col].valor;

            // Validación de integridad: No vaciar campos que ya tenían datos
            if (vNuevo === "" && vViejo !== "" && col != "FECHA QA") { msgError = `El campo [${col}] no puede quedar vacío.`; }
            
            if (String(vViejo) !== String(vNuevo)) {
              cambios.push(`[${col}: ${formatearParaLog(vViejo, col)} -> ${formatearParaLog(vNuevo, col)}]`);
              nuevaFila[bdIdx] = vNuevo;
            }
          }
        });
      }

      // 6. EJECUCIÓN (COMMIT)
      if (msgError !== "") {
        reportarError(wsInput, RANGO_MENSAJES, "❌ " + msgError, UX);
      } else if (cambios.length === 0) {
        const msj = wsInput.getRange(RANGO_MENSAJES);
        msj.setValue("ℹ️ Sin cambios detectados respecto a la BD.");
        msj.getFormat().getFill().setColor("#E2E3E5");
        msj.getFormat().getFont().setColor("#383D41");
      } else if (motivoVal === "") {
        reportarError(wsInput, RANGO_MENSAJES, "⚠️ Se requiere 'Motivo' para justificar el cambio.", UX);
      } else {
        wsBD.getProtection().unprotect(clave);
        wsHistorial.getProtection().unprotect(clave);

        // A. Update Base de Datos
        const rFila = tablaBD.getRangeBetweenHeaderAndTotal().getRow(rowIndex);
        const idxAudit = headersBD.indexOf("AUDIT TRAIL");
        if (idxAudit !== -1) nuevaFila[idxAudit] = new Date().toLocaleString('es-AR', { hour12: false, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' });
        rFila.setValues([nuevaFila]);
        
        // B. Insert Historial (Cálculo de ID_EVENTO)
        const tHist = wsHistorial.getTable(TABLE_HISTORIAL);
        const dataHist = tHist.getRangeBetweenHeaderAndTotal().getValues();
        let proximoId = 1;
        if (dataHist.length > 0) {
            proximoId = Math.max(...dataHist.map(f => Number(f[0]))) + 1;
        }

        const datosHist: { [key: string]: string | number | boolean } = {
            "ID_EVENTO": proximoId,
            "ID_DESVIO": idBuscado,
            "FECHA_CAMBIO": new Date().toLocaleString('es-AR', { hour12: false, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' }),
            "USUARIO": usuarioVal,
            "MOTIVO": motivoVal,
            "CAMBIOS": cambios.join(" | ")
        };

        const headersHist = tHist.getHeaderRowRange().getValues()[0] as string[];
        tHist.addRow(-1, headersHist.map(h => datosHist[h] ?? ""));

        // C. Mensaje de Éxito (Consistente con Registrar/Buscar)
        const msj = wsInput.getRange(RANGO_MENSAJES);
        msj.setValue(`✅ Desvío ${idBuscado} actualizado y auditado correctamente.`);
        msj.getFormat().getFill().setColor(UX.EXITO_BG);
        msj.getFormat().getFont().setColor(UX.EXITO_TXT);
        msj.getFormat().getFont().setBold(true);
        msj.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        msj.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        msj.select();

        // Limpieza dinámica del motivo
        if (formMap["MOTIVO"]) formMap["MOTIVO"].celda.clear(ExcelScript.ClearApplyTo.contents);
      }

    } catch (e) {
      reportarError(wsInput, RANGO_MENSAJES, "❌ Error técnico de ejecución.", UX);
      console.log(e);
    } finally {
      // 7. BLINDAJE FINAL
      if (clave !== "") {
        [wsInput, wsBD, wsHistorial].forEach(ws => {
          try { ws.getProtection().protect(undefined, clave); } catch (err) {}
        });
      }
    }
  }
}