function main(workbook: ExcelScript.Workbook) {
  // 1. CONFIGURACIÓN Y CONSTANTES
  const SHEET_INPUT = "INPUT_DESVIOS";
  const SHEET_BD = "BD_DESVIOS";
  const SHEET_MAESTROS = "MAESTROS";
  const TABLE_BD = "TablaDesvios"; 
  // 1. A) Coordenadas Exactas
  const COORD = {
    FECHA_SUCESO: "C6",
    FECHA_REGISTRO: "C7",
    FECHA_QA: "C8",
    PLANTA: "C10",
    TERCERISTA: "C12",
    DESC: "C14",
    ETAPA_OCURRENCIA: "C16",
    ETAPA_DETECCION: "C18",
    CLASIFICACION: "C20",
    IMPACTO: "C22",
    OBSERVACIONES: "C24",
    AUTOR: "C26",
    MOTIVO: "C28"
  };

  const RANGO_MENSAJES = "D1:F3"; 
  const CELL_CLAVE = "XFD1";

  // 1. B) Colores UX
  const UX = {
    EXITO_BG: "#D4EDDA",
    EXITO_TXT: "#155724",
    ERROR_BG: "#F8D7DA",
    ERROR_TXT: "#721C24"
  };

  // 1. C) FUNCIÓN HELPER ENCAPSULADA
  // Al estar dentro de main, no choca con otros scripts
  function reportarError(ws: ExcelScript.Worksheet, dir: string, texto: string, colors: typeof UX ) {
    const rango = ws.getRange(dir);
    rango.setValue(texto);
    rango.getFormat().getFill().setColor(UX.ERROR_BG);
    rango.getFormat().getFont().setColor(UX.ERROR_TXT);
    rango.getFormat().setWrapText(true);
    rango.select()
  }
  // ------------------------------------------------------------

  const wsInput = workbook.getWorksheet(SHEET_INPUT)!;
  const wsBD = workbook.getWorksheet(SHEET_BD)!;
  const wsMaestros = workbook.getWorksheet(SHEET_MAESTROS)!;
  if (!wsInput || !wsBD || !wsMaestros) {
      throw new Error("Faltan hojas críticas (INPUT, BD o MAESTROS).");
  }

  // 2. LIMPIEZA INICIAL DE UX
  let clave = "";
  try {
    // 2. A) Lectura de la clave antes que nada
    clave = wsMaestros.getRange(CELL_CLAVE).getText(); 
    // 2. B) Desprotección de Input desvíos para limpiar  y escribir errores.
    wsInput.getProtection().unprotect(clave)

    // 2. C) Limpieza visual
    const msj = wsInput.getRange(RANGO_MENSAJES);
    msj.clear(ExcelScript.ClearApplyTo.contents);
    msj.getFormat().getFill().clear();
  } catch (e) {}
  
  try {
    // 3. LECTURA DE DATOS
    
    const fSuceso = wsInput.getRange(COORD.FECHA_SUCESO).getValue();
    const fRegistro = wsInput.getRange(COORD.FECHA_REGISTRO).getValue();
    const fQA = wsInput.getRange(COORD.FECHA_QA).getValue();
    const planta = wsInput.getRange(COORD.PLANTA).getText();
    const tercerista = wsInput.getRange(COORD.TERCERISTA).getText();
    const descripcion = wsInput.getRange(COORD.DESC).getText();
    const etapaO = wsInput.getRange(COORD.ETAPA_OCURRENCIA).getText();
    const etapaD = wsInput.getRange(COORD.ETAPA_DETECCION).getText();
    const clasif = wsInput.getRange(COORD.CLASIFICACION).getText();
    const impacto = wsInput.getRange(COORD.IMPACTO).getText();
    const obs = wsInput.getRange(COORD.OBSERVACIONES).getText();
    const autor = wsInput.getRange(COORD.AUTOR).getText();

    // 4. VALIDACIONES
    let errores: string[] = [];

    if (!fSuceso) errores.push("Fecha Suceso.");
    if (!fRegistro) errores.push("Fecha Registro.");
    if (!planta) errores.push("Planta.");
    if (!tercerista) errores.push("Tercerista.");
    if (!descripcion) errores.push("Descripción.");
    if (!etapaO) errores.push("Etapa de Ocurrencia.");
    if (!etapaD) errores.push("Etapa de Detección.");
    if (!clasif) errores.push("Clasificación.");
    if (!impacto) errores.push("Impacto.");
    if ((!obs) && obs != "N/A") errores.push("Observaciones.");
    if (!autor) errores.push("Autor.");

    if (fSuceso && fRegistro && fRegistro < fSuceso) {
        errores.push("⛔ F. Registro es anterior a F. Suceso.");
    }
    
    // Nota: Si fQA está vacía (opcional en algunos procesos), no comparamos.

    if (fRegistro && fQA && fQA < fRegistro) {
        errores.push("⛔ F. Recepción QA es anterior a F. Registro.");
    }

    if (errores.length > 0) {
      reportarError(wsInput, RANGO_MENSAJES, "❌ FALTAN DATOS:\n" + errores.join(" - "), UX);
    } else {

      // 5. ESCRITURA SEGURA
      try {
        // 5. A) Descprotección de BD_DESVIOS
        wsBD.getProtection().unprotect(String(clave));
        const tablaDesvios = wsBD.getTable(TABLE_BD);
        if (!tablaDesvios) throw new Error("Falta TablaDesvios.");

        let nuevoId = 1;
        let cantidadFilas = tablaDesvios.getRowCount();
        if (cantidadFilas > 0) {
          let valoresID = tablaDesvios.getColumnByName("ID").getRangeBetweenHeaderAndTotal().getValues();
          let listaNumeros = valoresID.map(fila => Number(fila[0]));
          nuevoId = Math.max(...listaNumeros) + 1;
        }
        const idDesvio = nuevoId;                
        // 5. B) PREPARAR DATOS
        const datosDiccionario: { [key: string]: string | number | boolean } = {
          "ID": idDesvio,
          "Estado": "Abierto",
          "Fecha Suceso": fSuceso,
          "Fecha Registro": fRegistro,
          "Fecha QA": fQA,
          "Planta": planta,
          "Tercerista": tercerista,
          "Descripción": descripcion,
          "Etapa Ocurrencia": etapaO,
          "Etapa Detección": etapaD,
          "Clasificación": clasif,
          "Impacto": impacto,
          "Observaciones": obs,
          "Usuario": autor,
          "Audit Trail": new Date().toLocaleString()
        };

        // 5. C) ESCRITURA         
        // Lectura de los encabezados de la Fila 0
        const rangoHeaders = tablaDesvios.getHeaderRowRange(); 
        // Se obtienen los valores directos
        const encabezados = rangoHeaders.getValues()[0] as string[];

        // Mapeo de los datos
        const filaAInsertar = encabezados.map(col => datosDiccionario[col] ?? "");

        tablaDesvios.addRow(-1, filaAInsertar);

        const rangoTabla = tablaDesvios.getRange();
          rangoTabla.getFormat().setWrapText(true);
          rangoTabla.getFormat().autofitColumns();
          rangoTabla.getFormat().autofitRows();

        // 5. D) LIMPIEZA Y ÉXITO
        wsInput.getRange("C6:C28").clear(ExcelScript.ClearApplyTo.contents); 

        const msj = wsInput.getRange(RANGO_MENSAJES);
        msj.setValue(`✅ Desvío ${idDesvio} registrado correctamente.`);
        msj.getFormat().getFill().setColor(UX.EXITO_BG);
        msj.getFormat().getFont().setColor(UX.EXITO_TXT);
        msj.getFormat().getFont().setBold(true);
        msj.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        msj.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        msj.select()

      } catch (error) {
        reportarError(wsInput, RANGO_MENSAJES, "❌ ERROR CRÍTICO:\n" + error, UX);
      } }
      } finally {
      // 6. SEGURIDAD: Re-proteger de forma defensiva
      // Se usa el "Pattern de Silencio": Si falla es porque ya estaba protegida.
      try { if (wsBD) wsBD.getProtection().protect(undefined, String(clave)); } catch (e) {}
      try { if (wsInput) wsInput.getProtection().protect(undefined, String(clave)); } catch (e) {}
  }
}
