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
  function reportarError(ws: ExcelScript.Worksheet, dir: string, texto: string, colors: typeof UX) {
    try {
        // Intentamos escribir en la hoja
        const rango = ws.getRange(dir);
        rango.setValue(texto);
        rango.getFormat().getFill().setColor(UX.ERROR_BG);
        rango.getFormat().getFont().setColor(UX.ERROR_TXT);
        rango.getFormat().setWrapText(true);
        rango.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        rango.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        rango.select(); 
    } catch (writeError) {
        // SI FALLA (porque la hoja está bloqueada y no tenemos clave):
        // No podemos mostrarlo bonito en la celda.
        // Lanzamos el error al sistema para que Excel muestre su popup lateral.
        console.log("💥 ERROR TÉCNICO (DEBUG):");
        console.log(writeError);
        throw new Error("⛔ ERROR CRÍTICO DEL SISTEMA: " + texto);
    }
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
        // 5. 
        // B) PREPARAR DATOS (Normalizado para coincidir con encabezados BD)
        const datosDiccionario: { [key: string]: string | number | boolean } = {
          "ID": idDesvio,
          "ESTADO": "Abierto",
          "FECHA SUCESO": fSuceso,
          "FECHA REGISTRO": fRegistro,
          "FECHA QA": fQA,
          "PLANTA": planta,
          "TERCERISTA": tercerista,
          "DESCRIPCIÓN": descripcion,
          "ETAPA OCURRENCIA": etapaO,
          "ETAPA DETECCIÓN": etapaD,
          "CLASIFICACIÓN": clasif,
          "IMPACTO": impacto,
          "OBSERVACIONES": obs,
          "USUARIO": autor,
          "AUDIT TRAIL": new Date().toLocaleString('es-AR', { hour12: false, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' })
        };

        // C) ESCRITURA (Validación de mapeo)
        const rangoHeaders = tablaDesvios.getHeaderRowRange(); 
        const encabezados = rangoHeaders.getValues()[0] as string[];

        // CAMBIO CLAVE: Si el mapeo falla, lanzamos error en lugar de insertar fila vacía
        const filaAInsertar = encabezados.map(col => {
            const valor = datosDiccionario[col];
            if (valor === undefined && col !== "") {
                console.log(`Error de Mapeo: La columna '${col}' no existe en el diccionario.`);
            }
            return valor ?? "";
        });

        if (filaAInsertar.filter(item => item !== "").length <= 1) {
            throw new Error("Error de Integridad: Se intentó registrar una fila vacía. Verifique que los encabezados de la BD coincidan con el script.");
        }

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
          // 1. Limpieza del mensaje para evitar "[object Object]"
          let errorTxt = "";
          
          if (typeof error === "string") {
              errorTxt = error;
          } else {
              // Intentamos leer la propiedad .message. Si no existe, usamos JSON.stringify para ver qué tiene.
              errorTxt = (error as Error).message || JSON.stringify(error);
          }

          // 2. Reportar (Esto intentará escribir en la hoja, y si falla, irá a la consola)
          reportarError(wsInput, RANGO_MENSAJES, "❌ " + errorTxt, UX);
          
        } }
      } finally {
      // 6. SEGURIDAD: Re-proteger de forma defensiva
      // Se usa el "Pattern de Silencio": Si falla es porque ya estaba protegida.
      try { if (wsBD) wsBD.getProtection().protect(undefined, String(clave)); } catch (e) {}
      try { if (wsInput) wsInput.getProtection().protect(undefined, String(clave)); } catch (e) {}
  }
}
