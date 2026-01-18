function main(workbook: ExcelScript.Workbook) {
  // 1. CONFIGURACIÓN Y CONSTANTES
  // 1. A) Hojas
  const SHEET_INPUT = "INPUT_DESVIOS";
  const SHEET_BD = "BD_DESVIOS";
  const SHEET_MAESTROS = "MAESTROS";

  // 1. B) Rango de Mensajes UX
  const RANGO_MENSAJES = "D1:F3";
  const CELL_CLAVE = "XFD1";

  // 1. C) Mapa de Mapeo: "Nombre Columna en BD" => "Celda en INPUT"
  const MAPA_LECTURA: { [key: string]: string } = {
    "FECHA SUCESO": "C6",
    "ESTADO": "C3",
    "FECHA REGISTRO": "C7",
    "FECHA QA": "C8",
    "PLANTA": "C10",
    "TERCERISTA": "C12",
    "DESCRIPCIÓN": "C14",
    "ETAPA OCURRENCIA": "C16",
    "ETAPA DETECCIÓN": "C18",
    "CLASIFICACIÓN": "C20",
    "IMPACTO": "C22",
    "OBSERVACIONES": "C24",
    "USUARIO": "C26",
    "MOTIVO": "C28"
  };
  // 1. D) Celda input y limpieza
  const INPUT_ID = "C2"; 
  const RANGO_LIMPIEZA = "C3:C28"; 

  // 1. E) Colores UX
  const UX = {
    EXITO_BG: "#D4EDDA",
    EXITO_TXT: "#155724",
    ERROR_BG: "#F8D7DA",
    ERROR_TXT: "#721C24"
  };

  // 1. F) FUNCIÓN HELPER ENCAPSULADA (Office scripts no tolera la importación desde otro script, esta función se podría modularizar e importar en otro caso)
  function reportarError(ws: ExcelScript.Worksheet, dir: string, texto: string, colors: typeof UX) {
    const rango = ws.getRange(dir);
    rango.setValue(texto);
    rango.getFormat().getFill().setColor(UX.ERROR_BG);
    rango.getFormat().getFont().setColor(UX.ERROR_TXT);
    rango.getFormat().setWrapText(true);
    rango.select();
  }
  // ---------------------------------------------------------------

  const wsInput = workbook.getWorksheet(SHEET_INPUT)!;
  const wsBD = workbook.getWorksheet(SHEET_BD)!;
  const wsMaestros = workbook.getWorksheet(SHEET_MAESTROS)!;

  // 2. LIMPIEZA VISUAL INICIAL
  // 2. A) Se limpia la pantalla antes de empezar
  try {
    const msj = wsInput.getRange(RANGO_MENSAJES);
    msj.clear(ExcelScript.ClearApplyTo.contents);
    msj.getFormat().getFill().clear();
    
    // 2. B) Limpieza de los campos de datos (pero NO el ID en C2)
    wsInput.getRange(RANGO_LIMPIEZA).clear(ExcelScript.ClearApplyTo.contents);
  } catch (e) {}

  // 3. VALIDACIÓN DEL INPUT
  const idBuscado = wsInput.getRange(INPUT_ID).getValue();

  if (!idBuscado) {
    reportarError(wsInput, RANGO_MENSAJES, "⚠️ Por favor ingresa un ID numérico en la celda C2.", UX);
    return;
  }

  // 4. BÚSQUEDA Y LECTURA
  let clave = "";
  
  // Try Global para asegurar el Finally
  try {
    // 4. A)) Lectura Segura de Clave
    if (wsMaestros) {
      clave = wsMaestros.getRange(CELL_CLAVE).getText();
    }

    // 4. B) Unprotect Seguro
    // Nota: Aunque para leer no es 100% obligatorio desproteger, se hace para evitar bloqueos de lectura en rangos específicos.
    wsBD.getProtection().unprotect(String(clave));
    wsInput.getProtection().unprotect(String(clave));

    // 4. C) ESTRATEGIA DE BÚSQUEDA
    const rangoUsado = wsBD.getUsedRange();
    if (!rangoUsado) throw new Error("La Base de Datos está vacía.");

    const valoresBD = rangoUsado.getValues(); 
    if (valoresBD.length < 2) throw new Error("No hay datos en la BD (solo encabezados).");

    // 4. C) i) Identificar headers
    const encabezados = valoresBD[0] as string[];
    const indexID = encabezados.indexOf("ID");

    if (indexID === -1) throw new Error("No se encuentra la columna 'ID' en BD_DESVIOS.");

    // 4. C) ii) Buscar la fila
    let filaEncontrada: (string | number | boolean)[] | null = null;

    let i = 1; // Empezamos en 1 para saltar headers
    while (i < valoresBD.length && !filaEncontrada) {
        if (valoresBD[i][indexID] == idBuscado) {
            filaEncontrada = valoresBD[i];
        }
        i++;
    }

    // 4. D) RESULTADO
    if (filaEncontrada) {
        // 4. D) i) Volcado de datos dinámico
        for (const [columnaBD, celdaInput] of Object.entries(MAPA_LECTURA)) {
            const idx = encabezados.indexOf(columnaBD);
            if (idx !== -1 && columnaBD !== "USUARIO") {
                const valor = filaEncontrada[idx];
                wsInput.getRange(celdaInput).setValue(valor);
            }
        }

        // 4. D) ii) Mensaje Éxito
        const msj = wsInput.getRange(RANGO_MENSAJES);
        msj.setValue(`✅ Desvío #${idBuscado} cargado correctamente.`);
        msj.getFormat().getFill().setColor(UX.EXITO_BG);
        msj.getFormat().getFont().setColor(UX.EXITO_TXT);
        msj.getFormat().getFont().setBold(true);
        msj.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        msj.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        msj.select();

    } else {
        // 4. D) iii) NO ENCONTRADO
        reportarError(wsInput, RANGO_MENSAJES, `⛔ No se encontró el ID "${idBuscado}" en la base de datos.`, UX);
    }

  } catch (error) {
    let errorTxt = (error as Error).message || String(error);
    reportarError(wsInput, RANGO_MENSAJES, "❌ ERROR BÚSQUEDA:\n" + errorTxt, UX);
  } finally {
    // 5. REPROTECT FINAL (Blindado)
    try { if (wsBD) wsBD.getProtection().protect(undefined, String(clave)); } catch (e) {}
    try { if (wsInput) wsInput.getProtection().protect(undefined, String(clave)); } catch (e) {}
  }
}