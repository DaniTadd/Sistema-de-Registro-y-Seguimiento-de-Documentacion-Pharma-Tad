function main(workbook: ExcelScript.Workbook) {
  // --- BLOQUE 1: CONFIGURACIÓN Y SEGURIDAD INICIAL ---
  const hojaInput = workbook.getWorksheet("INPUT_DESVIOS");
  const hojaMaestros = workbook.getWorksheet("MAESTROS");
  const CLAVE_SEGURIDAD = hojaMaestros.getRange("XFD1").getText();
  // 1. DESPROTEGER INPUT (Necesario para limpiar el mensaje)
  try {
    hojaInput.getProtection().unprotect(CLAVE_SEGURIDAD);
  } catch (e) {
    console.log("Input ya estaba desprotegida o clave incorrecta.");
  }

  // 2. CONFIGURAR ÁREA DE MENSAJE (E4:H6)
  const celdaMensaje = hojaInput.getRange("E4:H6");

  // 3. COMBINAR CELDAS (Para que el texto no se repita)
  celdaMensaje.merge(false);

  // Limpieza visual
  celdaMensaje.clear(ExcelScript.ClearApplyTo.contents);
  celdaMensaje.getFormat().getFill().clear();
  celdaMensaje.getFormat().getFont().setColor("Black");
  celdaMensaje.getFormat().getFont().setBold(false);
  // Alineación para que se vea bonito centrado
  celdaMensaje.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  celdaMensaje.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  celdaMensaje.getFormat().setWrapText(true);

  // --- BLOQUE 2: LECTURA DE INPUTS ---
  const fechaSuceso = hojaInput.getRange("C4").getValue() as number;
  const fechaRegistroManual = hojaInput.getRange("C5").getValue() as number;
  const fechaRecepcionQA = hojaInput.getRange("C6").getValue() as number;
  const planta = hojaInput.getRange("C8").getText();
  const tercerista = hojaInput.getRange("C10").getText();
  const descripcion = hojaInput.getRange("C12").getText();
  const etapaOcurrencia = hojaInput.getRange("C14").getText();
  const etapaDeteccion = hojaInput.getRange("C16").getText();
  const clasificacion = hojaInput.getRange("C18").getText();
  const impacto = hojaInput.getRange("C20").getText();
  const observaciones = hojaInput.getRange("C22").getText();
  const usuarioResponsable = hojaInput.getRange("C24").getText();

  let esValido = true;
  let mensajeError = "";

  // --- BLOQUE 3: VALIDACIONES ---
  if (fechaRegistroManual < fechaSuceso) mensajeError += "• Fecha registro anterior al suceso.\n";
  if (fechaRecepcionQA < fechaRegistroManual) mensajeError += "• Fecha QA anterior al registro.\n";

  if (esValido && mensajeError.length > 0) esValido = false;

  // --- BLOQUE 4: EJECUCIÓN ---
  if (esValido) {
    try {
      const hojaBD = workbook.getWorksheet("BD_DESVIOS");
      if (!hojaBD) throw new Error("Falta hoja BD_DESVIOS.");

      // DESPROTEGER BD
      hojaBD.getProtection().unprotect(CLAVE_SEGURIDAD);

      const tablaDesvios = hojaBD.getTable("TablaDesvios");
      if (!tablaDesvios) throw new Error("Falta TablaDesvios.");

      // Cálculos ID
      let nuevoId = 1;
      let cantidadFilas = tablaDesvios.getRowCount();
      if (cantidadFilas > 0) {
        let valoresID = tablaDesvios.getColumnByName("ID").getRangeBetweenHeaderAndTotal().getValues();
        let listaNumeros = valoresID.map(fila => Number(fila[0]));
        nuevoId = Math.max(...listaNumeros) + 1;
      }
      let fechaHoraSistema = new Date().toLocaleString();

      // Diccionario
      const datosDiccionario: { [key: string]: string | number | boolean } = {
        "ID": nuevoId,
        "Fecha Registro": fechaRegistroManual,
        "Fecha Suceso": fechaSuceso,
        "Fecha QA": fechaRecepcionQA,
        "Planta": planta,
        "Tercerista": tercerista,
        "Descripción": descripcion,
        "Etapa Ocurrencia": etapaOcurrencia,
        "Etapa Detección": etapaDeteccion,
        "Clasificación": clasificacion,
        "Impacto": impacto,
        "Observaciones": observaciones,
        "Usuario": usuarioResponsable,
        "Audit Trail": fechaHoraSistema
      };

      // Escritura
      const rangoEncabezados = tablaDesvios.getHeaderRowRange();
      const encabezadosExcel = rangoEncabezados.getValues()[0] as string[];
      const nuevaFilaOrdenada = encabezadosExcel.map(columna => datosDiccionario[columna] ?? "");

      tablaDesvios.addRow(-1, nuevaFilaOrdenada);

      // Formato
      const rangoTabla = tablaDesvios.getRange();
      rangoTabla.getFormat().setWrapText(true);
      rangoTabla.getFormat().autofitColumns();
      rangoTabla.getFormat().autofitRows();

      // Limpieza Formulario (Hasta C24)
      hojaInput.getRange("C4:C24").clear(ExcelScript.ClearApplyTo.contents);

      // Éxito
      celdaMensaje.setValue(`✅ Desvío #${nuevoId} guardado.`);
      celdaMensaje.getFormat().getFill().setColor("#DFF6DD");
      celdaMensaje.getFormat().getFont().setColor("#006600");
      celdaMensaje.getFormat().getFont().setBold(true);

      // RE-PROTEGER BD
      hojaBD.getProtection().protect({
        allowInsertRows: false,
        allowDeleteRows: false,
        allowFormatCells: false,
        allowAutoFilter: true,
        allowSort: true
      }, CLAVE_SEGURIDAD);

    } catch (error) {
      try {
        workbook.getWorksheet("BD_DESVIOS").getProtection().protect({ allowAutoFilter: true }, CLAVE_SEGURIDAD);
      } catch (e) { }

      celdaMensaje.setValue("⛔ ERROR:\n" + error.message);
      celdaMensaje.getFormat().getFill().setColor("#FFDDDD");
    }

  } else {
    console.log(mensajeError);
    celdaMensaje.setValue("⚠️ DATOS INVÁLIDOS:\n" + mensajeError);
    celdaMensaje.getFormat().getFill().setColor("#FFFFCC");
    celdaMensaje.getFormat().getFont().setColor("#996600");
  }

  // --- CIERRE FINAL: RE-PROTEGER INPUT ---
  try {
    // Solo usamos las propiedades estándar permitidas en Office Scripts
    hojaInput.getProtection().protect({
      allowFormatCells: false,
      allowFormatColumns: false,
      allowFormatRows: false,
      allowInsertRows: false,
      allowDeleteRows: false,
      allowSort: false,
      allowAutoFilter: false
    }, CLAVE_SEGURIDAD);
  } catch (e) {
    console.log("No se pudo reproteger Input.");
  }
}