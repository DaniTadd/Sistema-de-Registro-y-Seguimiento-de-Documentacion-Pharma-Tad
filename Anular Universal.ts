/**
 * SCRIPT: ENTIDAD_ANULAR_UNIVERSAL
 * OBJETIVO: Realizar la anulación definitiva de un registro (Desvío/CAPA) con justificación obligatoria.
 * GARANTÍA: Establece el estado inmutable "ANULADO" y genera el registro correspondiente en el historial de auditoría.
 */
function main(
  workbook: ExcelScript.Workbook,
  entidadParam: string,           // Origen Power Automate: "DESVIO" o "CAPA"
  idDesdePanel: string,           // ID capturado en el panel de control
  motivoDesdePanel: string,       // Justificación obligatoria de la anulación
  userEmail: string               // Usuario que ejecuta la anulación
) {
  // --- 1. NORMALIZACIÓN DE PARÁMETROS ---
  const nombreEntidadNormalizado = String(entidadParam).replace(/[\[\]"]/g, "").toUpperCase().trim();

  // --- 2. DICCIONARIO DE CONFIGURACIÓN DE ENTIDADES (Metadata) ---
  const CONFIGURACION_ENTIDADES: { [key: string]: { tabla: string, historial: string, hojaInput: string, etiqueta: string, articulo: string, genero: string } } = {
    "DESVIO": { tabla: "TablaDesvios", historial: "TablaDesviosHistorial", hojaInput: "INP_DES", etiqueta: "desvío", articulo: "el", genero: "o" },
    "CAPA": { tabla: "TablaCAPAs", historial: "TablaCAPAsHistorial", hojaInput: "INP_CAPA", etiqueta: "CAPA", articulo: "la", genero: "a" }
  };

  const configuracionActiva = CONFIGURACION_ENTIDADES[nombreEntidadNormalizado];
  if (!configuracionActiva) throw new Error(`La entidad '${entidadParam}' no está configurada en el sistema.`);

  // Asignación de variables descriptivas para homogeneidad del proyecto
  const etiquetaEntidad = configuracionActiva.etiqueta;
  const articuloEntidad = configuracionActiva.articulo;
  const generoEntidad = configuracionActiva.genero;
  const nombreTablaPrincipal = configuracionActiva.tabla;
  const nombreTablaHistorial = configuracionActiva.historial;
  const nombreHojaEntrada = configuracionActiva.hojaInput;

  // Definición de Tipos e Interfaces para el manejo de resultados
  type ValorCelda = string | number | boolean;
  interface ResultadoAccion { success: boolean; message: string; logLevel: 'EXITO' | 'ERROR' | 'WARN' | 'INFO'; }
  interface MapaColoresUX { [key: string]: { fondo: string; texto: string } }

  const resultadoOperacion: ResultadoAccion = { success: true, message: "Inicio de proceso de anulación", logLevel: 'INFO' };
  const PALETA_COLORES_UX: MapaColoresUX = {
    EXITO: { fondo: "#D4EDDA", texto: "#155724" },
    ERROR: { fondo: "#F8D7DA", texto: "#721C24" },
    WARN: { fondo: "#FFF3CD", texto: "#856404" },
    INFO: { fondo: "#E2E3E5", texto: "#383D41" }
  };

  let hojaEntradaWS: ExcelScript.Worksheet | undefined,
    tablaBaseDatos: ExcelScript.Table | undefined,
    tablaHistorial: ExcelScript.Table | undefined,
    itemClaveSistema: ExcelScript.NamedItem | undefined;
  let claveProteccion: string = "";

  try {
    // --- I. VALIDACIÓN DE INFRAESTRUCTURA ---
    hojaEntradaWS = workbook.getWorksheet(nombreHojaEntrada);
    itemClaveSistema = workbook.getNamedItem("SISTEMA_CLAVE");

    if (!hojaEntradaWS) throw new Error(`Infraestructura: No se halló la hoja '${nombreHojaEntrada}'.`);
    if (!itemClaveSistema) throw new Error("Infraestructura: No se halló 'SISTEMA_CLAVE'.");
    
    // Posicionamiento preventivo del cursor para evitar errores de foco
    hojaEntradaWS.getRange("A1").select();
    
    claveProteccion = itemClaveSistema.getRange().getText();
    tablaBaseDatos = workbook.getTable(nombreTablaPrincipal);
    tablaHistorial = workbook.getTable(nombreTablaHistorial);

    if (!tablaBaseDatos || !tablaHistorial) throw new Error("Infraestructura: Tablas de datos o historial no encontradas.");

    // --- II. CAPTURA DE DATOS PARA EL PROCESO ---
    hojaEntradaWS.getProtection().unprotect(claveProteccion);
    const rangoEtiquetasFormulario = hojaEntradaWS.getRange("B:B").getUsedRange();
    const encabezadosTabla = tablaBaseDatos.getHeaderRowRange().getValues()[0].map(h => String(h).toUpperCase().replace(/\s/g, "_"));
    const nombreCampoPrimario = encabezadosTabla[0];

    // Prioridad absoluta a los parámetros validados por Power Automate
    const idABuscar = idDesdePanel.trim().toUpperCase();
    const motivoDeAnulacion = motivoDesdePanel.trim();

    // --- III. VALIDACIONES PREVIAS (SEGURIDAD DE OPERACIÓN) ---
    if (!idABuscar || idABuscar === "") {
      resultadoOperacion.success = false;
      resultadoOperacion.message = `Se requiere un ID de ${etiquetaEntidad} en el panel para proceder con la anulación.`;
      resultadoOperacion.logLevel = 'WARN';
    } else if (!motivoDeAnulacion || motivoDeAnulacion === "") {
      resultadoOperacion.success = false;
      resultadoOperacion.message = `Protocolo: No se puede anular un registro sin una justificación obligatoria.`;
      resultadoOperacion.logLevel = 'WARN';
    } else {
      
      // --- IV. LOCALIZACIÓN DEL REGISTRO EN BASE DE DATOS ---
      const matrizValoresDB = tablaBaseDatos.getRangeBetweenHeaderAndTotal().getValues();
      let indiceFilaEncontrada = -1;
      let contadorFilas = 0;
      let registroEncontrado = false;

      while (contadorFilas < matrizValoresDB.length && !registroEncontrado) {
        if (String(matrizValoresDB[contadorFilas][encabezadosTabla.indexOf(nombreCampoPrimario)]) === idABuscar) {
          indiceFilaEncontrada = contadorFilas;
          registroEncontrado = true;
        }
        contadorFilas++;
      }

      // Validación de existencia y estado actual
      if (!registroEncontrado) {
        resultadoOperacion.success = false;
        resultadoOperacion.message = `${etiquetaEntidad.charAt(0).toUpperCase() + etiquetaEntidad.slice(1)} #${idABuscar} no encontrad${generoEntidad} en la base de datos.`;
        resultadoOperacion.logLevel = 'ERROR';
      } else if (String(matrizValoresDB[indiceFilaEncontrada][encabezadosTabla.indexOf("ESTADO")]).toUpperCase() === "ANULADO") {
        resultadoOperacion.success = false;
        resultadoOperacion.message = `${articuloEntidad.charAt(0).toUpperCase() + articuloEntidad.slice(1)} ${etiquetaEntidad} #${idABuscar} ya se encuentra anulad${generoEntidad}.`;
        resultadoOperacion.logLevel = 'INFO';
      } else {
        
        // --- V. EJECUCIÓN DE LA ANULACIÓN (COMMIT DEFINITIVO) ---
        
        // 1. Actualización de la Tabla Principal
        tablaBaseDatos.getWorksheet().getProtection().unprotect(claveProteccion);
        const idxEstado = encabezadosTabla.indexOf("ESTADO");
        const idxAuditTrail = encabezadosTabla.indexOf("AUDIT_TRAIL");
        const idxUsuario = encabezadosTabla.indexOf("USUARIO");

        const rangoDatosDB = tablaBaseDatos.getRangeBetweenHeaderAndTotal();
        if (idxEstado !== -1) rangoDatosDB.getCell(indiceFilaEncontrada, idxEstado).setValue("ANULADO");
        if (idxUsuario !== -1) rangoDatosDB.getCell(indiceFilaEncontrada, idxUsuario).setValue(userEmail);
        if (idxAuditTrail !== -1) {
          rangoDatosDB.getCell(indiceFilaEncontrada, idxAuditTrail).setValue(
            new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false })
          );
        }

        // 2. Registro del Evento en la Tabla de Historial
        tablaHistorial.getWorksheet().getProtection().unprotect(claveProteccion);
        const filaRegistroHistorial = (tablaHistorial.getHeaderRowRange().getValues()[0] as string[]).map(h => {
          const headCaps = h.toUpperCase();
          if (headCaps === "ID_EVENTO") {
            const columnaIdEvento = tablaHistorial!.getColumnByName("ID_EVENTO").getRangeBetweenHeaderAndTotal();
            const valoresExistentes = columnaIdEvento ? columnaIdEvento.getValues() as number[][] : [];
            return tablaHistorial!.getRowCount() === 0 ? 1 : Math.max(...valoresExistentes.map(x => Number(x[0]))) + 1;
          }
          if (headCaps === nombreCampoPrimario) return idABuscar;
          if (headCaps === "USUARIO") return userEmail;
          if (headCaps === "MOTIVO") return motivoDeAnulacion;
          if (headCaps === "CAMBIOS") return "[REGISTRO ANULADO]";
          if (headCaps === "FECHA_CAMBIO") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
          return "";
        });
        tablaHistorial.addRow(-1, filaRegistroHistorial);

        resultadoOperacion.message = `✅ ${etiquetaEntidad.charAt(0).toUpperCase() + etiquetaEntidad.slice(1)} #${idABuscar} ANULAD${generoEntidad.toUpperCase()} correctamente.`;
        resultadoOperacion.logLevel = 'EXITO';

        // --- VI. LIMPIEZA DEL FORMULARIO DE INTERFAZ ---
        if (rangoEtiquetasFormulario) {
          const matrizEtiquetasForm = rangoEtiquetasFormulario.getValues() as ValorCelda[][];
          const offsetFilaForm = rangoEtiquetasFormulario.getRowIndex();
          auxiliarLimpiarFormulario(hojaEntradaWS, matrizEtiquetasForm, offsetFilaForm, nombreCampoPrimario);
        }
      }
    }
  } catch (excepcionSistema) {
    resultadoOperacion.success = false;
    resultadoOperacion.message = `❌ Error Crítico del Sistema: ${String(excepcionSistema)}`;
    resultadoOperacion.logLevel = 'ERROR';
  } finally {
    // --- IX. PROTOCOLO DE CIERRE Y SEGURIDAD ---
    if (itemClaveSistema && hojaEntradaWS) {
      const passFinal = itemClaveSistema.getRange().getText();
      auxiliarActualizarInterfazUX(hojaEntradaWS, resultadoOperacion, PALETA_COLORES_UX, passFinal);
      auxiliarProtegerHoja(hojaEntradaWS, passFinal, resultadoOperacion);
      if (tablaBaseDatos) auxiliarProtegerHoja(tablaBaseDatos.getWorksheet(), passFinal, resultadoOperacion);
      if (tablaHistorial) auxiliarProtegerHoja(tablaHistorial.getWorksheet(), passFinal, resultadoOperacion);
    }
  }

  // --- FUNCIONES AUXILIARES (HELPERS) ---

  function auxiliarActualizarInterfazUX(hoja: ExcelScript.Worksheet, res: ResultadoAccion, colores: MapaColoresUX, pass: string) {
    const itemUI = hoja.getNamedItem("UI_FEEDBACK");
    if (!itemUI) return; 
    const rangoUI = itemUI.getRange();
    const estiloUX = colores[res.logLevel];
    const fechaHora = new Date();
    const marcaTiempo = fechaHora.toLocaleTimeString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
    const iconoLatido = (fechaHora.getSeconds() % 2 === 0) ? "⚡" : "✨";
    try {
      hoja.getProtection().unprotect(pass);
      rangoUI.setValue(`[${marcaTiempo}] ${iconoLatido} ${res.message}`);
      rangoUI.getFormat().getFill().setColor(estiloUX.fondo);
      rangoUI.getFormat().getFont().setColor(estiloUX.texto);
      rangoUI.getFormat().getFont().setBold(true);
      rangoUI.getFormat().setWrapText(true);
    } catch (e) {
      try { rangoUI.setValue(res.message); } catch (e2) {}
    }
  }

  function auxiliarProtegerHoja(hoja: ExcelScript.Worksheet | undefined, pass: string, res: ResultadoAccion) {
    if (hoja) {
      try { hoja.getProtection().protect({ allowAutoFilter: true }, pass); }
      catch (e) { res.message += ` [⚠️ Seguridad: ${hoja.getName()}]`; }
    }
  }

  function auxiliarLimpiarFormulario(hoja: ExcelScript.Worksheet, matriz: ValorCelda[][], filaInicio: number, campoId: string) {
    matriz.forEach((fila, i) => {
      const claveCampo = String(fila[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
      // Mantenemos la lógica de inmutabilidad visual para Usuario e ID
      if (claveCampo !== "" && claveCampo !== campoId && claveCampo !== "USUARIO") {
        hoja.getRangeByIndexes(i + filaInicio, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
      }
    });
  }
}