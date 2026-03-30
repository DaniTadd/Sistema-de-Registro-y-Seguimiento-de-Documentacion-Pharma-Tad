/**
 * SCRIPT: ENTIDAD_CAMBIAR_ESTADO_AUDITABLE
 * OBJETIVO: Gestionar el cierre y reapertura de registros (Toggle) con bloqueo de registros anulados.
 * GARANTÍA: Asegura que cada cambio de estado quede justificado con un motivo y registrado en el historial.
 */
function main(
  workbook: ExcelScript.Workbook,
  entidadParam: string,           // Origen Power Automate: "DESVIO" o "CAPA"
  idDesdePanel: string,           // ID capturado en el panel de control
  motivoDesdePanel: string,       // Justificación técnica del cambio de estado
  userEmail: string               // Usuario que autoriza la transición
) {
  // --- 1. NORMALIZACIÓN DE PARÁMETROS ---
  const nombreEntidadNormalizado = String(entidadParam).replace(/[\[\]"]/g, "").toUpperCase().trim();

  // --- 2. DICCIONARIO DE CONFIGURACIÓN DE ENTIDADES (Metadata) ---
  const CONFIGURACION_ENTIDADES: { [key: string]: { tabla: string, historial: string, hojaInput: string, etiqueta: string, articulo: string, genero: string } } = {
    "DESVIO": { tabla: "TablaDesvios", historial: "TablaDesviosHistorial", hojaInput: "INP_DES", etiqueta: "desvío", articulo: "el", genero: "o" },
    "CAPA": { tabla: "TablaCAPAs", historial: "TablaCAPAsHistorial", hojaInput: "INP_CAPA", etiqueta: "CAPA", articulo: "la", genero: "a" }
  };

  const configuracionActiva = CONFIGURACION_ENTIDADES[nombreEntidadNormalizado];
  if (!configuracionActiva) throw new Error(`La entidad '${entidadParam}' no está configurada.`);

  // Asignación de variables descriptivas para homogeneidad
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

  const resultadoOperacion: ResultadoAccion = { success: true, message: "Inicio de transición", logLevel: 'INFO' };
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

    claveProteccion = itemClaveSistema.getRange().getText();
    tablaBaseDatos = workbook.getTable(nombreTablaPrincipal);
    tablaHistorial = workbook.getTable(nombreTablaHistorial);

    if (!tablaBaseDatos || !tablaHistorial) throw new Error("Infraestructura: Tablas de datos o historial no encontradas.");

    // --- II. CAPTURA DE DATOS DEL PANEL ---
    hojaEntradaWS.getProtection().unprotect(claveProteccion);
    const rangoEtiquetasFormulario = hojaEntradaWS.getRange("B:B").getUsedRange();
    const encabezadosTabla = tablaBaseDatos.getHeaderRowRange().getValues()[0].map(h => String(h).toUpperCase().replace(/\s/g, "_"));
    const nombreCampoPrimario = encabezadosTabla[0];

    // Prioridad absoluta a los parámetros externos (Power Automate)
    const idABuscar = idDesdePanel.trim().toUpperCase();
    const motivoDeCambio = motivoDesdePanel.trim();

    // --- III. VALIDACIONES PREVIAS A LA TRANSICIÓN ---
    if (!idABuscar || idABuscar === "") {
      resultadoOperacion.success = false;
      resultadoOperacion.message = `Se requiere un ID de ${etiquetaEntidad} en el panel para realizar esta acción.`;
      resultadoOperacion.logLevel = 'WARN';
    } else if (!motivoDeCambio || motivoDeCambio === "") {
      resultadoOperacion.success = false;
      resultadoOperacion.message = `La justificación (motivo) es obligatoria para cambiar el estado de ${articuloEntidad} ${etiquetaEntidad}.`;
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

      if (!registroEncontrado) {
        resultadoOperacion.success = false;
        resultadoOperacion.message = `${etiquetaEntidad.charAt(0).toUpperCase() + etiquetaEntidad.slice(1)} #${idABuscar} no encontrad${generoEntidad}.`;
        resultadoOperacion.logLevel = 'ERROR';
      } else {
        const estadoActual = String(matrizValoresDB[indiceFilaEncontrada][encabezadosTabla.indexOf("ESTADO")]).toUpperCase();

        // 1. Regla de Inmutabilidad: Registros ANULADOS no pueden cambiar
        if (estadoActual === "ANULADO") {
          resultadoOperacion.success = false;
          resultadoOperacion.message = `Protocolo: Un registro ANULADO es definitivo y no permite cambios de estado.`;
          resultadoOperacion.logLevel = 'WARN';
        } else {
          
          // 2. Lógica de Transición (Toggle: ABIERTO <-> CERRADO)
          const nuevoEstado = (estadoActual === "ABIERTO") ? "CERRADO" : "ABIERTO";

          // --- V. EJECUCIÓN DEL CAMBIO (COMMIT) ---
          tablaBaseDatos.getWorksheet().getProtection().unprotect(claveProteccion);
          const idxEstado = encabezadosTabla.indexOf("ESTADO");
          const idxAuditTrail = encabezadosTabla.indexOf("AUDIT_TRAIL");
          const idxUsuario = encabezadosTabla.indexOf("USUARIO");

          // Actualización física en la tabla principal
          const rangoFilaDB = tablaBaseDatos.getRangeBetweenHeaderAndTotal();
          rangoFilaDB.getCell(indiceFilaEncontrada, idxEstado).setValue(nuevoEstado);
          
          if (idxUsuario !== -1) rangoFilaDB.getCell(indiceFilaEncontrada, idxUsuario).setValue(userEmail);
          if (idxAuditTrail !== -1) {
            rangoFilaDB.getCell(indiceFilaEncontrada, idxAuditTrail).setValue(
              new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false })
            );
          }

          // --- VI. REGISTRO EN EL HISTORIAL (AUDITORÍA DE TRANSICIÓN) ---
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
            if (headCaps === "MOTIVO") return motivoDeCambio;
            if (headCaps === "CAMBIOS") return `ESTADO: [${estadoActual}] -> [${nuevoEstado}]`;
            if (headCaps === "FECHA_CAMBIO") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
            return "";
          });
          tablaHistorial.addRow(-1, filaRegistroHistorial);

          resultadoOperacion.message = `✅ ${etiquetaEntidad.charAt(0).toUpperCase() + etiquetaEntidad.slice(1)} #${idABuscar} ${nuevoEstado === "ABIERTO" ? "reabiert" + generoEntidad : "cerrad" + generoEntidad} con éxito.`;
          resultadoOperacion.logLevel = 'EXITO';

          // --- VII. LIMPIEZA DEL FORMULARIO ---
          if (rangoEtiquetasFormulario) {
            const matrizEtiquetasForm = rangoEtiquetasFormulario.getValues() as ValorCelda[][];
            const offsetFilaForm = rangoEtiquetasFormulario.getRowIndex();
            auxiliarLimpiarFormulario(hojaEntradaWS, matrizEtiquetasForm, offsetFilaForm, nombreCampoPrimario);
          }
        }
      }
    }
  } catch (excepcionSistema) {
    resultadoOperacion.success = false;
    resultadoOperacion.message = `❌ Error Crítico: ${String(excepcionSistema)}`;
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
      // Mantenemos inalterada la lógica de limpieza (Usuario y ID no se limpian por persistencia visual)
      if (claveCampo !== "" && claveCampo !== campoId && claveCampo !== "USUARIO") {
        hoja.getRangeByIndexes(i + filaInicio, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
      }
    });
  }
}