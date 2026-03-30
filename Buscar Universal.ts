/**
 * SCRIPT: ENTIDAD_BUSCAR_UNIVERSAL
 * OBJETIVO: Localizar un registro por ID y cargar sus datos en el formulario de entrada.
 * GARANTÍA: Asegura que el usuario visualice la información más reciente de la base de datos antes de cualquier edición.
 */
function main(
  workbook: ExcelScript.Workbook,
  entidadParam: string,           // Origen Power Automate: "DESVIO" o "CAPA"
  idDesdePanel: string,           // ID capturado en la interfaz de usuario
  userEmail: string               // Usuario que realiza la consulta
) {
  // --- 1. NORMALIZACIÓN DE PARÁMETROS ---
  const nombreEntidadNormalizado = String(entidadParam).replace(/[\[\]"]/g, "").toUpperCase().trim();

  // --- 2. DICCIONARIO DE CONFIGURACIÓN DE ENTIDADES (Metadata) ---
  const CONFIGURACION_ENTIDADES: { [key: string]: { tabla: string, hojaInput: string, etiqueta: string, articulo: string, genero: string } } = {
    "DESVIO": { tabla: "TablaDesvios", hojaInput: "INP_DES", etiqueta: "desvío", articulo: "el", genero: "o" },
    "CAPA": { tabla: "TablaCapas", hojaInput: "INP_CAPAS", etiqueta: "CAPA", articulo: "la", genero: "a" }
  };

  const configuracionActiva = CONFIGURACION_ENTIDADES[nombreEntidadNormalizado];
  if (!configuracionActiva) throw new Error(`La entidad '${entidadParam}' no está configurada.`);

  // Asignación de variables descriptivas
  const etiquetaEntidad = configuracionActiva.etiqueta;
  const articuloEntidad = configuracionActiva.articulo;
  const generoEntidad = configuracionActiva.genero;
  const nombreTablaPrincipal = configuracionActiva.tabla;
  const nombreHojaEntrada = configuracionActiva.hojaInput;

  // Definición de Tipos e Interfaces
  type ValorCelda = string | number | boolean;
  interface ResultadoAccion { success: boolean; message: string; logLevel: 'EXITO' | 'ERROR' | 'WARN' | 'INFO'; }
  interface MapaColoresUX { [key: string]: { fondo: string; texto: string } }

  const resultadoOperacion: ResultadoAccion = { success: true, message: "Inicio de búsqueda", logLevel: 'INFO' };
  const PALETA_COLORES_UX: MapaColoresUX = {
    EXITO: { fondo: "#D4EDDA", texto: "#155724" },
    ERROR: { fondo: "#F8D7DA", texto: "#721C24" },
    WARN: { fondo: "#FFF3CD", texto: "#856404" },
    INFO: { fondo: "#E2E3E5", texto: "#383D41" }
  };

  let hojaEntradaWS: ExcelScript.Worksheet | undefined,
    tablaBaseDatos: ExcelScript.Table | undefined,
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

    if (!tablaBaseDatos) throw new Error(`Infraestructura: La tabla '${nombreTablaPrincipal}' no existe.`);

    // Limpieza de filtros para asegurar que la búsqueda recorra todos los registros
    const autoFiltro = tablaBaseDatos.getAutoFilter();
    if (autoFiltro) autoFiltro.clearCriteria();

    // --- II. CAPTURA DEL IDENTIFICADOR A LOCALIZAR ---
    hojaEntradaWS.getProtection().unprotect(claveProteccion);
    const idABuscar = idDesdePanel.trim().toUpperCase();

    if (!idABuscar || idABuscar === "") {
      resultadoOperacion.success = false;
      resultadoOperacion.message = `Se requiere ingresar un ID de ${etiquetaEntidad} en el panel para buscar.`;
      resultadoOperacion.logLevel = 'WARN';
    } else {
      
      const rangoEtiquetasFormulario = hojaEntradaWS.getRange("B:B").getUsedRange();
      const encabezadosTabla = tablaBaseDatos.getHeaderRowRange().getValues()[0].map(h => String(h).toUpperCase().replace(/\s/g, "_"));
      const nombreCampoPrimario = encabezadosTabla[0];

      if (rangoEtiquetasFormulario) {
        const matrizEtiquetas = rangoEtiquetasFormulario.getValues() as ValorCelda[][];
        const indiceFilaInicialForm = rangoEtiquetasFormulario.getRowIndex();
        const mapaCoordenadasFormulario: { [key: string]: number } = {};

        // Mapeamos las posiciones (filas) de cada campo en el formulario de Excel
        matrizEtiquetas.forEach((fila, i) => {
          const etiquetaLimpia = String(fila[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
          if (etiquetaLimpia !== "") {
            mapaCoordenadasFormulario[etiquetaLimpia] = i + indiceFilaInicialForm;
          }
        });

        // --- III. PROCESO DE BÚSQUEDA EN BASE DE DATOS ---
        const matrizValoresDB = tablaBaseDatos.getRangeBetweenHeaderAndTotal().getValues();
        const matrizTextosDB = tablaBaseDatos.getRangeBetweenHeaderAndTotal().getTexts();
        let indiceFilaEncontrada = -1;
        let contadorFilas = 0;
        let registroEncontrado = false;

        // Búsqueda secuencial por coincidencia de ID
        while (contadorFilas < matrizValoresDB.length && !registroEncontrado) {
          if (String(matrizValoresDB[contadorFilas][encabezadosTabla.indexOf(nombreCampoPrimario)]) === idABuscar) {
            indiceFilaEncontrada = contadorFilas;
            registroEncontrado = true;
          }
          contadorFilas++;
        }

        if (registroEncontrado) {
          // --- 1. LIMPIEZA PREVENTIVA DEL FORMULARIO ---
          matrizEtiquetas.forEach((fila, i) => {
            const campoLimpio = String(fila[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
            if (campoLimpio !== "") {
              hojaEntradaWS!.getRangeByIndexes(i + indiceFilaInicialForm, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
            }
          });

          // --- 2. POBLACIÓN DEL FORMULARIO CON DATOS DE ORIGEN ---
          encabezadosTabla.forEach((nombreEncabezado, indiceColumna) => {
            if (mapaCoordenadasFormulario[nombreEncabezado] !== undefined) {
              const rangoDestino = hojaEntradaWS!.getRangeByIndexes(mapaCoordenadasFormulario[nombreEncabezado], 2, 1, 1);

              // Tratamiento especial para campos de fecha (preservar formato numérico de Excel)
              if (nombreEncabezado.includes("FECHA")) {
                rangoDestino.setValue(matrizValoresDB[indiceFilaEncontrada][indiceColumna]); 
                rangoDestino.setNumberFormatLocal("dd/mm/aaaa");
              } else {
                // Para el resto de campos, usamos el texto formateado de la tabla
                rangoDestino.setValue(matrizTextosDB[indiceFilaEncontrada][indiceColumna]);
              }
            }
          });

          // --- 3. REFUERZO DE IDENTIDAD (Seguridad Visual) ---
          // Aseguramos que el campo ID del formulario muestre exactamente el ID buscado
          if (mapaCoordenadasFormulario[nombreCampoPrimario] !== undefined) {
            hojaEntradaWS!.getRangeByIndexes(mapaCoordenadasFormulario[nombreCampoPrimario], 2, 1, 1).setValue(idABuscar);
          }

          resultadoOperacion.message = `✅ ${etiquetaEntidad.charAt(0).toUpperCase() + etiquetaEntidad.slice(1)} #${idABuscar} cargad${generoEntidad} con éxito.`;
          resultadoOperacion.logLevel = 'EXITO';

        } else {
          resultadoOperacion.success = false;
          resultadoOperacion.message = `${articuloEntidad.charAt(0).toUpperCase() + articuloEntidad.slice(1)} ${etiquetaEntidad} #${idABuscar} no existe en la base de datos.`;
          resultadoOperacion.logLevel = 'ERROR';
        }
      }
    }
  } catch (excepcionSistema) {
    resultadoOperacion.success = false;
    resultadoOperacion.message = `❌ Error de Sistema: ${String(excepcionSistema)}`;
    resultadoOperacion.logLevel = 'ERROR';
  } finally {
    // --- IX. PROTOCOLO DE CIERRE Y SEGURIDAD ---
    if (itemClaveSistema && hojaEntradaWS) {
      const passFinal = itemClaveSistema.getRange().getText();
      auxiliarActualizarInterfazUX(hojaEntradaWS, resultadoOperacion, PALETA_COLORES_UX, passFinal);
      auxiliarProtegerHoja(hojaEntradaWS, passFinal, resultadoOperacion);
    }
  }

  // --- FUNCIONES AUXILIARES (HELPERS) ---

  function auxiliarActualizarInterfazUX(hoja: ExcelScript.Worksheet, res: ResultadoAccion, colores: MapaColoresUX, pass: string) {
    const itemFeedback = hoja.getNamedItem("UI_FEEDBACK");
    if (!itemFeedback) return; 
    const rangoFeedback = itemFeedback.getRange();
    const estiloUX = colores[res.logLevel];
    const fechaHora = new Date();
    const marcaTiempo = fechaHora.toLocaleTimeString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
    const iconoLatido = (fechaHora.getSeconds() % 2 === 0) ? "⚡" : "✨";
    try {
      hoja.getProtection().unprotect(pass);
      rangoFeedback.setValue(`[${marcaTiempo}] ${iconoLatido} ${res.message}`);
      rangoFeedback.getFormat().getFill().setColor(estiloUX.fondo);
      rangoFeedback.getFormat().getFont().setColor(estiloUX.texto);
      rangoFeedback.getFormat().getFont().setBold(true);
      rangoFeedback.getFormat().setWrapText(true);
    } catch (e) {
      try { rangoFeedback.setValue(res.message); } catch (e2) {}
    }
  }

  function auxiliarProtegerHoja(hoja: ExcelScript.Worksheet | undefined, pass: string, res: ResultadoAccion) {
    if (hoja) {
      try { hoja.getProtection().protect({ allowAutoFilter: true }, pass); }
      catch (e) { res.message += ` [⚠️ Seguridad: ${hoja.getName()}]`; }
    }
  }
}