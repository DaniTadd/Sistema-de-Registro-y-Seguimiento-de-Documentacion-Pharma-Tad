/**
 * SCRIPT: ENTIDAD_ACTUALIZAR_CON_AUDITORIA
 * OBJETIVO: Gestionar la actualización de registros (Desvíos/CAPAS) con validación de integridad y log de cambios.
 * GARANTÍA: Mantiene trazabilidad total en Tablas de Historial y campos de Audit Trail.
 */
function main(
  workbook: ExcelScript.Workbook,
  entidadParam: string,           // Origen Power Automate: "DESVIO" o "CAPA"
  idDesdePanel: string,           // ID proporcionado en la interfaz de usuario de PA
  motivoDesdePanel: string,       // Justificación técnica del cambio
  userEmail: string               // Usuario responsable de la transacción
) {
  // --- 1. NORMALIZACIÓN DE PARÁMETROS ---
  const nombreEntidadNormalizado = String(entidadParam).replace(/[\[\]"]/g, "").toUpperCase().trim();

  // --- 2. DICCIONARIO DE CONFIGURACIÓN DE ENTIDADES (Metadata) ---
  const CONFIGURACION_ENTIDADES: { [key: string]: { tabla: string, historial: string, hojaInput: string, etiqueta: string, articulo: string, genero: string } } = {
    "DESVIO": { tabla: "TablaDesvios", historial: "TablaDesviosHistorial", hojaInput: "INP_DES", etiqueta: "desvío", articulo: "el", genero: "o" },
    "CAPA": { tabla: "TablaCapas", historial: "TablaCapasHistorial", hojaInput: "INP_CAPAS", etiqueta: "CAPA", articulo: "la", genero: "a" }
  };

  const configuracionActiva = CONFIGURACION_ENTIDADES[nombreEntidadNormalizado];
  if (!configuracionActiva) throw new Error(`La entidad '${entidadParam}' no tiene una configuración definida.`);

  // Asignación de variables descriptivas para la lógica dinámica
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

  const resultadoOperacion: ResultadoAccion = { success: true, message: "Inicio de proceso", logLevel: 'INFO' };
  const PALETA_COLORES_UX: MapaColoresUX = {
    EXITO: { fondo: "#D4EDDA", texto: "#155724" },
    ERROR: { fondo: "#F8D7DA", texto: "#721C24" },
    WARN: { fondo: "#FFF3CD", texto: "#856404" },
    INFO: { fondo: "#E2E3E5", texto: "#383D41" }
  };

  let hojaEntradaWS: ExcelScript.Worksheet | undefined,
    tablaBaseDatos: ExcelScript.Table | undefined,
    tablaHistorial: ExcelScript.Table | undefined,
    hojaMaestrosWS: ExcelScript.Worksheet | undefined,
    itemClaveSistema: ExcelScript.NamedItem | undefined;
  let claveProteccion: string = "";

  try {
    // --- I. VALIDACIÓN DE INFRAESTRUCTURA Y SEGURIDAD ---
    hojaEntradaWS = workbook.getWorksheet(nombreHojaEntrada);
    hojaMaestrosWS = workbook.getWorksheet("MAESTROS");
    itemClaveSistema = workbook.getNamedItem("SISTEMA_CLAVE");

    if (!hojaEntradaWS) throw new Error(`Infraestructura: No se halló la hoja de entrada '${nombreHojaEntrada}'.`);
    if (!hojaMaestrosWS) throw new Error("Infraestructura: Hoja MAESTROS no disponible.");
    if (!itemClaveSistema) throw new Error("Infraestructura: No se halló el ítem 'SISTEMA_CLAVE'.");

    claveProteccion = itemClaveSistema.getRange().getText();
    tablaBaseDatos = workbook.getTable(nombreTablaPrincipal);
    tablaHistorial = workbook.getTable(nombreTablaHistorial);

    if (!tablaBaseDatos) throw new Error(`Infraestructura: La tabla '${nombreTablaPrincipal}' no existe.`);
    if (!tablaHistorial) throw new Error(`Infraestructura: La tabla '${nombreTablaHistorial}' no existe.`);

    // --- II. CAPTURA Y PROCESAMIENTO DEL FORMULARIO (COLUMNA B y C) ---
    hojaEntradaWS.getProtection().unprotect(claveProteccion);
    const rangoEtiquetas = hojaEntradaWS.getRange("B:B").getUsedRange();

    if (rangoEtiquetas) {
      const matrizEtiquetas = rangoEtiquetas.getValues() as ValorCelda[][];
      const indiceFilaInicial = rangoEtiquetas.getRowIndex();
      const objetoDatosFormulario: { [key: string]: ValorCelda } = {};
      const listaCamposObligatorios: string[] = [];

      matrizEtiquetas.forEach((fila, i) => {
        const etiquetaLimpia = String(fila[0]).trim().toUpperCase();
        if (etiquetaLimpia !== "") {
          const claveCampo = etiquetaLimpia.replace("*", "").trim().replace(/\s/g, "_");
          if (etiquetaLimpia.endsWith("*")) listaCamposObligatorios.push(claveCampo);
          
          const valorIngresado = hojaEntradaWS!.getRangeByIndexes(i + indiceFilaInicial, 2, 1, 1).getValue();
          objetoDatosFormulario[claveCampo] = (valorIngresado === null || String(valorIngresado).trim() === "") 
            ? (etiquetaLimpia.endsWith("*") ? "" : "N/A") 
            : valorIngresado;
        }
      });

      const encabezadosTabla = tablaBaseDatos.getHeaderRowRange().getValues()[0].map(h => String(h).toUpperCase().replace(/\s/g, "_"));
      const nombreCampoPrimario = encabezadosTabla[0]; // ID_DESVIO o ID_CAPA
      
      // --- III. SALVAGUARDA DE INTEGRIDAD ESTRICTA (Mismatch Check) ---
      const idEnPanelParam = idDesdePanel.trim().toUpperCase();
      const idEnHojaInput = String(objetoDatosFormulario[nombreCampoPrimario] || "").trim().toUpperCase();

      if (idEnPanelParam === "" || idEnHojaInput === "" || idEnHojaInput === "N/A" || idEnPanelParam !== idEnHojaInput) {
        resultadoOperacion.success = false;
        resultadoOperacion.message = `⚠️ ERROR: Mismatch de ID. Panel: [${idEnPanelParam || "VACÍO"}] | Hoja: [${idEnHojaInput || "VACÍO"}]. Sincronice el registro antes de actualizar.`;
        resultadoOperacion.logLevel = 'ERROR';
        throw new Error(resultadoOperacion.message);
      }

      const idABuscar = idEnPanelParam;

      if (!idABuscar || idABuscar === "" || idABuscar === "N/A") {
        resultadoOperacion.success = false;
        resultadoOperacion.message = `Se requiere un ID de ${etiquetaEntidad} válido.`;
        resultadoOperacion.logLevel = 'WARN';
      } else {
        // --- IV. LOCALIZACIÓN DEL REGISTRO Y VALIDACIÓN DE ESTADO ---
        const matrizValoresDB = tablaBaseDatos.getRangeBetweenHeaderAndTotal().getValues();
        const matrizTextosDB = tablaBaseDatos.getRangeBetweenHeaderAndTotal().getTexts();
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
          const valorEstadoActual = String(matrizValoresDB[indiceFilaEncontrada][encabezadosTabla.indexOf("ESTADO")]).toUpperCase();

          if (["CERRADO", "ANULADO"].includes(valorEstadoActual)) {
            resultadoOperacion.success = false;
            resultadoOperacion.message = `⚠️ No se permite modificar ${articuloEntidad} ${etiquetaEntidad} en estado ${valorEstadoActual}.`;
            resultadoOperacion.logLevel = 'WARN';
          } else {
            // --- V. DETECCIÓN DE CAMBIOS Y VALIDACIÓN DE REGLAS ---
            const valoresFilaOriginal = [...matrizValoresDB[indiceFilaEncontrada]];
            const logDeCambios: string[] = [];
            const listaErroresValidacion: string[] = [];
            const cambiosPendientes: { columna: number, valor: ValorCelda }[] = [];

            // Validación de Campos Obligatorios
            listaCamposObligatorios.forEach(campo => {
              if (objetoDatosFormulario[campo] === "" || objetoDatosFormulario[campo] === null) {
                listaErroresValidacion.push(`Falta campo obligatorio: ${campo.replace(/_/g, " ")}`);
              }
            });

            // Comparación de cada columna para identificar modificaciones
            encabezadosTabla.forEach((header, colIdx) => {
              if (["AUDIT_TRAIL", "ESTADO", "USUARIO"].includes(header)) return;
              if (objetoDatosFormulario.hasOwnProperty(header)) {
                const textoOriginal = matrizTextosDB[indiceFilaEncontrada][colIdx]; 
                const valorNuevo = objetoDatosFormulario[header];         
                let textoNuevoFormateado = "";

                // Manejo de formatos de fecha para comparación textual
                if (typeof valorNuevo === "number" && header.includes("FECHA")) {
                  const milisegundos = Math.round((valorNuevo - 25569) * 86400 * 1000) + 43200000;
                  textoNuevoFormateado = new Date(milisegundos).toLocaleDateString('es-AR', {
                    timeZone: 'America/Argentina/Buenos_Aires',
                    day: '2-digit', month: '2-digit', year: 'numeric'
                  });
                } else {
                  textoNuevoFormateado = String(valorNuevo);
                }

                if (textoOriginal.trim() !== textoNuevoFormateado.trim()) {
                  logDeCambios.push(`${header}: [${textoOriginal}] -> [${textoNuevoFormateado}]`);
                  cambiosPendientes.push({ columna: colIdx, valor: valorNuevo });
                }
              }
            });

            // --- VI. APLICACIÓN DE REGLAS DE NEGOCIO (TABLA MAESTRA) ---
            const tablaReglasMaestra = hojaMaestrosWS.getTable("TablaReglas");
            if (tablaReglasMaestra) {
              tablaReglasMaestra.getRangeBetweenHeaderAndTotal().getValues().forEach(regla => {
                if (nombreTablaPrincipal.toUpperCase().includes(String(regla[0]).toUpperCase())) {
                  const claveCampoA = String(regla[1]).toUpperCase().replace(/\s/g, "_");
                  const operador = String(regla[2]);
                  const referenciaRaw = String(regla[3]);
                  const mensajeErrorRegla = String(regla[4]);
                  const valorAValidar = objetoDatosFormulario[claveCampoA];

                  if (operador === "EXISTE_EN") {
                    if (valorAValidar && valorAValidar !== "N/A" && referenciaRaw.includes("[")) {
                      const [nombreTablaRef, parteColumna] = referenciaRaw.split("[");
                      const nombreColumnaRef = parteColumna.replace("]", "");
                      const tablaReferencia = workbook.getTable(nombreTablaRef);
                      if (tablaReferencia) {
                        const valoresMaestros = tablaReferencia.getColumnByName(nombreColumnaRef).getRangeBetweenHeaderAndTotal().getValues();
                        const existeValor = valoresMaestros.some(filaMaestra => String(filaMaestra[0]) === String(valorAValidar));
                        if (!existeValor) listaErroresValidacion.push(mensajeErrorRegla);
                      }
                    }
                  } 
                  else if (operador === "<=" || operador === ">=") {
                    const claveCampoB = referenciaRaw.toUpperCase().replace(/\s/g, "_");
                    const valorB = objetoDatosFormulario[claveCampoB] !== undefined 
                      ? objetoDatosFormulario[claveCampoB] 
                      : valoresFilaOriginal[encabezadosTabla.indexOf(claveCampoB)];

                    if (valorAValidar && valorB && valorAValidar !== "N/A" && valorB !== "N/A") {
                      const fechaA = auxiliarParsearFechaANumero(valorAValidar);
                      const fechaB = auxiliarParsearFechaANumero(valorB);
                      if (isNaN(fechaA) || isNaN(fechaB)) {
                        listaErroresValidacion.push("Error: Formato de fecha inválido.");
                      } else if (operador === "<=" && !(fechaA <= fechaB)) {
                        listaErroresValidacion.push(mensajeErrorRegla);
                      } else if (operador === ">=" && !(fechaA >= fechaB)) {
                        listaErroresValidacion.push(mensajeErrorRegla);
                      }
                    }
                  }
                }
              });
            }

            // --- VII. DECISIÓN DE PERSISTENCIA (COMMIT) ---
            if (listaErroresValidacion.length > 0) {
              resultadoOperacion.success = false;
              resultadoOperacion.message = "⚠️ " + listaErroresValidacion.join(" | ");
              resultadoOperacion.logLevel = 'WARN';
            } else if (logDeCambios.length === 0) {
              resultadoOperacion.message = "ℹ️ No se detectaron cambios para guardar.";
              resultadoOperacion.logLevel = 'INFO';
            } else if (!motivoDesdePanel || motivoDesdePanel.trim() === "") {
              resultadoOperacion.success = false;
              resultadoOperacion.message = `⚠️ Se requiere el motivo de modificación en el panel de control.`;
              resultadoOperacion.logLevel = 'WARN';
            } else {
              // --- VIII. EJECUCIÓN DE LA ACTUALIZACIÓN Y AUDITORÍA ---
              tablaBaseDatos.getWorksheet().getProtection().unprotect(claveProteccion);
              const rangoFilaAfectada = tablaBaseDatos.getRangeBetweenHeaderAndTotal().getRow(indiceFilaEncontrada);
              
              // Aplicar cambios a la tabla principal
              cambiosABase.forEach(cambio => rangoFilaAfectada.getCell(0, cambio.columna).setValue(cambio.valor));
              
              // Actualizar Audit Trail
              const idxAuditTrail = encabezadosTabla.indexOf("AUDIT_TRAIL");
              if (idxAuditTrail !== -1) rangoFilaAfectada.getCell(0, idxAuditTrail).setValue(new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false }));
              
              const idxUsuario = encabezadosTabla.indexOf("USUARIO");
              if (idxUsuario !== -1) rangoFilaAfectada.getCell(0, idxUsuario).setValue(userEmail);

              // Registro en Tabla Historial
              tablaHistorial.getWorksheet().getProtection().unprotect(claveProteccion);
              const nuevaFilaHistorial = (tablaHistorial.getHeaderRowRange().getValues()[0] as string[]).map(h => {
                const headCaps = h.toUpperCase();
                if (headCaps === "ID_EVENTO") {
                  const columnaIdEvento = tablaHistorial!.getColumnByName("ID_EVENTO").getRangeBetweenHeaderAndTotal();
                  const valoresExistentes = columnaIdEvento ? columnaIdEvento.getValues() as number[][] : [];
                  return tablaHistorial!.getRowCount() === 0 ? 1 : Math.max(...valoresExistentes.map(x => Number(x[0]))) + 1;
                }
                if (headCaps === nombreCampoPrimario) return idABuscar;
                if (headCaps === "USUARIO") return userEmail;
                if (headCaps === "MOTIVO") return motivoDesdePanel.trim();
                if (headCaps === "CAMBIOS") return logDeCambios.join(" | ");
                if (headCaps === "FECHA_CAMBIO") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
                return "";
              });
              tablaHistorial.addRow(-1, nuevaFilaHistorial);

              resultadoOperacion.message = `✅ ${idABuscar} actualizad${generoEntidad} con éxito.`;
              resultadoOperacion.logLevel = 'EXITO';
              auxiliarLimpiarFormulario(hojaEntradaWS, matrizEtiquetas, indiceFilaInicial, nombreCampoPrimario);
            }
          }
        }
      }
    }
  } catch (excepcion) {
    resultadoOperacion.success = false;
    resultadoOperacion.message = `❌ Error de Sistema: ${String(excepcion)}`;
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

  function auxiliarParsearFechaANumero(valor: ValorCelda): number {
    if (typeof valor === "number") return valor;
    const partes = String(valor).split("/");
    if (partes.length === 3) {
      const fechaObjeto = new Date(parseInt(partes[2]), parseInt(partes[1]) - 1, parseInt(partes[0]));
      return (fechaObjeto.getFullYear() === parseInt(partes[2])) ? fechaObjeto.getTime() : NaN;
    }
    return NaN;
  }

  function auxiliarActualizarInterfazUX(hoja: ExcelScript.Worksheet, res: ResultadoAccion, colores: MapaColoresUX, pass: string) {
    const itemUI = hoja.getNamedItem("UI_FEEDBACK");
    if (!itemUI) return; 
    const rangoUI = itemUI.getRange();
    const estilo = colores[res.logLevel];
    const fechaHora = new Date();
    const marcaTiempo = fechaHora.toLocaleTimeString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
    const iconoLatido = (fechaHora.getSeconds() % 2 === 0) ? "⚡" : "✨";
    try {
      hoja.getProtection().unprotect(pass);
      rangoUI.setValue(`[${marcaTiempo}] ${iconoLatido} ${res.message}`);
      rangoUI.getFormat().getFill().setColor(estilo.fondo);
      rangoUI.getFormat().getFont().setColor(estilo.texto);
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
      if (claveCampo !== "" && claveCampo !== campoId && claveCampo !== "MOTIVO") {
        hoja.getRangeByIndexes(i + filaInicio, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
      }
    });
  }
}