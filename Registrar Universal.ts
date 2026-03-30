/**
 * SCRIPT: ENTIDAD_REGISTRAR_UNIVERSAL
 * OBJETIVO: Crear nuevos registros (Desvíos/CAPAS/Afectaciones) con generación de ID automático y validación inicial.
 * GARANTÍA: Asegura que todo nuevo ingreso cumpla con las reglas de integridad referencial y cronológica.
 */
function main(
  workbook: ExcelScript.Workbook,
  entidadParam: string,           // Origen Power Automate: "DESVIO", "CAPA" o "AFECTACION"
  userEmail: string               // Usuario que realiza el registro inicial
) {
  // --- 1. NORMALIZACIÓN DE PARÁMETROS ---
  const nombreEntidadNormalizado = String(entidadParam).replace(/[\[\]"]/g, "").toUpperCase().trim();

  // --- 2. DICCIONARIO DE CONFIGURACIÓN DE ENTIDADES (Metadata) ---
  const CONFIGURACION_ENTIDADES: { [key: string]: { tabla: string, historial: string, hojaInput: string, prefijo: string, etiqueta: string, articulo: string, genero: string } } = {
    "DESVIO": { tabla: "TablaDesvios", historial: "TablaDesviosHistorial", hojaInput: "INP_DES", prefijo: "D-", etiqueta: "desvío", articulo: "el", genero: "o" },
    "CAPA": { tabla: "TablaCapas", historial: "TablaCapasHistorial", hojaInput: "INP_CAPAS", prefijo: "C-", etiqueta: "CAPA", articulo: "la", genero: "a" },
    "AFECTACION": { tabla: "TablaAfectacion", historial: "TablaAfectacionHistorial", hojaInput: "INP_AFECT", prefijo: "AF-", etiqueta: "AFECTACION", articulo: "la", genero: "a" }
  };

  const configuracionActiva = CONFIGURACION_ENTIDADES[nombreEntidadNormalizado];
  if (!configuracionActiva) throw new Error(`La entidad '${entidadParam}' no está configurada.`);

  // Asignación de variables descriptivas
  const etiquetaEntidad = configuracionActiva.etiqueta;
  const articuloEntidad = configuracionActiva.articulo;
  const generoEntidad = configuracionActiva.genero;
  const prefijoID = configuracionActiva.prefijo;
  const nombreTablaPrincipal = configuracionActiva.tabla;
  const nombreTablaHistorial = configuracionActiva.historial;
  const nombreHojaEntrada = configuracionActiva.hojaInput;

  // Definición de Tipos e Interfaces
  type ValorCelda = string | number | boolean;
  interface ResultadoAccion { success: boolean; message: string; logLevel: 'EXITO' | 'ERROR' | 'WARN' | 'INFO'; }
  interface MapaColoresUX { [key: string]: { fondo: string; texto: string } }

  const resultadoOperacion: ResultadoAccion = { success: true, message: "Inicio de registro", logLevel: 'INFO' };
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
  let nuevaFilaHistorial: ValorCelda[] | undefined = undefined;

  try {
    // --- I. VALIDACIÓN DE INFRAESTRUCTURA ---
    hojaEntradaWS = workbook.getWorksheet(nombreHojaEntrada);
    hojaMaestrosWS = workbook.getWorksheet("MAESTROS");
    itemClaveSistema = workbook.getNamedItem("SISTEMA_CLAVE");

    if (!hojaEntradaWS) throw new Error(`Infraestructura: No se halló la hoja '${nombreHojaEntrada}'.`);
    if (!hojaMaestrosWS) throw new Error("Infraestructura: Hoja MAESTROS no disponible.");
    if (!itemClaveSistema) throw new Error("Infraestructura: No se halló 'SISTEMA_CLAVE'.");

    claveProteccion = itemClaveSistema.getRange().getText();
    tablaBaseDatos = workbook.getTable(nombreTablaPrincipal);
    
    // Limpieza preventiva de filtros para asegurar lectura correcta de IDs
    const autoFiltro = tablaBaseDatos.getAutoFilter();
    if (autoFiltro) autoFiltro.clearCriteria();
    
    tablaHistorial = workbook.getTable(nombreTablaHistorial);

    if (!tablaBaseDatos) throw new Error(`Infraestructura: La tabla '${nombreTablaPrincipal}' no existe.`);
    if (!tablaHistorial) throw new Error(`Infraestructura: La tabla '${nombreTablaHistorial}' no existe.`);

    // --- II. CAPTURA Y PROCESAMIENTO DE DATOS DEL FORMULARIO ---
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
          // Normalización de datos: Campos vacíos u opcionales se marcan como N/A
          objetoDatosFormulario[claveCampo] = (valorIngresado === null || String(valorIngresado).trim() === "") 
            ? (etiquetaLimpia.endsWith("*") ? "" : "N/A") 
            : valorIngresado;
        }
      });

      const encabezadosTabla = tablaBaseDatos.getHeaderRowRange().getValues()[0].map(h => String(h).toUpperCase().replace(/\s/g, "_"));
      const nombreCampoPrimario = encabezadosTabla[0];
      const listaErroresValidacion: string[] = [];

      // --- 1. VALIDACIÓN TÉCNICA DE FORMATOS (FECHAS) ---
      for (let clave in objetoDatosFormulario) {
        if (clave.includes("FECHA") && objetoDatosFormulario[clave] !== "N/A" && objetoDatosFormulario[clave] !== "") {
          if (isNaN(auxiliarParsearFechaANumero(objetoDatosFormulario[clave]))) {
            listaErroresValidacion.push(`Formato de fecha inválido en: ${clave.replace(/_/g, " ")}`);
          }
        }
      }

      // --- 2. MOTOR DE REGLAS DE NEGOCIO (TRANSVERSAL) ---
      const tablaReglasMaestra = hojaMaestrosWS.getTable("TablaReglas");
      if (tablaReglasMaestra) {
        tablaReglasMaestra.getRangeBetweenHeaderAndTotal().getValues().forEach(regla => {
          if (nombreTablaPrincipal.toUpperCase().includes(String(regla[0]).toUpperCase())) {
            const campoA = String(regla[1]).toUpperCase().replace(/\s/g, "_");
            const operador = String(regla[2]);
            const referenciaRaw = String(regla[3]);
            const mensajeErrorRegla = String(regla[4]);
            const valorAValidar = objetoDatosFormulario[campoA];

            // Regla: Integridad Referencial Simple
            if (operador === "EXISTE_EN") {
              if (valorAValidar && valorAValidar !== "N/A" && referenciaRaw.includes("[")) {
                const [nombreTablaRef, parteColumna] = referenciaRaw.split("[");
                const nombreColumnaRef = parteColumna.replace("]", "");
                const tablaReferencia = workbook.getTable(nombreTablaRef);
                if (tablaReferencia) {
                  const valoresMaestros = tablaReferencia.getColumnByName(nombreColumnaRef).getRangeBetweenHeaderAndTotal().getValues();
                  const existeEnMaestro = valoresMaestros.some(filaMaestra => String(filaMaestra[0]) === String(valorAValidar));
                  if (!existeEnMaestro) listaErroresValidacion.push(mensajeErrorRegla);
                }
              }
            } 
            // Regla: Integridad Referencial con Validación de Estado (Condicional)
            else if (operador === "ESTA_ABIERTO") {
              if (valorAValidar && valorAValidar !== "N/A" && referenciaRaw.includes(";")) {
                  const [parteObjetivo, parteColEstado, valorEsperado] = referenciaRaw.split(";"); 
                  const [nombreTablaRef, parteColID] = parteObjetivo.split("[");
                  const nombreColID = parteColID.replace("]", "");
                  const nombreColEstado = parteColEstado.replace("[", "").replace("]", "");

                  const tablaReferencia = workbook.getTable(nombreTablaRef);
                  if (tablaReferencia) {
                      const matrizIDs = tablaReferencia.getColumnByName(nombreColID).getRangeBetweenHeaderAndTotal().getValues();
                      const matrizEstados = tablaReferencia.getColumnByName(nombreColEstado).getRangeBetweenHeaderAndTotal().getValues();

                      const indiceFilaMaestra = matrizIDs.findIndex(filaID => String(filaID[0]) === String(valorAValidar));

                      if (indiceFilaMaestra !== -1) {
                          const estadoActualEnMaestro = String(matrizEstados[indiceFilaMaestra][0]);
                          if (estadoActualEnMaestro !== valorEsperado) {
                              listaErroresValidacion.push(mensajeErrorRegla);
                          }
                      }
                  }
              }
            }
            // Reglas Cronológicas o Comparativas
            else if (operador === "<=" || operador === ">=") {
              const campoB = referenciaRaw.toUpperCase().replace(/\s/g, "_");
              const valorB = objetoDatosFormulario[campoB];
              if (valorAValidar && valorB && valorAValidar !== "N/A" && valorB !== "N/A") {
                const fechaA = auxiliarParsearFechaANumero(valorAValidar), fechaB = auxiliarParsearFechaANumero(valorB);
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

      // --- 3. VALIDACIÓN DE CAMPOS OBLIGATORIOS ---
      listaCamposObligatorios.forEach(campo => { 
        if (objetoDatosFormulario[campo] === "" || objetoDatosFormulario[campo] === null) {
          listaErroresValidacion.push(`Falta campo obligatorio: ${campo.replace(/_/g, " ")}`); 
        }
      });

      // --- 4. GENERACIÓN DE IDENTIFICADOR ÚNICO CORRELATIVO ---
      const columnaIDs = tablaBaseDatos.getColumnByName(nombreCampoPrimario).getRangeBetweenHeaderAndTotal();
      const valoresIDsExistentes = columnaIDs ? columnaIDs.getValues() as string[][] : [];
      let siguienteNumeroCorrelativo = 1;
      
      if (valoresIDsExistentes.length > 0) {
        // Extraemos la parte numérica de los IDs actuales para encontrar el máximo
        const numerosExtraidos = valoresIDsExistentes.map(filaID => parseInt(String(filaID[0]).replace(/\D/g, '')) || 0);
        siguienteNumeroCorrelativo = Math.max(...numerosExtraidos) + 1;
      }
      const idGeneradoFinal = prefijoID + siguienteNumeroCorrelativo;

      // --- III. PERSISTENCIA DE DATOS (COMMIT) ---
      if (listaErroresValidacion.length > 0) {
        resultadoOperacion.success = false;
        resultadoOperacion.message = "⚠️ Validación Fallida:\n" + listaErroresValidacion.join("\n");
        resultadoOperacion.logLevel = 'WARN';
      } else {
        try {
          // A. INSERCIÓN EN TABLA PRINCIPAL (BASE DE DATOS)
          tablaBaseDatos.getWorksheet().getProtection().unprotect(claveProteccion);
          const nuevaFilaBaseDatos = encabezadosTabla.map(encabezado => {
            if (encabezado === nombreCampoPrimario) return idGeneradoFinal;
            if (encabezado === "ESTADO") return "ABIERTO";
            if (encabezado === "USUARIO") return userEmail;
            if (encabezado === "AUDIT_TRAIL") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
            
            if (objetoDatosFormulario.hasOwnProperty(encabezado)) {
              const valorCampo = objetoDatosFormulario[encabezado];
              return (valorCampo === "" || valorCampo === null || valorCampo === "N/A") ? "N/A" : valorCampo;
            }
            return null;
          });
          tablaBaseDatos.addRow(-1, nuevaFilaBaseDatos);

          // B. INSERCIÓN EN TABLA DE HISTORIAL (AUDITORÍA)
          tablaHistorial.getWorksheet().getProtection().unprotect(claveProteccion);
          const filaRegistroHistorial = (tablaHistorial.getHeaderRowRange().getValues()[0] as string[]).map(h => {
            const headCaps = h.toUpperCase();
            if (headCaps === "ID_EVENTO") {
              const columnaIdEvento = tablaHistorial!.getColumnByName("ID_EVENTO").getRangeBetweenHeaderAndTotal();
              const valoresExistentes = columnaIdEvento ? columnaIdEvento.getValues() as number[][] : [];
              return tablaHistorial!.getRowCount() === 0 ? 1 : Math.max(...valoresExistentes.map(x => Number(x[0]))) + 1;
            }
            if (headCaps === nombreCampoPrimario) return idGeneradoFinal;
            if (headCaps === "USUARIO") return userEmail;
            if (headCaps === "MOTIVO") return "Registro inicial del sistema.";
            if (headCaps === "CAMBIOS") return "[NUEVO REGISTRO CREADO]";
            if (headCaps === "FECHA_CAMBIO") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
            return "";
          });
          tablaHistorial.addRow(-1, filaRegistroHistorial);

          resultadoOperacion.message = `✅ ${articuloEntidad.toUpperCase()} ${etiquetaEntidad.toUpperCase()} #${idGeneradoFinal} se ha cread${generoEntidad} correctamente.`;
          resultadoOperacion.logLevel = 'EXITO';
          
          // Limpieza del formulario tras éxito
          auxiliarLimpiarFormulario(hojaEntradaWS, matrizEtiquetas, indiceFilaInicial, nombreCampoPrimario);

        } catch (errorInterno) {
          resultadoOperacion.success = false;
          resultadoOperacion.logLevel = 'ERROR';
          resultadoOperacion.message = `❌ Error de escritura para ${idGeneradoFinal}: ${String(errorInterno)}`;
        }
      }
    }
  } catch (excepcionSistema) {
    resultadoOperacion.success = false;
    resultadoOperacion.message = `❌ Error Crítico: ${String(excepcionSistema)}`;
    resultadoOperacion.logLevel = 'ERROR';
  } finally {
    // --- IX. PROTOCOLO DE CIERRE Y SEGURIDAD ---
    if (hojaEntradaWS) {
      const passFinal: string = itemClaveSistema ? itemClaveSistema.getRange().getText() : "";
      auxiliarActualizarInterfazUX(hojaEntradaWS, resultadoOperacion, PALETA_COLORES_UX, passFinal);
      if (itemClaveSistema) {
        auxiliarProtegerHoja(hojaEntradaWS, passFinal, resultadoOperacion);
        if (tablaBaseDatos) auxiliarProtegerHoja(tablaBaseDatos.getWorksheet(), passFinal, resultadoOperacion);
        if (tablaHistorial) auxiliarProtegerHoja(tablaHistorial.getWorksheet(), passFinal, resultadoOperacion);
        if (hojaMaestrosWS) auxiliarProtegerHoja(hojaMaestrosWS, passFinal, resultadoOperacion);
      }
    }
  }

  // --- FUNCIONES AUXILIARES (HELPERS) ---

  function auxiliarParsearFechaANumero(valor: ValorCelda): number {
    if (typeof valor === "number") return valor;
    const partes = String(valor).split("/");
    if (partes.length === 3) {
      const objetoFecha = new Date(parseInt(partes[2]), parseInt(partes[1]) - 1, parseInt(p[0]));
      return (objetoFecha.getFullYear() === parseInt(partes[2])) ? objetoFecha.getTime() : NaN;
    }
    return NaN;
  }

  function auxiliarActualizarInterfazUX(hoja: ExcelScript.Worksheet, res: ResultadoAccion, colores: MapaColoresUX, pass: string) {
    const itemFeedback = hoja.getNamedItem("UI_FEEDBACK");
    const itemPreparacion = hoja.getNamedItem("UI_PREPARACION");
    if (!itemFeedback || !itemPreparacion) return; 

    const rangoFeedback = itemFeedback.getRange();
    const rangoPreparacion = itemPreparacion.getRange();
    const estiloUX = colores[res.logLevel];
    const fechaActual = new Date();
    const marcaTiempo = fechaActual.toLocaleTimeString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
    const iconoLatido = (fechaActual.getSeconds() % 2 === 0) ? "⚡" : "✨";
    
    try {
      hoja.getProtection().unprotect(pass);
      rangoFeedback.setValue(`[${marcaTiempo}] ${iconoLatido} ${res.message}`);
      rangoFeedback.getFormat().getFill().setColor(estiloUX.fondo);
      rangoFeedback.getFormat().getFont().setColor(estiloUX.texto);
      rangoFeedback.getFormat().getFont().setBold(true);
      rangoFeedback.getFormat().setWrapText(true);
      
      // Limpieza automática del indicador de preparación tras finalizar el registro
      rangoPreparacion.setValue("");
      rangoPreparacion.getFormat().getFill().clear();
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

  function auxiliarLimpiarFormulario(hoja: ExcelScript.Worksheet, matriz: ValorCelda[][], filaInicio: number, campoId: string) {
    matriz.forEach((fila, i) => {
      const claveCampo = String(fila[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
      // No limpiamos el ID generado ni el usuario responsable, solo campos de entrada
      if (claveCampo !== "" && claveCampo !== campoId && claveCampo !== "USUARIO" && claveCampo !== "MOTIVO") {
        hoja.getRangeByIndexes(i + filaInicio, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
      }
    });
  }
}