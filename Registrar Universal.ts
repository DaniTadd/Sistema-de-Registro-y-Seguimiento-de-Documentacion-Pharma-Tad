/**
 * SCRIPT: ENTIDAD_REGISTRAR_UNIVERSAL
 * OBJETIVO: Crear nuevos registros (Desvíos/CAPAS/Afectaciones) con generación de ID automático y validación inicial.
 * GARANTÍA: Asegura que todo nuevo ingreso cumpla con las reglas de integridad referencial y cronológica.
 */
function main(
  workbook: ExcelScript.Workbook,
  entidadParam: string,           // Origen Power Automate: "DESVIO", "CAPA" o "AFECTACION"
  hashBBDD_Maestro: string,      // Desde SharePoint
  hashHistorial_Maestro: string, // Desde SharePoint
  salt: string,                  // Desde SharePoint (la Sal)
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

  // --- 0.1 INICIALIZACIÓN DE TABLAS PARA VALIDACIÓN ---
  tablaBaseDatos = workbook.getTable(nombreTablaPrincipal);
  tablaHistorial = workbook.getTable(nombreTablaHistorial);

  // --- 0.2 VERIFICACIÓN DE INTEGRIDAD ALCOA+ (Pre-Operación) ---
  const hashVivoBBDD = await generarFirmaDigital(tablaBaseDatos, salt);
  const hashVivoHistorial = await generarFirmaDigital(tablaHistorial, salt);

  // Variable para capturar si hay error
  let mensajeErrorIntegridad = "";

  if (hashBBDD_Maestro !== "0" && hashVivoBBDD !== hashBBDD_Maestro) {
    mensajeErrorIntegridad = `ERROR: Integridad violada en ${nombreTablaPrincipal}.`;
  } else if (hashHistorial_Maestro !== "0" && hashVivoHistorial !== hashHistorial_Maestro) {
    mensajeErrorIntegridad = `ERROR: Integridad violada en Historial.`;
  }

  // SI HAY ERROR, INFORMAMOS AL USUARIO EN LA PLANILLA ANTES DE ABORTAR
  if (mensajeErrorIntegridad !== "") {
    // 1. Buscamos la hoja de entrada y el rango de estado
    const hojaInput = workbook.getWorksheet(nombreHojaEntrada);
    const celdaEstado = hojaInput.getNamedItem("UI_FEEDBACK");

    // 2. Aplicamos el formato de la PALETA_COLORES_UX que ya tenés
    celdaEstado.setValue(mensajeErrorIntegridad + " Contacte al Administrador.");
    celdaEstado.getFormat().getFill().setColor(PALETA_COLORES_UX.ERROR.fondo);
    celdaEstado.getFormat().getFont().setColor(PALETA_COLORES_UX.ERROR.texto);

    // 3. Ahora sí, lanzamos el error para que Power Automate se entere
    throw new Error(mensajeErrorIntegridad);
  }

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

          resultadoOperacion.message = `✅ ${articuloEntidad.toUpperCase()} ${etiquetaEntidad.toUpperCase()} #${idGeneradoFinal} se ha cread${generoEntidad} correctamente y se ha sellado digitalmente.`;
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
    const nuevoSelloBBDD = await generarFirmaDigital(tablaBaseDatos, salt);
    const nuevoSelloHistorial = await generarFirmaDigital(tablaHistorial, salt);

    // Este es el RETURN que verá Power Automate
    return {
      success: true,
      message: `Registro de ${etiquetaEntidad} completado con éxito.`,
      nuevoHashBBDD: nuevoSelloBBDD,
      nuevoHashHistorial: nuevoSelloHistorial,
      status: "SUCCESS"}
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
/**
 * Genera una firma digital SHA-256 única.
 * Versión "Pure JS": Funciona en cualquier entorno (Excel Desktop, Web y Power Automate).
 */
async function generarFirmaDigital(tabla: ExcelScript.Table, salt: string): Promise<string> {
    const rangoCuerpo = tabla.getRangeBetweenHeaderAndTotal();
    const datosParaHash = rangoCuerpo 
        ? JSON.stringify(rangoCuerpo.getValues()) + salt 
        : "TABLA_VACIA" + salt;

    return sha256(datosParaHash);
}

/**
 * Implementación manual de SHA-256 para entornos sin API 'crypto'.
 */
function sha256(s: string): string {
    const chrsz = 8;
    const hexcase = 0;

    function safe_add(x: number, y: number): number {
        const lsw = (x & 0xFFFF) + (y & 0xFFFF);
        const msw = (x >> 16) + (y >> 16) + (lsw >> 16);
        return (msw << 16) | (lsw & 0xFFFF);
    }

    function S(X: number, n: number): number { return (X >>> n) | (X << (32 - n)); }
    function R(X: number, n: number): number { return (X >>> n); }
    function Ch(x: number, y: number, z: number): number { return ((x & y) ^ ((~x) & z)); }
    function Maj(x: number, y: number, z: number): number { return ((x & y) ^ (x & z) ^ (y & z)); }
    function Sigma0256(x: number): number { return (S(x, 2) ^ S(x, 13) ^ S(x, 22)); }
    function Sigma1256(x: number): number { return (S(x, 6) ^ S(x, 11) ^ S(x, 25)); }
    function Gamma0256(x: number): number { return (S(x, 7) ^ S(x, 18) ^ R(x, 3)); }
    function Gamma1256(x: number): number { return (S(x, 17) ^ S(x, 19) ^ R(x, 10)); }

    function core_sha256(m: number[], l: number): number[] {
        const K = [0x428A2F98, 0x71374491, 0xB5C0FBCF, 0xE9B5DBA5, 0x3956C25B, 0x59F111F1, 0x923F82A4, 0xAB1C5ED5, 0xD807AA98, 0x12835B01, 0x243185BE, 0x550C7DC3, 0x72BE5D74, 0x80DEB1FE, 0x9BDC06A7, 0xC19BF174, 0xE49B69C1, 0xEFBE4786, 0x0FC19DC6, 0x240CA1CC, 0x2DE92C6F, 0x4A7484AA, 0x5CB0A9DC, 0x76F988DA, 0x983E5152, 0xA831C66D, 0xB00327C8, 0xBF597FC7, 0xC6E00BF3, 0xD5A79147, 0x06CA6351, 0x14292967, 0x27B70A85, 0x2E1B2138, 0x4D2C6DFC, 0x53380D13, 0x650A7354, 0x766A0ABB, 0x81C2C92E, 0x92722C85, 0xA2BFE8A1, 0xA81A664B, 0xC24B8B70, 0xC76C51A3, 0xD192E819, 0xD6990624, 0xF40E3585, 0x106AA070, 0x19A4C116, 0x1E376C08, 0x2748774C, 0x34B0BCB5, 0x391C0CB3, 0x4ED8AA4A, 0x5B9CCA4F, 0x682E6FF3, 0x748F82EE, 0x78A5636F, 0x84C87814, 0x8CC70208, 0x90BEFFFA, 0xA4506CEB, 0xBEF9A3F7, 0xC67178F2];
        const HASH = [0x6A09E667, 0xBB67AE85, 0x3C6EF372, 0xA54FF53A, 0x510E527F, 0x9B05688C, 0x1F83D9AB, 0x5BE0CD19];
        const W = new Array(64);
        let a, b, c, d, e, f, g, h;

        m[l >> 5] |= 0x80 << (24 - l % 32);
        m[((l + 64 >> 9) << 4) + 15] = l;

        for (let i = 0; i < m.length; i += 16) {
            a = HASH[0]; b = HASH[1]; c = HASH[2]; d = HASH[3]; e = HASH[4]; f = HASH[5]; g = HASH[6]; h = HASH[7];
            for (let j = 0; j < 64; j++) {
                if (j < 16) W[j] = m[j + i];
                else W[j] = safe_add(safe_add(safe_add(Gamma1256(W[j - 2]), W[j - 7]), Gamma0256(W[j - 15])), W[j - 16]);
                const T1 = safe_add(safe_add(safe_add(safe_add(h, Sigma1256(e)), Ch(e, f, g)), K[j]), W[j]);
                const T2 = safe_add(Sigma0256(a), Maj(a, b, c));
                h = g; g = f; f = e; e = safe_add(d, T1); d = c; c = b; b = a; a = safe_add(T1, T2);
            }
            HASH[0] = safe_add(a, HASH[0]); HASH[1] = safe_add(b, HASH[1]); HASH[2] = safe_add(c, HASH[2]); HASH[3] = safe_add(d, HASH[3]);
            HASH[4] = safe_add(e, HASH[4]); HASH[5] = safe_add(f, HASH[5]); HASH[6] = safe_add(g, HASH[6]); HASH[7] = safe_add(h, HASH[7]);
        }
        return HASH;
    }

    function str2binb(str: string): number[] {
        const bin: number[] = [];
        const mask = (1 << chrsz) - 1;
        for (let i = 0; i < str.length * chrsz; i += chrsz) {
            bin[i >> 5] |= (str.charCodeAt(i / chrsz) & mask) << (24 - i % 32);
        }
        return bin;
    }

    function binb2hex(binarray: number[]): string {
        const hex_tab = hexcase ? "0123456789ABCDEF" : "0123456789abcdef";
        let str = "";
        for (let i = 0; i < binarray.length * 4; i++) {
            str += hex_tab.charAt((binarray[i >> 2] >> ((3 - i % 4) * 8 + 4)) & 0xF) + hex_tab.charAt((binarray[i >> 2] >> ((3 - i % 4) * 8)) & 0xF);
        }
        return str;
    }

    return binb2hex(core_sha256(str2binb(s), s.length * chrsz));
}