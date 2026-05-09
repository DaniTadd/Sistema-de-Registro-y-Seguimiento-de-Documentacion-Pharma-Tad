/**
 * TIPOS E INTERFACES GLOBALES
 */
type ValorCelda = string | number | boolean;
interface ResultadoAccion { success: boolean; message: string; logLevel: 'EXITO' | 'ERROR' | 'WARN' | 'INFO'; }
interface MapaColoresUX { [key: string]: { fondo: string; texto: string } }

/**
 * SCRIPT: ENTIDAD_BUSCAR_UNIVERSAL
 * OBJETIVO: Localizar un registro por ID y poblar el formulario garantizando la integridad previa de los datos.
 */
async function main(workbook: ExcelScript.Workbook, idABuscar: string) {
  // --- 1. DETECCIÓN DEL MÓDULO ---
  const nombreHojaActiva: string = workbook.getActiveWorksheet().getName();

  const CONFIGURACION_ENTIDADES: { [key: string]: { tabla: string, historial: string, prefijo: string, etiqueta: string, articulo: string, genero: string } } = {
    "INP_DES": { tabla: "TablaDesvios", historial: "TablaDesviosHistorial", prefijo: "D-", etiqueta: "desvío", articulo: "el", genero: "o" },
    "INP_CAPAS": { tabla: "TablaCapas", historial: "TablaCapasHistorial", prefijo: "C-", etiqueta: "CAPA", articulo: "la", genero: "a" },
    "INP_AFECT": { tabla: "TablaAfectacion", historial: "TablaAfectacionHistorial", prefijo: "AF-", etiqueta: "AFECTACION", articulo: "la", genero: "a" }
  };

  const configuracionActiva = CONFIGURACION_ENTIDADES[nombreHojaActiva];
  if (!configuracionActiva) throw new Error(`La hoja '${nombreHojaActiva}' no es un formulario válido de búsqueda.`);

  const etiquetaEntidad: string = configuracionActiva.etiqueta;
  const articuloEntidad: string = configuracionActiva.articulo;
  const generoEntidad: string = configuracionActiva.genero;
  const nombreTablaPrincipal: string = configuracionActiva.tabla;
  const nombreTablaHistorial: string = configuracionActiva.historial;
  const nombreHojaEntrada: string = nombreHojaActiva;

  const resultadoOperacion: ResultadoAccion = { success: true, message: "Inicio de búsqueda", logLevel: 'INFO' };
  const PALETA_COLORES_UX: MapaColoresUX = {
    EXITO: { fondo: "#D4EDDA", texto: "#155724" },
    ERROR: { fondo: "#F8D7DA", texto: "#721C24" },
    WARN: { fondo: "#FFF3CD", texto: "#856404" },
    INFO: { fondo: "#E2E3E5", texto: "#383D41" }
  };

  let hojaEntradaWS: ExcelScript.Worksheet | undefined;
  let tablaBaseDatos: ExcelScript.Table | undefined;
  let tablaHistorial: ExcelScript.Table | undefined;
  let tablaUsuarios: ExcelScript.Table | undefined;
  let itemClaveSistema: ExcelScript.NamedItem | undefined;
  
  let claveProteccion: string = "";

  // --- 0.1 INICIALIZACIÓN Y SEGURIDAD ---
  tablaBaseDatos = workbook.getTable(nombreTablaPrincipal);
  tablaHistorial = workbook.getTable(nombreTablaHistorial);
  tablaUsuarios = workbook.getTable("TablaUsuarios");
  
  const itemSalt = workbook.getNamedItem("SISTEMA_SALT");
  if (!itemSalt) throw new Error("Infraestructura: Falta configuración criptográfica (SISTEMA_SALT).");
  const salt: string = String(itemSalt.getRange().getValue());
  
  const tablaIntegridad = workbook.getTable("TablaIntegridad");
  if (!tablaIntegridad || !tablaUsuarios || !tablaBaseDatos || !tablaHistorial) {
      throw new Error("Infraestructura: Faltan tablas críticas del sistema.");
  }

  const dataIntegridad: ValorCelda[][] = tablaIntegridad.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][];
  
  const regBBDD = dataIntegridad.find((f: ValorCelda[]) => String(f[0]) === nombreTablaPrincipal);
  const regHistorial = dataIntegridad.find((f: ValorCelda[]) => String(f[0]) === nombreTablaHistorial);
  const regUsuarios = dataIntegridad.find((f: ValorCelda[]) => String(f[0]) === "TablaUsuarios");
  let hashinicio: string = "INICIAR"
  const hashBBDD_Maestro: string = regBBDD ? String(regBBDD[1]) : hashinicio;
  const hashHistorial_Maestro: string = regHistorial ? String(regHistorial[1]) : hashinicio;
  const hashUsuarios_Maestro: string = regUsuarios ? String(regUsuarios[1]) : hashinicio;

  // --- 0.2 VERIFICACIÓN DE INTEGRIDAD EN VIVO ---
  const hashVivoBBDD: string = await generarFirmaDigital(tablaBaseDatos, salt);
  const hashVivoHistorial: string = await generarFirmaDigital(tablaHistorial, salt);
  const hashVivoUsuarios: string = await generarFirmaDigital(tablaUsuarios, salt);

  let mensajeErrorIntegridad: string = "";

  try {
    hojaEntradaWS = workbook.getWorksheet(nombreHojaEntrada);
    itemClaveSistema = workbook.getNamedItem("SISTEMA_CLAVE");

    if (!hojaEntradaWS || !itemClaveSistema) {
        throw new Error("Infraestructura: Componentes visuales o claves maestras no disponibles.");
    }

    claveProteccion = String(itemClaveSistema.getRange().getValue());
    hojaEntradaWS.getProtection().unprotect(claveProteccion);
    // --- GUARDIA ALCOA+ (Solo lectura, pero bloquea si hay corrupción) ---
    if (hashBBDD_Maestro !== hashinicio && hashVivoBBDD !== hashBBDD_Maestro) {
      mensajeErrorIntegridad = `ERROR: Integridad violada en ${nombreTablaPrincipal}.`;
    } else if (hashHistorial_Maestro !== hashinicio && hashVivoHistorial !== hashHistorial_Maestro) {
      mensajeErrorIntegridad = `ERROR: Integridad violada en ${nombreTablaHistorial}.`;
    } else if (hashUsuarios_Maestro !== hashinicio && hashVivoUsuarios !== hashUsuarios_Maestro) {
      mensajeErrorIntegridad = `ERROR: Integridad violada en TablaUsuarios.`;
    }

    if (mensajeErrorIntegridad !== "") {
      resultadoOperacion.success = false;
      resultadoOperacion.message = mensajeErrorIntegridad + " Lectura bloqueada. Contacte a Calidad.";
      resultadoOperacion.logLevel = 'ERROR';
      throw new Error(mensajeErrorIntegridad);
    }

    // Limpieza de filtros para asegurar búsqueda en todo el set de datos
    const autoFiltro = tablaBaseDatos.getAutoFilter();
    if (autoFiltro) autoFiltro.clearCriteria();

    // --- II. PROCESAMIENTO DE LA BÚSQUEDA ---
    const idBuscadoNormalizado = idABuscar ? idABuscar.trim().toUpperCase() : "";

    if (!idBuscadoNormalizado || idBuscadoNormalizado === "") {
      resultadoOperacion.success = false;
      resultadoOperacion.message = `Se requiere un ID válido para buscar.`;
      resultadoOperacion.logLevel = 'WARN';
    } else {
      const rangoEtiquetasFormulario = hojaEntradaWS.getRange("B:B").getUsedRange();
      const encabezadosTabla: string[] = tablaBaseDatos.getHeaderRowRange().getValues()[0].map((h: ValorCelda) => String(h).toUpperCase().replace(/\s/g, "_"));
      const nombreCampoPrimario: string = encabezadosTabla[0];

      if (rangoEtiquetasFormulario) {
        const matrizEtiquetas: ValorCelda[][] = rangoEtiquetasFormulario.getValues() as ValorCelda[][];
        const indiceFilaInicialForm: number = rangoEtiquetasFormulario.getRowIndex();
        const mapaCoordenadasFormulario: { [key: string]: number } = {};

        // Mapeo de coordenadas del formulario
        matrizEtiquetas.forEach((fila: ValorCelda[], i: number) => {
          const etiquetaLimpia: string = String(fila[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
          if (etiquetaLimpia !== "") {
            mapaCoordenadasFormulario[etiquetaLimpia] = i + indiceFilaInicialForm;
          }
        });

        // --- III. BÚSQUEDA EN BASE DE DATOS ---
        const matrizValoresDB: ValorCelda[][] = tablaBaseDatos.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][];
        const matrizTextosDB: string[][] = tablaBaseDatos.getRangeBetweenHeaderAndTotal().getTexts();
        let indiceFilaEncontrada: number = -1;
        let contadorFilas: number = 0;
        let registroEncontrado: boolean = false;

        while (contadorFilas < matrizValoresDB.length && !registroEncontrado) {
          if (String(matrizValoresDB[contadorFilas][encabezadosTabla.indexOf(nombreCampoPrimario)]).toUpperCase() === idBuscadoNormalizado) {
            indiceFilaEncontrada = contadorFilas;
            registroEncontrado = true;
          }
          contadorFilas++;
        }

        if (registroEncontrado) {
          // 1. Limpieza del formulario antes de poblar
          matrizEtiquetas.forEach((fila: ValorCelda[], i: number) => {
            const campoLimpio: string = String(fila[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
            if (campoLimpio !== "") {
              hojaEntradaWS!.getRangeByIndexes(i + indiceFilaInicialForm, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
            }
          });

          // 2. Población de datos
          encabezadosTabla.forEach((nombreEncabezado: string, indiceColumna: number) => {
            if (mapaCoordenadasFormulario[nombreEncabezado] !== undefined) {
              const rangoDestino = hojaEntradaWS!.getRangeByIndexes(mapaCoordenadasFormulario[nombreEncabezado], 2, 1, 1);

              if (nombreEncabezado.includes("FECHA")) {
                rangoDestino.setValue(matrizValoresDB[indiceFilaEncontrada][indiceColumna]); 
                rangoDestino.setNumberFormatLocal("dd/mm/yyyy");
              } else {
                rangoDestino.setValue(matrizTextosDB[indiceFilaEncontrada][indiceColumna]);
              }
            }
          });

          // 3. Refuerzo de Identidad Visual
          if (mapaCoordenadasFormulario[nombreCampoPrimario] !== undefined) {
            hojaEntradaWS!.getRangeByIndexes(mapaCoordenadasFormulario[nombreCampoPrimario], 2, 1, 1).setValue(idBuscadoNormalizado);
          }

          resultadoOperacion.message = `✅ ${etiquetaEntidad.charAt(0).toUpperCase() + etiquetaEntidad.slice(1)} #${idBuscadoNormalizado} cargad${generoEntidad} con éxito. Datos verificados (ALCOA+).`;
          resultadoOperacion.logLevel = 'EXITO';

        } else {
          resultadoOperacion.success = false;
          resultadoOperacion.message = `❌ ${articuloEntidad.charAt(0).toUpperCase() + articuloEntidad.slice(1)} ${etiquetaEntidad} #${idBuscadoNormalizado} no existe.`;
          resultadoOperacion.logLevel = 'ERROR';
        }
      }
    }
  } catch (e) {
    resultadoOperacion.success = false;
    resultadoOperacion.logLevel = (resultadoOperacion.logLevel === 'INFO') ? 'ERROR' : resultadoOperacion.logLevel;
    if (resultadoOperacion.message === "Inicio de búsqueda") resultadoOperacion.message = `❌ Error: ${String(e)}`;
  } finally {
    // --- IX. PROTOCOLO DE CIERRE Y SEGURIDAD ---
    if (hojaEntradaWS) {
      auxiliarActualizarInterfazUX(hojaEntradaWS, resultadoOperacion, PALETA_COLORES_UX, claveProteccion);
      auxiliarProtegerHoja(hojaEntradaWS, claveProteccion, resultadoOperacion);
    }
  }

  // --- FUNCIONES AUXILIARES (HELPERS) ---

  function auxiliarActualizarInterfazUX(hoja: ExcelScript.Worksheet, res: ResultadoAccion, colores: MapaColoresUX, pass: string): void {
    const itemFeedback = hoja.getNamedItem("UI_FEEDBACK");
    if (!itemFeedback) return; 
    const rangoFeedback = itemFeedback.getRange();
    const estiloUX = colores[res.logLevel];
    try {
      hoja.getProtection().unprotect(pass);
      rangoFeedback.setValue(`[${new Date().toLocaleTimeString('es-AR', { hour12: false })}] ${res.message}`);
      rangoFeedback.getFormat().getFill().setColor(estiloUX.fondo);
      rangoFeedback.getFormat().getFont().setColor(estiloUX.texto);
      rangoFeedback.getFormat().getFont().setBold(true);
      rangoFeedback.getFormat().setWrapText(true);
    } catch (e) {}
  }

  function auxiliarProtegerHoja(hoja: ExcelScript.Worksheet | undefined, pass: string, res: ResultadoAccion): void {
    if (hoja) {
      try { hoja.getProtection().protect({ allowAutoFilter: true }, pass); }
      catch (e) { res.message += ` [⚠️ Seguridad: ${hoja.getName()}]`; }
    }
  }
}

// --- CORE CRIPTOGRÁFICO ---
async function generarFirmaDigital(tabla: ExcelScript.Table, salt: string): Promise<string> {
  const r = tabla.getRangeBetweenHeaderAndTotal();
  return sha256((r ? JSON.stringify(r.getValues()) : "TABLA_VACIA") + salt);
}

function sha256(s: string): string {
  let a: number = 0, b: number = 0, c: number = 0, d: number = 0, e: number = 0, f: number = 0, g: number = 0, h: number = 0;
  const chrsz: number = 8;
  function safe_add(x: number, y: number): number {
    const lsw: number = (x & 0xFFFF) + (y & 0xFFFF); 
    const msw: number = (x >> 16) + (y >> 16) + (lsw >> 16);
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
    const K: number[] = [0x428A2F98, 0x71374491, 0xB5C0FBCF, 0xE9B5DBA5, 0x3956C25B, 0x59F111F1, 0x923F82A4, 0xAB1C5ED5, 0xD807AA98, 0x12835B01, 0x243185BE, 0x550C7DC3, 0x72BE5D74, 0x80DEB1FE, 0x9BDC06A7, 0xC19BF174, 0xE49B69C1, 0xEFBE4786, 0x0FC19DC6, 0x240CA1CC, 0x2DE92C6F, 0x4A7484AA, 0x5CB0A9DC, 0x76F988DA, 0x983E5152, 0xA831C66D, 0xB00327C8, 0xBF597FC7, 0xC6E00BF3, 0xD5A79147, 0x06CA6351, 0x14292967, 0x27B70A85, 0x2E1B2138, 0x4D2C6DFC, 0x53380D13, 0x650A7354, 0x766A0ABB, 0x81C2C92E, 0x92722C85, 0xA2BFE8A1, 0xA81A664B, 0xC24B8B70, 0xC76C51A3, 0xD192E819, 0xD6990624, 0xF40E3585, 0x106AA070, 0x19A4C116, 0x1E376C08, 0x2748774C, 0x34B0BCB5, 0x391C0CB3, 0x4ED8AA4A, 0x5B9CCA4F, 0x682E6FF3, 0x748F82EE, 0x78A5636F, 0x84C87814, 0x8CC70208, 0x90BEFFFA, 0xA4506CEB, 0xBEF9A3F7, 0xC67178F2];
    const H: number[] = [0x6A09E667, 0xBB67AE85, 0x3C6EF372, 0xA54FF53A, 0x510E527F, 0x9B05688C, 0x1F83D9AB, 0x5BE0CD19];
    const W: number[] = new Array(64);
    m[l >> 5] |= 0x80 << (24 - l % 32); m[((l + 64 >> 9) << 4) + 15] = l;
    for (let i = 0; i < m.length; i += 16) {
      a = H[0]; b = H[1]; c = H[2]; d = H[3]; e = H[4]; f = H[5]; g = H[6]; h = H[7];
      for (let j = 0; j < 64; j++) {
        if (j < 16) W[j] = m[j + i];
        else W[j] = safe_add(safe_add(safe_add(Gamma1256(W[j - 2]), W[j - 7]), Gamma0256(W[j - 15])), W[j - 16]);
        const T1: number = safe_add(safe_add(safe_add(safe_add(h, Sigma1256(e)), Ch(e, f, g)), K[j]), W[j]);
        const T2: number = safe_add(Sigma0256(a), Maj(a, b, c));
        h = g; g = f; f = e; e = safe_add(d, T1); d = c; c = b; b = a; a = safe_add(T1, T2);
      }
      H[0] = safe_add(a, H[0]); H[1] = safe_add(b, H[1]); H[2] = safe_add(c, H[2]); H[3] = safe_add(d, H[3]);
      H[4] = safe_add(e, H[4]); H[5] = safe_add(f, H[5]); H[6] = safe_add(g, H[6]); H[7] = safe_add(h, H[7]);
    }
    return H;
  }
  function str2binb(str: string): number[] {
    const bin: number[] = []; const mask: number = (1 << chrsz) - 1;
    for (let i = 0; i < str.length * chrsz; i += chrsz) bin[i >> 5] |= (str.charCodeAt(i / chrsz) & mask) << (24 - i % 32);
    return bin;
  }
  function binb2hex(binarray: number[]): string {
    const hex_tab = "0123456789abcdef"; let str = "";
    for (let i = 0; i < binarray.length * 4; i++) str += hex_tab.charAt((binarray[i >> 2] >> ((3 - i % 4) * 8 + 4)) & 0xF) + hex_tab.charAt((binarray[i >> 2] >> ((3 - i % 4) * 8)) & 0xF);
    return str;
  }
  return binb2hex(core_sha256(str2binb(s), s.length * chrsz));
}