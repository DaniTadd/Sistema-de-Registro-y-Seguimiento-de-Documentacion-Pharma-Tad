/**
 * TIPOS E INTERFACES GLOBALES
 */
type ValorCelda = string | number | boolean;
interface ResultadoAccion {
    success: boolean;
    message: string;
    logLevel: 'EXITO' | 'ERROR' | 'WARN' | 'INFO';
    logCambios?: string[];

}
interface MapaColoresUX { [key: string]: { fondo: string; texto: string } }

/**
 * SCRIPT: ENTIDAD_ACTUALIZAR_UNIVERSAL
 * OBJETIVO: Modificar registros existentes con Commit Atómico en memoria, validando reglas y generando Audit Trail estricto.
 */
async function main(
    workbook: ExcelScript.Workbook,
    usuarioEjecutor: string,
    motivoModificacion: string,
    claveFirma: string
) {
    const resultadoOperacion: ResultadoAccion = { success: true, message: "Inicio de actualización", logLevel: 'INFO' };
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
    let tablaIntegridad: ExcelScript.Table | undefined;
    let hojaMaestrosWS: ExcelScript.Worksheet | undefined;
    let claveProteccion: string = "";

    try {
        // --- 1. DETECCIÓN DEL MÓDULO E INFRAESTRUCTURA ---
        const nombreHojaActiva: string = workbook.getActiveWorksheet().getName();
        const CONFIGURACION_ENTIDADES: { [key: string]: { tabla: string, historial: string, prefijo: string, etiqueta: string, articulo: string, genero: string, estadoInicial: string, estadosBloqueantes: string[] } } = {
            "INP_DES": { tabla: "TablaDesvios", historial: "TablaDesviosHistorial", prefijo: "SB-", etiqueta: "DESVÍO", articulo: "el", genero: "o", estadoInicial: "ABIERTO", estadosBloqueantes: ["CERRADO", "ANULADO"] },
            "INP_CAPAS": { tabla: "TablaCapas", historial: "TablaCapasHistorial", prefijo: "CAPA-", etiqueta: "CAPA", articulo: "la", genero: "a", estadoInicial: "ABIERTO", estadosBloqueantes: ["CERRADO", "ANULADO"] },
            "INP_AFECT": { tabla: "TablaAfectacion", historial: "TablaAfectacionHistorial", prefijo: "AF-", etiqueta: "AFECTACION", articulo: "la", genero: "a", estadoInicial: "VIGENTE", estadosBloqueantes: ["ANULADO"] },
            "INP_REC": { tabla: "TablaReclamos", historial: "TablaReclamosHistorial", prefijo: "R-", etiqueta: "RECLAMO", articulo: "el", genero: "o", estadoInicial: "ABIERTO", estadosBloqueantes: ["CERRADO", "ANULADO"] },
            "INP_EQ": { tabla: "TablaEquipos", historial: "TablaEquiposHistorial", prefijo: "EQ-", etiqueta: "EQUIPO", articulo: "el", genero: "o", estadoInicial: "EN USO", estadosBloqueantes: ["DADO DE BAJA", "EN REPARACIÓN"] }
        };

        const configuracionActiva = CONFIGURACION_ENTIDADES[nombreHojaActiva];
        if (!configuracionActiva) {
            throw new Error(`La hoja '${nombreHojaActiva}' no es un formulario de actualización válido.`);
        }

        hojaEntradaWS = workbook.getWorksheet(nombreHojaActiva);
        hojaMaestrosWS = workbook.getWorksheet("MAESTROS");
        const itemClaveSistema = workbook.getNamedItem("SISTEMA_CLAVE");
        const itemSalt = workbook.getNamedItem("SISTEMA_SALT");

        tablaBaseDatos = workbook.getTable(configuracionActiva.tabla);
        tablaHistorial = workbook.getTable(configuracionActiva.historial);
        tablaUsuarios = workbook.getTable("TablaUsuarios");
        tablaIntegridad = workbook.getTable("TablaIntegridad");

        if (!hojaEntradaWS || !hojaMaestrosWS || !itemClaveSistema || !itemSalt || !tablaIntegridad || !tablaUsuarios || !tablaBaseDatos || !tablaHistorial) {
            throw new Error("Infraestructura: Faltan componentes, variables criptográficas o tablas críticas.");
        }

        claveProteccion = String(itemClaveSistema.getRange().getValue());
        const salt: string = String(itemSalt.getRange().getValue());
        hojaEntradaWS.getProtection().unprotect(claveProteccion);

        // --- 2. VERIFICACIÓN DE INTEGRIDAD EN VIVO (ALCOA+) ---
        const dataIntegridad: ValorCelda[][] = tablaIntegridad.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][];
        const regBBDD = dataIntegridad.find((f: ValorCelda[]) => String(f[0]) === configuracionActiva.tabla);
        const regHistorial = dataIntegridad.find((f: ValorCelda[]) => String(f[0]) === configuracionActiva.historial);
        const regUsuarios = dataIntegridad.find((f: ValorCelda[]) => String(f[0]) === "TablaUsuarios");

        const hashInicio: string = "INICIAR";
        const hashBBDD_Maestro: string = regBBDD ? String(regBBDD[1]) : hashInicio;
        const hashHistorial_Maestro: string = regHistorial ? String(regHistorial[1]) : hashInicio;
        const hashUsuarios_Maestro: string = regUsuarios ? String(regUsuarios[1]) : hashInicio;

        const hashVivoBBDD: string = await generarFirmaDigital(tablaBaseDatos, salt);
        const hashVivoHistorial: string = await generarFirmaDigital(tablaHistorial, salt);
        const hashVivoUsuarios: string = await generarFirmaDigital(tablaUsuarios, salt);

        let mensajeErrorIntegridad: string = "";
        if (hashBBDD_Maestro !== hashInicio && hashVivoBBDD !== hashBBDD_Maestro) mensajeErrorIntegridad = `ERROR: Integridad violada en ${configuracionActiva.tabla}.`;
        else if (hashHistorial_Maestro !== hashInicio && hashVivoHistorial !== hashHistorial_Maestro) mensajeErrorIntegridad = `ERROR: Integridad violada en ${configuracionActiva.historial}.`;
        else if (hashUsuarios_Maestro !== hashInicio && hashVivoUsuarios !== hashUsuarios_Maestro) mensajeErrorIntegridad = `ERROR: Integridad violada en TablaUsuarios.`;

        if (mensajeErrorIntegridad !== "") {
            throw new Error(mensajeErrorIntegridad + " Contacte a Calidad.");
        }

        // --- 3. BATCH READ: CAPTURA DEL FORMULARIO EN MEMORIA ---
        const rangoEtiquetas = hojaEntradaWS.getRange("B:B").getUsedRange();
        let continuarEjecucionSESE = true;

        if (rangoEtiquetas) {
            const matrizDatosFormulario: ValorCelda[][] = rangoEtiquetas.getResizedRange(0, 1).getValues() as ValorCelda[][];
            const indiceFilaInicial: number = rangoEtiquetas.getRowIndex();
            const objetoDatosFormulario: { [key: string]: string } = {};
            const listaCamposObligatorios: string[] = [];

            matrizDatosFormulario.forEach((fila: ValorCelda[]) => {
                const etiquetaLimpia: string = String(fila[0]).trim().toUpperCase();
                if (etiquetaLimpia !== "") {
                    const claveCampo: string = etiquetaLimpia.replace("*", "").trim().replace(/\s/g, "_");
                    if (etiquetaLimpia.endsWith("*")) listaCamposObligatorios.push(claveCampo);

                    const valorIngresado: ValorCelda = fila[1];
                    objetoDatosFormulario[claveCampo] = (valorIngresado === null || String(valorIngresado).trim() === "")
                        ? (etiquetaLimpia.endsWith("*") ? "" : "N/A")
                        : String(valorIngresado);
                }
            });

            // --- PUENTE DE NORMALIZACIÓN DINÁMICA (PATRÓN FILTRO_) ---
            Object.keys(objetoDatosFormulario).forEach((claveCampo: string) => {
                if (claveCampo.startsWith("FILTRO_")) {
                    const claveColumnaReal: string = claveCampo.replace("FILTRO_", "");
                    const valorSeleccionado: string = objetoDatosFormulario[claveCampo];

                    if (valorSeleccionado !== "" && valorSeleccionado !== "N/A") {
                        objetoDatosFormulario[claveColumnaReal] = valorSeleccionado;
                    }
                }
            });
            // ---------------------------------------------------------

            const listaErroresValidacion: string[] = [];
            const usuarioIngresado: string = usuarioEjecutor.trim();
            let actualizarFirmaUsuario: boolean = false;

            if (usuarioIngresado === "") {
                listaErroresValidacion.push("Se requiere usuario autenticado para firmar cambios.");
            } else {
                const matrizUsuarios: ValorCelda[][] = tablaUsuarios.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][];
                const indiceUsuario: number = matrizUsuarios.findIndex((f: ValorCelda[]) => String(f[0]).trim().toUpperCase() === usuarioIngresado.trim().toUpperCase());

                if (indiceUsuario === -1) {
                    listaErroresValidacion.push(`Firma inválida: Usuario '${usuarioIngresado}' no registrado.`);
                } else {
                    const hashGuardado: string = String(matrizUsuarios[indiceUsuario][1]);
                    const hashFirmaCalculado: string = sha256(claveFirma);

                    if (hashGuardado === "INICIAR" || hashGuardado === "0" || hashGuardado === "") {
                        matrizUsuarios[indiceUsuario][1] = hashFirmaCalculado;
                        tablaUsuarios.getWorksheet().getProtection().unprotect(claveProteccion);
                        tablaUsuarios.getRangeBetweenHeaderAndTotal().setValues(matrizUsuarios);
                        actualizarFirmaUsuario = true;
                    } else if (hashGuardado !== hashFirmaCalculado) {
                        listaErroresValidacion.push("Firma inválida: Credenciales incorrectas.");
                    }
                }
            }

            const encabezadosTabla: string[] = tablaBaseDatos.getHeaderRowRange().getValues()[0].map((h: ValorCelda) => String(h).toUpperCase().replace(/\s/g, "_"));
            const nombreCampoPrimario: string = encabezadosTabla.find(header => header.startsWith("ID_")) || encabezadosTabla[0];

            // Telemetría de diagnóstico empírico para inspeccionar las claves mapeadas en memoria RAM
            console.log("--- [SYS_DEBUG_ACTUALIZAR] ---");
            console.log("Nombre Campo Primario esperado en BBDD:", nombreCampoPrimario);
            console.log("Objeto Datos Formulario completo:", JSON.stringify(objetoDatosFormulario));

            // Resolución robusta: Busca el ID tanto en la clave primaria directa como en su equivalente BUSQUEDA_
            let idABuscar: string = String(objetoDatosFormulario[nombreCampoPrimario] || "").trim().toUpperCase();
            if (!idABuscar || idABuscar === "N/A") {
                const claveBusquedaAlternativa = Object.keys(objetoDatosFormulario).find(k => k.startsWith("BUSQUEDA_") && k.includes(nombreCampoPrimario.replace("ID_", "")));
                if (claveBusquedaAlternativa) {
                    idABuscar = String(objetoDatosFormulario[claveBusquedaAlternativa] || "").trim().toUpperCase();
                }
            }
            console.log("ID a buscar resuelto:", idABuscar);
            console.log("------------------------------");

            if (!idABuscar || idABuscar === "" || idABuscar === "N/A") {
                listaErroresValidacion.push(`Se requiere un ID de ${configuracionActiva.etiqueta} válido para actualizar.`);
                continuarEjecucionSESE = false;
            }

            // =====================================================================
            // --- MOTOR DE AUTORIZACIÓN (ACL - ZERO TRUST) ---
            // =====================================================================
            if (continuarEjecucionSESE && listaErroresValidacion.length === 0) {
                // 1. Inferencia Temprana de Transacción
                const estadoEvaluadoACL = objetoDatosFormulario["ESTADO"] !== undefined ? String(objetoDatosFormulario["ESTADO"]).toUpperCase().trim() : "";
                let transaccionRequerida = "MODIFICACION";

                if (estadoEvaluadoACL === "BAJA" || estadoEvaluadoACL === "DADO DE BAJA") {
                    transaccionRequerida = "BAJA";
                } else if (estadoEvaluadoACL === "ANULADO" || estadoEvaluadoACL === "ANULACION") {
                    transaccionRequerida = "ANULACION";
                }

                // 2. Extracción de Matriz de Permisos
                const tablaPermisos = workbook.getTable("TablaPermisos");
                if (!tablaPermisos) {
                    throw new Error("Integridad fallida: TablaPermisos no localizada en el repositorio local.");
                }

                const matrizACL = tablaPermisos.getRangeBetweenHeaderAndTotal().getTexts();
                let idxPermiso = 0;
                let autorizacionConcedida = false;

                // 3. Evaluación Lineal (Fail-Closed)
                while (idxPermiso < matrizACL.length && !autorizacionConcedida) {
                    const usuarioBase = String(matrizACL[idxPermiso][0]).trim().toUpperCase();
                    const transaccionBase = String(matrizACL[idxPermiso][1]).trim().toUpperCase();
                    const estadoAcceso = String(matrizACL[idxPermiso][2]).trim().toUpperCase();

                    const usuarioCoincide = (usuarioBase === usuarioIngresado.toUpperCase());
                    const transaccionCoincide = (transaccionBase === transaccionRequerida || transaccionBase === "*");
                    const accesoPermitido = (estadoAcceso === "CONCEDIDO");

                    if (usuarioCoincide && transaccionCoincide && accesoPermitido) {
                        autorizacionConcedida = true;
                    }
                    idxPermiso++;
                }

                // 4. Compuerta Lógica
                if (!autorizacionConcedida) {
                    listaErroresValidacion.push(`ACCESO DENEGADO: El usuario ${usuarioIngresado} no posee privilegios para la transacción [${transaccionRequerida}].`);
                    continuarEjecucionSESE = false;
                }
            }
            // =====================================================================

            if (continuarEjecucionSESE && listaErroresValidacion.length === 0) {
                const matrizValoresDB: ValorCelda[][] = tablaBaseDatos.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][];
                const matrizTextosDB: string[][] = tablaBaseDatos.getRangeBetweenHeaderAndTotal().getTexts();
                let indiceFilaEncontrada: number = -1;
                let contadorFilas: number = 0;
                let registroEncontrado: boolean = false;

                while (contadorFilas < matrizValoresDB.length && !registroEncontrado) {
                    if (String(matrizValoresDB[contadorFilas][encabezadosTabla.indexOf(nombreCampoPrimario)]).toUpperCase() === idABuscar) {
                        indiceFilaEncontrada = contadorFilas;
                        registroEncontrado = true;
                    }
                    contadorFilas++;
                }

                if (!registroEncontrado) {
                    listaErroresValidacion.push(`${configuracionActiva.etiqueta.charAt(0).toUpperCase() + configuracionActiva.etiqueta.slice(1)} #${idABuscar} no encontrad${configuracionActiva.genero}.`);
                } else {
                    const valorEstadoActual: string = String(matrizValoresDB[indiceFilaEncontrada][encabezadosTabla.indexOf("ESTADO")]).toUpperCase();

                    if (configuracionActiva.estadosBloqueantes.includes(valorEstadoActual)) {
                        listaErroresValidacion.push(`Protección ALCOA+: No se permite modificar registros en estado ${valorEstadoActual}.`);
                    } else {

                        // --- CONTROL DE CONCURRENCIA OPTIMISTA (OCC) ---
                        const indiceAuditTrailBBDD: number = encabezadosTabla.indexOf("AUDIT_TRAIL");
                        if (indiceAuditTrailBBDD !== -1) {
                            const estampaTiempoBBDD: string = String(matrizValoresDB[indiceFilaEncontrada][indiceAuditTrailBBDD]).trim();
                            const estampaTiempoFormulario: string = String(objetoDatosFormulario["AUDIT_TRAIL"] || "").trim();

                            // Si el formulario posee estampa y difiere de la BBDD actual, otro usuario la modificó
                            if (estampaTiempoFormulario !== "" && estampaTiempoFormulario !== "N/A" && estampaTiempoBBDD !== estampaTiempoFormulario) {
                                listaErroresValidacion.push(`Conflicto de concurrencia (Lost Update): El registro fue modificado por otro usuario en la sesión. Vuelva a buscar el registro para actualizar su vista.`);
                            }
                        }

                        // --- 4. DETECCIÓN DE CAMBIOS Y REGLAS DE NEGOCIO ---
                        const valoresFilaOriginal: ValorCelda[] = [...matrizValoresDB[indiceFilaEncontrada]];
                        const logDeCambios: string[] = [];
                        const cambiosPendientes: { columna: number, valor: string }[] = [];

                        listaCamposObligatorios.forEach((campo: string) => {
                            if (objetoDatosFormulario[campo] === "" || objetoDatosFormulario[campo] === null || objetoDatosFormulario[campo] === "N/A") {
                                listaErroresValidacion.push(`Falta campo obligatorio: ${campo.replace(/_/g, " ")}`);
                            }
                        });

                        encabezadosTabla.forEach((header: string, colIdx: number) => {
                            if (["AUDIT_TRAIL", "ESTADO", "USUARIO", nombreCampoPrimario].includes(header)) return;

                            if (objetoDatosFormulario.hasOwnProperty(header)) {
                                const textoOriginal: string = matrizTextosDB[indiceFilaEncontrada][colIdx];
                                const valorNuevo: string = String(objetoDatosFormulario[header]);
                                let textoNuevoFormateado: string = "";
                                let esCambioReal: boolean = false;

                                if (header.includes("FECHA") && valorNuevo !== "N/A" && valorNuevo !== "") {
                                    const fechaOriginalEstandarizada: string = auxiliarFormatearFechaAString(textoOriginal);
                                    const fechaNuevaEstandarizada: string = auxiliarFormatearFechaAString(valorNuevo);

                                    if (fechaOriginalEstandarizada !== fechaNuevaEstandarizada) {
                                        esCambioReal = true;
                                        textoNuevoFormateado = fechaNuevaEstandarizada;
                                    }
                                } else {
                                    textoNuevoFormateado = valorNuevo;
                                    if (textoOriginal.trim() !== textoNuevoFormateado.trim()) {
                                        esCambioReal = true;
                                    }
                                }

                                if (esCambioReal) {
                                    logDeCambios.push(`${header}: [${textoOriginal}] -> [${textoNuevoFormateado}]`);
                                    cambiosPendientes.push({ columna: colIdx, valor: objetoDatosFormulario[header] });
                                }
                            }
                        });

                        for (const clave in objetoDatosFormulario) {
                            if (clave.includes("FECHA") && objetoDatosFormulario[clave] !== "N/A" && objetoDatosFormulario[clave] !== "") {
                                if (isNaN(auxiliarParsearFechaANumero(objetoDatosFormulario[clave]))) {
                                    listaErroresValidacion.push(`Formato de fecha inválido en: ${clave.replace(/_/g, " ")}`);
                                }
                            }
                        }

                        // --- VI. MOTOR DE REGLAS DINÁMICO (BATCHING I/O Y EVALUACIÓN EN RAM) ---

                        // 1. Elevación de Variables (Hoisting para Event Sourcing)
                        const reglasAplicables: ValorCelda[][] = [];
                        const reglasEventos: ValorCelda[][] = [];
                        const diccionarioTablasAuxiliares: { [nombreTabla: string]: { encabezados: string[], datos: ValorCelda[][], maxIdEvento?: number } } = {};

                        const tablaReglasMaestra = hojaMaestrosWS.getTable("TablaReglas");
                        if (tablaReglasMaestra) {
                            const matrizReglasTodas: ValorCelda[][] = tablaReglasMaestra.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][];

                            matrizReglasTodas.forEach((regla: ValorCelda[]) => {
                                const entidadRegla: string = String(regla[0]).toUpperCase();

                                if (configuracionActiva.tabla.toUpperCase().includes(entidadRegla)) {
                                    reglasAplicables.push(regla);

                                    const operador: string = String(regla[2]);
                                    const referenciaRaw: string = String(regla[3]);

                                    if (operador === "EXISTE_EN" || operador === "ESTADO_DISTINTO_A" || operador === "ES_UNICO_ALFANUMERICO") {
                                        const nombreTablaRequerida: string = referenciaRaw.split("[")[0];

                                        if (nombreTablaRequerida && !diccionarioTablasAuxiliares[nombreTablaRequerida]) {
                                            const tablaAuxiliar = workbook.getTable(nombreTablaRequerida);
                                            if (tablaAuxiliar) {
                                                const encabezadosAuxiliares: string[] = tablaAuxiliar.getHeaderRowRange().getValues()[0].map((h: ValorCelda) => String(h).toUpperCase().replace(/\s/g, "_"));
                                                const datosAuxiliares: ValorCelda[][] = tablaAuxiliar.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][];
                                                diccionarioTablasAuxiliares[nombreTablaRequerida] = { encabezados: encabezadosAuxiliares, datos: datosAuxiliares };
                                            }
                                        }
                                    } // --- NUEVO: INTERCEPCIÓN DE EVENTOS ---
                                    else if (operador === "LOG_EVENTO") {
                                        reglasEventos.push(regla);
                                        const nombreTablaRequerida: string = referenciaRaw.trim();
                                        if (nombreTablaRequerida && !diccionarioTablasAuxiliares[nombreTablaRequerida]) {
                                            const tablaAuxiliar = workbook.getTable(nombreTablaRequerida);
                                            if (tablaAuxiliar) {
                                                const encabezadosAuxiliares: string[] = tablaAuxiliar.getHeaderRowRange().getValues()[0].map((h: ValorCelda) => String(h).toUpperCase().replace(/\s/g, "_"));
                                                let maxIdObj: number = 0;
                                                if (tablaAuxiliar.getRowCount() > 0) {
                                                    const colIdEv = tablaAuxiliar.getColumnByName("ID_EVENTO");
                                                    if (colIdEv) {
                                                        const vals = colIdEv.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][];
                                                        maxIdObj = Math.max(...vals.map(v => Number(v[0])));
                                                    }
                                                }
                                                diccionarioTablasAuxiliares[nombreTablaRequerida] = {
                                                    encabezados: encabezadosAuxiliares,
                                                    datos: [],
                                                    maxIdEvento: maxIdObj
                                                };
                                            }
                                        }
                                    }
                                    // --------------------------------------
                                }
                            });

                            // 2. Fase de Procesamiento (Evaluación 100% en RAM)
                            reglasAplicables.forEach((regla: ValorCelda[]) => {
                                const claveCampoA: string = String(regla[1]).toUpperCase().replace(/\s/g, "_");
                                const operador: string = String(regla[2]);
                                const referenciaRaw: string = String(regla[3]);
                                const mensajeErrorRegla: string = String(regla[4]);
                                const valorAValidar: string = objetoDatosFormulario[claveCampoA];

                                if (valorAValidar && valorAValidar !== "N/A") {

                                    if (operador === "EXISTE_EN") {
                                        const partesReferencia: string[] = referenciaRaw.split("[");
                                        const nombreTabla: string = partesReferencia[0];
                                        const colIdDestino: string = partesReferencia[1].replace("]", "");

                                        const tablaEnMemoria = diccionarioTablasAuxiliares[nombreTabla];
                                        if (tablaEnMemoria) {
                                            const indiceColumnaId: number = tablaEnMemoria.encabezados.indexOf(colIdDestino);
                                            if (indiceColumnaId !== -1) {
                                                const registroExiste: boolean = tablaEnMemoria.datos.some((fila: ValorCelda[]) => String(fila[indiceColumnaId]) === String(valorAValidar));
                                                if (!registroExiste) listaErroresValidacion.push(mensajeErrorRegla);
                                            }
                                        }
                                    }
                                    else if (operador === "ESTADO_DISTINTO_A") {
                                        const segmentosCondicion: string[] = referenciaRaw.split(";");
                                        if (segmentosCondicion.length === 3) {
                                            const partesTabla: string[] = segmentosCondicion[0].split("[");
                                            const nombreTabla: string = partesTabla[0];
                                            const colIdDestino: string = partesTabla[1].replace("]", "");
                                            const colEstadoDestino: string = segmentosCondicion[1].replace("[", "").replace("]", "");
                                            const estadoProhibido: string = segmentosCondicion[2].toUpperCase();

                                            const tablaEnMemoria = diccionarioTablasAuxiliares[nombreTabla];
                                            if (tablaEnMemoria) {
                                                const indiceColId: number = tablaEnMemoria.encabezados.indexOf(colIdDestino);
                                                const indiceColEstado: number = tablaEnMemoria.encabezados.indexOf(colEstadoDestino);

                                                if (indiceColId !== -1 && indiceColEstado !== -1) {
                                                    let estadoEncontrado: string = "";
                                                    let indiceBusqueda: number = 0;
                                                    let registroHallado: boolean = false;

                                                    while (indiceBusqueda < tablaEnMemoria.datos.length && !registroHallado) {
                                                        if (String(tablaEnMemoria.datos[indiceBusqueda][indiceColId]).toUpperCase() === String(valorAValidar).toUpperCase()) {
                                                            estadoEncontrado = String(tablaEnMemoria.datos[indiceBusqueda][indiceColEstado]).toUpperCase();
                                                            registroHallado = true;
                                                        }
                                                        indiceBusqueda++;
                                                    }

                                                    if (registroHallado && estadoEncontrado === estadoProhibido) {
                                                        listaErroresValidacion.push(mensajeErrorRegla);
                                                    }
                                                }
                                            }
                                        }
                                    } else if (operador === "ES_UNICO_ALFANUMERICO") {
                                        const nombreTablaDestino: string = referenciaRaw.trim();
                                        const tablaEnMemoria = diccionarioTablasAuxiliares[nombreTablaDestino];

                                        if (tablaEnMemoria) {
                                            const indiceColumnaId: number = tablaEnMemoria.encabezados.indexOf(claveCampoA);
                                            const indicePK: number = tablaEnMemoria.encabezados.indexOf(nombreCampoPrimario);

                                            if (indiceColumnaId !== -1 && indicePK !== -1) {
                                                // Sanitización: Solo alfanuméricos en mayúscula
                                                const valorNormalizadoInput: string = String(valorAValidar).replace(/[^a-zA-Z0-9]/g, "").toUpperCase();
                                                let colisionDetectada: boolean = false;
                                                let idxBusqueda: number = 0;

                                                while (idxBusqueda < tablaEnMemoria.datos.length && !colisionDetectada) {
                                                    const valorFilaOriginal = tablaEnMemoria.datos[idxBusqueda][indiceColumnaId];
                                                    const pkFilaOriginal = String(tablaEnMemoria.datos[idxBusqueda][indicePK]).toUpperCase();

                                                    if (valorFilaOriginal && String(valorFilaOriginal) !== "N/A" && String(valorFilaOriginal) !== "") {
                                                        const valorNormalizadoDB: string = String(valorFilaOriginal).replace(/[^a-zA-Z0-9]/g, "").toUpperCase();

                                                        // Exclusión del registro actual (evita falso positivo contra sí mismo)
                                                        if (valorNormalizadoDB === valorNormalizadoInput && pkFilaOriginal !== idABuscar) {
                                                            colisionDetectada = true;
                                                        }
                                                    }
                                                    idxBusqueda++;
                                                }

                                                if (colisionDetectada) {
                                                    listaErroresValidacion.push(mensajeErrorRegla);
                                                }
                                            }
                                        }
                                    } else if (operador === "<=" || operador === ">=") {
                                        const claveCampoB: string = referenciaRaw.toUpperCase().replace(/\s/g, "_");

                                        // Recuperación híbrida: Si el campo B no está en el formulario actual, se extrae de la copia original en memoria (BBDD)
                                        const valorB: string = objetoDatosFormulario[claveCampoB] !== undefined && objetoDatosFormulario[claveCampoB] !== "N/A"
                                            ? objetoDatosFormulario[claveCampoB]
                                            : String(valoresFilaOriginal[encabezadosTabla.indexOf(claveCampoB)]);

                                        if (valorAValidar && valorB && valorAValidar !== "N/A" && valorB !== "N/A") {
                                            const fechaA: number = auxiliarParsearFechaANumero(valorAValidar);
                                            const fechaB: number = auxiliarParsearFechaANumero(valorB);
                                            if (isNaN(fechaA) || isNaN(fechaB)) {
                                                listaErroresValidacion.push(`Error: Formato de fecha inválido comparando reglas de negocio para ${claveCampoA.replace(/_/g, " ")}.`);
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

                        // --- 5. DECISIÓN DE PERSISTENCIA Y COMMIT ATÓMICO ---
                        if (listaErroresValidacion.length === 0) {
                            if (logDeCambios.length === 0) {
                                resultadoOperacion.message = "ℹ️ No se detectaron cambios reales para guardar.";
                                resultadoOperacion.logLevel = 'INFO';
                            } else if (!motivoModificacion || motivoModificacion.trim() === "") {
                                listaErroresValidacion.push("Se requiere justificar el motivo de la modificación (ALCOA+).");
                            } else {
                                // --- NUEVO: PREPARACIÓN DE EVENTOS DINÁMICOS (DELTA DETECTION) ---
                                const eventosAInsertar: { tablaDestino: string, filaData: string[] }[] = [];

                                reglasEventos.forEach((regla: ValorCelda[]) => {
                                    const campoFormulario: string = String(regla[1]).toUpperCase().replace(/\s/g, "_");
                                    const tablaDestino: string = String(regla[3]).trim();
                                    const idxCampoBD = encabezadosTabla.indexOf(campoFormulario);

                                    if (idxCampoBD !== -1 && objetoDatosFormulario[campoFormulario] !== undefined && objetoDatosFormulario[campoFormulario] !== "N/A") {
                                        const valorNuevo = String(objetoDatosFormulario[campoFormulario]);
                                        const valorOriginal = String(valoresFilaOriginal[idxCampoBD]);

                                        // Lógica Delta: Solo registra si el valor mutó realmente
                                        if (valorNuevo !== valorOriginal) {
                                            const infoTabla = diccionarioTablasAuxiliares[tablaDestino];
                                            if (infoTabla) {
                                                infoTabla.maxIdEvento = (infoTabla.maxIdEvento || 0) + 1;

                                                const nuevaFilaEvento: string[] = infoTabla.encabezados.map((enc: string) => {
                                                    if (enc === "ID_EVENTO") return String(infoTabla.maxIdEvento);
                                                    if (enc === "FECHA_EVENTO" || enc === "FECHA") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
                                                    if (enc === "USUARIO") return usuarioIngresado;
                                                    if (enc === "MOTIVO") return motivoModificacion.trim();
                                                    if (enc === nombreCampoPrimario) return idABuscar;

                                                    if (objetoDatosFormulario[enc] !== undefined && objetoDatosFormulario[enc] !== "N/A") {
                                                        return objetoDatosFormulario[enc];
                                                    }
                                                    return "";
                                                });
                                                eventosAInsertar.push({ tablaDestino, filaData: nuevaFilaEvento });
                                            }
                                        }
                                    }
                                });
                                // -----------------------------------------------------------------

                                tablaBaseDatos.getWorksheet().getProtection().unprotect(claveProteccion);
                                const rangoFilaAfectada = tablaBaseDatos.getRangeBetweenHeaderAndTotal().getRow(indiceFilaEncontrada);

                                // Ejecución de Escritura Atómica en RAM
                                const filaAtómicaModificada: ValorCelda[] = [...valoresFilaOriginal];
                                cambiosPendientes.forEach(cambio => {
                                    filaAtómicaModificada[cambio.columna] = cambio.valor;
                                });

                                const idxAuditTrail = encabezadosTabla.indexOf("AUDIT_TRAIL");
                                if (idxAuditTrail !== -1) filaAtómicaModificada[idxAuditTrail] = new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });

                                const idxUsuario = encabezadosTabla.indexOf("USUARIO");
                                if (idxUsuario !== -1) filaAtómicaModificada[idxUsuario] = usuarioIngresado;

                                // Volcado Síncrono a Disco (Batch Write)
                                rangoFilaAfectada.setValues([filaAtómicaModificada]);

                                tablaHistorial.getWorksheet().getProtection().unprotect(claveProteccion);
                                const nuevaFilaHistorial: ValorCelda[] = (tablaHistorial.getHeaderRowRange().getValues()[0] as string[]).map((h: string) => {
                                    const headCaps = h.toUpperCase();
                                    if (headCaps === "ID_EVENTO") return (tablaHistorial!.getRowCount() === 0 ? 1 : Math.max(...(tablaHistorial!.getColumnByName("ID_EVENTO").getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][]).map((v: ValorCelda[]) => Number(v[0]))) + 1);
                                    if (headCaps === nombreCampoPrimario) return idABuscar;
                                    if (headCaps === "USUARIO") return usuarioIngresado;
                                    if (headCaps === "MOTIVO") return motivoModificacion.trim();
                                    if (headCaps === "CAMBIOS") return logDeCambios.join(" | ");
                                    if (headCaps === "FECHA_CAMBIO") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
                                    return "";
                                });
                                tablaHistorial.addRow(-1, nuevaFilaHistorial);

                                // --- NUEVO: ESCRITURA JIT DE EVENTOS ---
                                eventosAInsertar.forEach((evento: { tablaDestino: string, filaData: string[] }) => {
                                    const tablaTarget = workbook.getTable(evento.tablaDestino);
                                    if (tablaTarget) {
                                        tablaTarget.getWorksheet().getProtection().unprotect(claveProteccion);
                                        tablaTarget.addRow(-1, evento.filaData);
                                    }
                                });
                                // ---------------------------------------

                                // =====================================================================
                                // --- TRIGGER DE NOTIFICACIÓN ASÍNCRONA: ACTUALIZAR (MODIFICACIÓN/BAJA) ---
                                // =====================================================================

                                const idRecibido = idABuscar;
                                const estadoEvaluado = objetoDatosFormulario["ESTADO"] !== undefined ? String(objetoDatosFormulario["ESTADO"]).toUpperCase().trim() : "";

                                let transaccionDetectada: TipoTransaccion = "MODIFICACION";

                                if (estadoEvaluado === "BAJA" || estadoEvaluado === "DADO DE BAJA") {
                                    transaccionDetectada = "BAJA";
                                } else if (estadoEvaluado === "ANULADO" || estadoEvaluado === "ANULACION") {
                                    transaccionDetectada = "ANULACION";
                                }

                                const diccionarioFilaActualizada: { [key: string]: string } = {};

                                // 1. Lectura del Vector Atómico (Single Source of Truth)
                                encabezadosTabla.forEach((header: string, index: number) => {
                                    let valorFinal = filaAtómicaModificada[index] !== undefined
                                        ? String(filaAtómicaModificada[index])
                                        : "N/A";

                                    // 2. Formateo dinámico de fechas crudas (Serial a DD/MM/YYYY)
                                    if (header.includes("FECHA") && valorFinal !== "N/A" && valorFinal !== "") {
                                        valorFinal = auxiliarFormatearFechaAString(valorFinal);
                                    }

                                    diccionarioFilaActualizada[header] = valorFinal;
                                });

                                // 3. Inyección explícita del MOTIVO (No es columna de la BBDD matriz, pero se exige en el reporte)
                                diccionarioFilaActualizada["MOTIVO_MODIFICACION"] = motivoModificacion.trim();

                                const payloadActualizacion: PayloadNotificacion = {
                                    entidad: configuracionActiva.tabla.replace("Tabla", "").toUpperCase(),
                                    transaccion: transaccionDetectada,
                                    idRegistro: idRecibido,
                                    usuario: usuarioIngresado,
                                    fechaHora: new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false }),
                                    datosFila: diccionarioFilaActualizada,
                                    logCambios: logDeCambios.length > 0 ? logDeCambios : ["Modificación general de atributos sin diferencias detectadas."]
                                };

                                auxiliarEncolarNotificacion(workbook, payloadActualizacion, claveProteccion);
                                // =====================================================================

                                resultadoOperacion.message = `✅ ${idABuscar} actualizad${configuracionActiva.genero} con éxito.`;
                                resultadoOperacion.logLevel = 'EXITO';
                                auxiliarLimpiarFormulario(hojaEntradaWS, matrizDatosFormulario, indiceFilaInicial, nombreCampoPrimario);

                                // --- ACTUALIZACIÓN DE SELLOS MAESTROS ---
                                tablaIntegridad.getWorksheet().getProtection().unprotect(claveProteccion);
                                const nuevoSelloBBDD: string = await generarFirmaDigital(tablaBaseDatos, salt);
                                const nuevoSelloHistorial: string = await generarFirmaDigital(tablaHistorial, salt);

                                let matrizSeg: ValorCelda[][] = tablaIntegridad.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][];

                                const idxBBDD: number = matrizSeg.findIndex((f: ValorCelda[]) => String(f[0]) === configuracionActiva.tabla);
                                if (idxBBDD !== -1) matrizSeg[idxBBDD][1] = nuevoSelloBBDD;

                                const idxHist: number = matrizSeg.findIndex((f: ValorCelda[]) => String(f[0]) === configuracionActiva.historial);
                                if (idxHist !== -1) matrizSeg[idxHist][1] = nuevoSelloHistorial;

                                if (actualizarFirmaUsuario) {
                                    const nuevoSelloUsuarios: string = await generarFirmaDigital(tablaUsuarios, salt);
                                    const idxUsu: number = matrizSeg.findIndex((f: ValorCelda[]) => String(f[0]) === "TablaUsuarios");
                                    if (idxUsu !== -1) matrizSeg[idxUsu][1] = nuevoSelloUsuarios;
                                }

                                // --- NUEVO: SELLADO CRIPTOGRÁFICO DE EVENTOS ---
                                const tablasEventosUnicas = Array.from(new Set(eventosAInsertar.map(e => e.tablaDestino)));
                                for (let i = 0; i < tablasEventosUnicas.length; i++) {
                                    const nombreT: string = tablasEventosUnicas[i];
                                    const tablaTargetObj = workbook.getTable(nombreT);
                                    if (tablaTargetObj) {
                                        const idxSeg: number = matrizSeg.findIndex((f: ValorCelda[]) => String(f[0]) === nombreT);
                                        if (idxSeg !== -1) matrizSeg[idxSeg][1] = await generarFirmaDigital(tablaTargetObj, salt);
                                    }
                                }
                                // -----------------------------------------------

                                tablaIntegridad.getRangeBetweenHeaderAndTotal().setValues(matrizSeg);
                            }
                        }
                    }
                }
            }

            // Consolidación de Errores UX
            if (listaErroresValidacion.length > 0) {
                resultadoOperacion.success = false;
                resultadoOperacion.message = "⚠️ Validación Fallida:\n" + listaErroresValidacion.join("\n");
                resultadoOperacion.logLevel = 'WARN';
            }
        }
    } catch (e) {
        resultadoOperacion.success = false;
        resultadoOperacion.logLevel = 'ERROR';
        resultadoOperacion.message = `❌ Fallo Crítico de Infraestructura: ${String(e)}`;
    } finally {
        // --- 6. PROTOCOLO DE CIERRE SEGURO ---
        if (hojaEntradaWS) {
            auxiliarActualizarInterfazUX(hojaEntradaWS, resultadoOperacion, PALETA_COLORES_UX, claveProteccion);
            auxiliarProtegerHoja(hojaEntradaWS, claveProteccion, resultadoOperacion);
            if (tablaBaseDatos) auxiliarProtegerHoja(tablaBaseDatos.getWorksheet(), claveProteccion, resultadoOperacion);
            if (tablaHistorial) auxiliarProtegerHoja(tablaHistorial.getWorksheet(), claveProteccion, resultadoOperacion);
            if (tablaUsuarios) auxiliarProtegerHoja(tablaUsuarios.getWorksheet(), claveProteccion, resultadoOperacion);
            if (tablaIntegridad) auxiliarProtegerHoja(tablaIntegridad.getWorksheet(), claveProteccion, resultadoOperacion);
        }
    }

    // --- FUNCIONES AUXILIARES (HELPERS) ---

    function auxiliarParsearFechaANumero(v: ValorCelda): number {
        if (typeof v === "number") return v;
        const numeroSerial = Number(v);
        if (!isNaN(numeroSerial) && String(v).trim() !== "") return numeroSerial;

        const partes: string[] = String(v).split("/");
        if (partes.length === 3) {
            const fechaObjeto = new Date(parseInt(partes[2]), parseInt(partes[1]) - 1, parseInt(partes[0]));
            return (fechaObjeto.getFullYear() === parseInt(partes[2])) ? fechaObjeto.getTime() : NaN;
        }
        return NaN;
    }

    function auxiliarActualizarInterfazUX(hoja: ExcelScript.Worksheet, res: ResultadoAccion, colores: MapaColoresUX, pass: string): void {
        const itemF = hoja.getNamedItem("UI_FEEDBACK");
        const itemP = hoja.getNamedItem("UI_PREPARACION");

        if (itemF) {
            const rf = itemF.getRange();
            const est = colores[res.logLevel];
            try {
                hoja.getProtection().unprotect(pass);
                rf.setValue(`[${new Date().toLocaleTimeString('es-AR', { hour12: false })}] ${res.message}`);
                rf.getFormat().getFill().setColor(est.fondo);
                rf.getFormat().getFont().setColor(est.texto);
                rf.getFormat().getFont().setBold(true);

                if (itemP) {
                    itemP.getRange().setValue("");
                    itemP.getRange().getFormat().getFill().clear();
                }
                rf.select();
            } catch (e) {
                // Requerimiento Zero Silent Failures
                console.log("Falla al aplicar UX. La hoja podría estar bloqueada con otra clave: ", e);
            }
        } else {
            // Requerimiento Zero Silent Failures de Infraestructura
            console.log("Falla de Infraestructura evitada: No se encontró el ítem nombrado 'UI_FEEDBACK'.");
        }
    }

    function auxiliarProtegerHoja(hoja: ExcelScript.Worksheet | undefined, pass: string, res: ResultadoAccion): void {
        if (hoja) {
            try {
                hoja.getProtection().protect({ allowAutoFilter: true }, pass);
            } catch (e) {
                res.success = false;
                res.logLevel = 'ERROR';
                res.message += ` | Falla crítica protegiendo: ${hoja.getName()}.`;
            }
        }
    }

    function auxiliarLimpiarFormulario(hoja: ExcelScript.Worksheet, matriz: ValorCelda[][], filaInicio: number, campoId: string): void {
        matriz.forEach((fila: ValorCelda[], i: number) => {
            const claveCampo: string = String(fila[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
            if (claveCampo !== "" && claveCampo !== campoId && claveCampo !== "MOTIVO") {
                hoja.getRangeByIndexes(i + filaInicio, 2, 1, 1).clear(ExcelScript.ClearApplyTo.contents);
            }
        });
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

function auxiliarFormatearFechaAString(valor: ValorCelda): string {
    const numeroSerial = Number(valor);
    if (!isNaN(numeroSerial) && String(valor).trim() !== "") {
        const milisegundos: number = Math.round((numeroSerial - 25569) * 86400 * 1000) + 43200000;
        const fechaObjeto = new Date(milisegundos);
        const dia: string = String(fechaObjeto.getDate()).padStart(2, '0');
        const mes: string = String(fechaObjeto.getMonth() + 1).padStart(2, '0');
        return `${dia}/${mes}/${fechaObjeto.getFullYear()}`;
    }

    const partes: string[] = String(valor).split("/");
    if (partes.length === 3) {
        const dia: string = String(parseInt(partes[0])).padStart(2, '0');
        const mes: string = String(parseInt(partes[1])).padStart(2, '0');
        return `${dia}/${mes}/${partes[2]}`;
    }

    return String(valor).trim();
}

// --- INTERFACES DE NOTIFICACIÓN ---
type TipoTransaccion = 'CREACION' | 'MODIFICACION' | 'CAMBIO_ESTADO' | 'ANULACION' | 'BAJA';

interface PayloadNotificacion {
    entidad: string;
    transaccion: TipoTransaccion;
    idRegistro: string;
    usuario: string;
    fechaHora: string;
    datosFila: { [key: string]: string };
    logCambios: string[];
}

// --- HELPER DE MENSAJERÍA (SESE & Fail-Safe) ---
function auxiliarEncolarNotificacion(
    workbook: ExcelScript.Workbook,
    payload: PayloadNotificacion,
    claveProteccion: string
): void {
    try {
        const tablaOutbox = workbook.getTable("TablaNotificaciones_Outbox");

        if (tablaOutbox) {
            let textoDetalleCompleto: string = "";
            const clavesDatos = Object.keys(payload.datosFila);
            let idxDatos = 0;

            while (idxDatos < clavesDatos.length) {
                const clave = clavesDatos[idxDatos];
                textoDetalleCompleto += `${clave.replace(/_/g, " ")}: ${payload.datosFila[clave]}\n`;
                idxDatos++;
            }

            let textoResumenCambios: string = "";
            if (payload.logCambios && payload.logCambios.length > 0) {
                let idxLog = 0;
                while (idxLog < payload.logCambios.length) {
                    textoResumenCambios += `${payload.logCambios[idxLog]}\n`;
                    idxLog++;
                }
            } else {
                textoResumenCambios = "No se detectaron modificaciones específicas. (Alta/Creación)";
            }

            const paqueteJSON = JSON.stringify({
                Entidad: payload.entidad,
                Transaccion: payload.transaccion,
                Firma: `${payload.usuario} - ${payload.fechaHora}`,
                ResumenCambios: textoResumenCambios.trim(),
                DetalleCompleto: textoDetalleCompleto.trim()
            });

            const idMensaje = `MSG-${new Date().getTime()}`;
            const nuevaFila = [
                idMensaje,
                payload.fechaHora,
                "PENDIENTE",
                paqueteJSON
            ];

            const hojaOutbox = tablaOutbox.getWorksheet();
            hojaOutbox.getProtection().unprotect(claveProteccion);
            tablaOutbox.addRow(-1, nuevaFila);
            // REMOVIDO PARA PERMITIR BORRADO DESDE POWER AUTOMATE:
            // hojaOutbox.getProtection().protect({ allowAutoFilter: true }, claveProteccion);
        } else {
            console.log("[SYS_WARNING] 'TablaNotificaciones_Outbox' no encontrada.");
        }
    } catch (errorOutbox) {
        console.log(`[SYS_WARNING] Falla aislada al encolar notificación: ${(errorOutbox as Error).message}`);
    }
}
