/**
 * TIPOS E INTERFACES GLOBALES
 */
type ValorCelda = string | number | boolean;
interface ResultadoAccion { success: boolean; message: string; logLevel: 'EXITO' | 'ERROR' | 'WARN' | 'INFO'; }
interface MapaColoresUX { [key: string]: { fondo: string; texto: string } }

/**
 * SCRIPT: ENTIDAD_CAMBIAR_ESTADO
 * OBJETIVO: Gestionar cierre/reapertura (Toggle) con escaneo Top-Down de dependencias, Batching I/O y Commit Atómico.
 */
async function main(
  workbook: ExcelScript.Workbook,
  usuarioEjecutor: string, 
  idConfirmacion: string, 
  motivoDeCambio: string, 
  claveFirma: string
) {
    const resultadoOperacion: ResultadoAccion = { success: true, message: "Inicio de transición", logLevel: 'INFO' };
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
        // --- 1. DETECCIÓN DEL MÓDULO E INFRAESTRUCTURA (SESE) ---
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
            throw new Error(`La hoja '${nombreHojaActiva}' no es un formulario válido para transición de estado.`);
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
        const rangoEtiquetasFormulario = hojaEntradaWS.getRange("B:B").getUsedRange();
        let continuarEjecucionSESE = true;

        if (rangoEtiquetasFormulario) {
            const matrizDatosFormulario: ValorCelda[][] = rangoEtiquetasFormulario.getResizedRange(0, 1).getValues() as ValorCelda[][];
            const indiceFilaInicial: number = rangoEtiquetasFormulario.getRowIndex();
            const objetoDatosFormulario: { [key: string]: string } = {};

            matrizDatosFormulario.forEach((fila: ValorCelda[]) => {
                const etiquetaLimpia: string = String(fila[0]).trim().toUpperCase();
                if (etiquetaLimpia !== "") {
                    const claveCampo: string = etiquetaLimpia.replace("*", "").trim().replace(/\s/g, "_");
                    const valorIngresado: ValorCelda = fila[1]; 
                    objetoDatosFormulario[claveCampo] = (valorIngresado === null || String(valorIngresado).trim() === "") ? "" : String(valorIngresado);
                }
            });

            const listaErroresValidacion: string[] = [];
            const usuarioIngresado: string = usuarioEjecutor.trim();
            let actualizarFirmaUsuario: boolean = false;

            const encabezadosTabla: string[] = tablaBaseDatos.getHeaderRowRange().getValues()[0].map((h: ValorCelda) => String(h).toUpperCase().replace(/\s/g, "_"));
            const nombreCampoPrimario: string = encabezadosTabla[0];
            const idEnPantalla: string = String(objetoDatosFormulario[nombreCampoPrimario] || "").trim().toUpperCase();
            const idTipeado: string = idConfirmacion.trim().toUpperCase();
            const idABuscar: string = String(objetoDatosFormulario[nombreCampoPrimario] || "").trim().toUpperCase();

            // --- 4. AUTENTICACIÓN Y VALIDACIONES PREVIAS ---
            if (!usuarioIngresado || usuarioIngresado === "") {
                listaErroresValidacion.push("El campo USUARIO es obligatorio para la firma electrónica.");
            } else {
                const matrizUsuarios: ValorCelda[][] = tablaUsuarios.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][];
                const indiceUsuario: number = matrizUsuarios.findIndex((f: ValorCelda[]) => String(f[0]).trim().toUpperCase() === usuarioIngresado.trim().toUpperCase());
                
                if (indiceUsuario === -1) {
                    listaErroresValidacion.push(`Firma inválida: Usuario '${usuarioIngresado}' no registrado.`);
                } else {
                    const hashGuardado: string = String(matrizUsuarios[indiceUsuario][1]);
                    const hashFirmaCalculado: string = sha256(claveFirma);

                    if (hashGuardado === "INICIAR") {
                        matrizUsuarios[indiceUsuario][1] = hashFirmaCalculado;
                        tablaUsuarios.getWorksheet().getProtection().unprotect(claveProteccion);
                        tablaUsuarios.getRangeBetweenHeaderAndTotal().setValues(matrizUsuarios);
                        actualizarFirmaUsuario = true;
                    } else if (hashGuardado !== hashFirmaCalculado) {
                        listaErroresValidacion.push("Firma inválida: Credenciales incorrectas.");
                    }
                }
            }

            if (!idABuscar || idABuscar === "") listaErroresValidacion.push(`Se requiere un ID de ${configuracionActiva.etiqueta} válido.`);
            if (idTipeado !== idEnPantalla) listaErroresValidacion.push(`Mismatch: El ID ingresado [${idTipeado}] no coincide con la pantalla [${idEnPantalla}].`);
            if (!motivoDeCambio || motivoDeCambio.trim() === "") listaErroresValidacion.push(`La justificación es obligatoria (ALCOA+).`);

            if (listaErroresValidacion.length === 0) {
                const matrizValoresDB: ValorCelda[][] = tablaBaseDatos.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][];
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
                    const estadoActual: string = String(matrizValoresDB[indiceFilaEncontrada][encabezadosTabla.indexOf("ESTADO")]).toUpperCase();

                    if (estadoActual === "ANULADO") {
                        listaErroresValidacion.push(`Un registro ANULADO es definitivo y no permite cambios de estado.`);
                    } else {
                        // --- RESOLUCIÓN DINÁMICA DE ESTADO (UNIVERSAL) ---
                        let nuevoEstado: string = "";
                        const inputNuevoEstado: string = objetoDatosFormulario["NUEVO_ESTADO"];

                        if (inputNuevoEstado && inputNuevoEstado !== "") {
                            nuevoEstado = inputNuevoEstado.toUpperCase();
                            const estadosValidos = [configuracionActiva.estadoInicial, ...configuracionActiva.estadosBloqueantes];
                            if (!estadosValidos.includes(nuevoEstado)) {
                                listaErroresValidacion.push(`El estado '${nuevoEstado}' no está autorizado para la entidad ${configuracionActiva.etiqueta}.`);
                            }
                        } else {
                            // Comportamiento Toggle (Interruptor binario inferido desde el Diccionario)
                            nuevoEstado = (estadoActual === configuracionActiva.estadoInicial) 
                                ? configuracionActiva.estadosBloqueantes[0] 
                                : configuracionActiva.estadoInicial;
                        }

                        // --- 5. MOTOR DE REGLAS: INTEGRIDAD TOP-DOWN (BATCHING I/O) ---
                        // Validamos reglas de dependencias solo si vamos hacia un estado bloqueante
                        if (configuracionActiva.estadosBloqueantes.includes(nuevoEstado)) {
                            const tablaReglasMaestra = hojaMaestrosWS.getTable("TablaReglas");
                            if (tablaReglasMaestra) {
                                const matrizReglasTodas: ValorCelda[][] = tablaReglasMaestra.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][];
                                const diccionarioTablasAuxiliares: { [nombreTabla: string]: { encabezados: string[], datos: ValorCelda[][] } } = {};
                                const reglasDependencias: ValorCelda[][] = [];

                                // Fase A: Precarga Batching
                                matrizReglasTodas.forEach((regla: ValorCelda[]) => {
                                    if (configuracionActiva.tabla.toUpperCase().includes(String(regla[0]).toUpperCase())) {
                                        const operador: string = String(regla[2]);
                                        if (operador === "DEPENDENCIAS_CERRADAS") {
                                            reglasDependencias.push(regla);
                                            const nombreTablaRequerida: string = String(regla[3]).split("[")[0];
                                            
                                            if (nombreTablaRequerida && !diccionarioTablasAuxiliares[nombreTablaRequerida]) {
                                                const tablaAuxiliar = workbook.getTable(nombreTablaRequerida);
                                                if (tablaAuxiliar) {
                                                    diccionarioTablasAuxiliares[nombreTablaRequerida] = {
                                                        encabezados: tablaAuxiliar.getHeaderRowRange().getValues()[0].map((h: ValorCelda) => String(h).toUpperCase().replace(/\s/g, "_")),
                                                        datos: tablaAuxiliar.getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][]
                                                    };
                                                }
                                            }
                                        }
                                    }
                                });

                                // Fase B: Evaluación Analítica en Memoria RAM
                                reglasDependencias.forEach((regla: ValorCelda[]) => {
                                    const referenciaRaw: string = String(regla[3]);
                                    const mensajeErrorRegla: string = String(regla[4]);
                                    
                                    const [nombreTablaRef, parteColumna] = referenciaRaw.split("[");
                                    const nombreColumnaRef: string = parteColumna.replace("]", "");
                                    const tablaEnMemoria = diccionarioTablasAuxiliares[nombreTablaRef];

                                    if (tablaEnMemoria) {
                                        const idxForanea: number = tablaEnMemoria.encabezados.indexOf(nombreColumnaRef);
                                        const idxEstadoRef: number = tablaEnMemoria.encabezados.indexOf("ESTADO");

                                        if (idxForanea !== -1 && idxEstadoRef !== -1) {
                                            // Determinar el estado inicial dinámicamente de la dependencia
                                            let estadoInicialDependencia = "ABIERTO";
                                            for (const key in CONFIGURACION_ENTIDADES) {
                                                if (CONFIGURACION_ENTIDADES[key].tabla === nombreTablaRef) {
                                                    estadoInicialDependencia = CONFIGURACION_ENTIDADES[key].estadoInicial;
                                                    break;
                                                }
                                            }

                                            const tieneDependenciaAbierta: boolean = tablaEnMemoria.datos.some((filaRef: ValorCelda[]) => {
                                                return String(filaRef[idxForanea]).toUpperCase() === idABuscar && String(filaRef[idxEstadoRef]).toUpperCase() === estadoInicialDependencia;
                                            });

                                            if (tieneDependenciaAbierta) {
                                                listaErroresValidacion.push(mensajeErrorRegla);
                                            }
                                        }
                                    }
                                });
                            }
                        }

                        // --- 6. DECISIÓN DE PERSISTENCIA Y COMMIT ATÓMICO ---
                        if (listaErroresValidacion.length === 0) {
                            tablaBaseDatos.getWorksheet().getProtection().unprotect(claveProteccion);
                            const idxEstado: number = encabezadosTabla.indexOf("ESTADO");
                            const idxAuditTrail: number = encabezadosTabla.indexOf("AUDIT_TRAIL");
                            const idxUsuario: number = encabezadosTabla.indexOf("USUARIO");

                            // Escritura Atómica
                            const filaAtomicaModificada: ValorCelda[] = [...matrizValoresDB[indiceFilaEncontrada]];
                            filaAtomicaModificada[idxEstado] = nuevoEstado;
                            if (idxUsuario !== -1) filaAtomicaModificada[idxUsuario] = usuarioIngresado;
                            if (idxAuditTrail !== -1) filaAtomicaModificada[idxAuditTrail] = new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
                            
                            tablaBaseDatos.getRangeBetweenHeaderAndTotal().getRow(indiceFilaEncontrada).setValues([filaAtomicaModificada]);

                            // Registro en Historial
                            tablaHistorial.getWorksheet().getProtection().unprotect(claveProteccion);
                            const filaRegistroHistorial: ValorCelda[] = (tablaHistorial.getHeaderRowRange().getValues()[0] as string[]).map((h: string) => {
                                const headCaps = h.toUpperCase();
                                if (headCaps === "ID_EVENTO") return (tablaHistorial!.getRowCount() === 0 ? 1 : Math.max(...(tablaHistorial!.getColumnByName("ID_EVENTO").getRangeBetweenHeaderAndTotal().getValues() as ValorCelda[][]).map((v: ValorCelda[]) => Number(v[0]))) + 1);
                                if (headCaps === nombreCampoPrimario) return idABuscar;
                                if (headCaps === "USUARIO") return usuarioIngresado;
                                if (headCaps === "MOTIVO") return motivoDeCambio.trim();
                                if (headCaps === "CAMBIOS") return `ESTADO: [${estadoActual}] -> [${nuevoEstado}]`;
                                if (headCaps === "FECHA_CAMBIO") return new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour12: false });
                                return "";
                            });
                            tablaHistorial.addRow(-1, filaRegistroHistorial);

                            const verboOperacion = nuevoEstado === configuracionActiva.estadoInicial ? "reabiert" : "modificad";
                            resultadoOperacion.message = `✅ ${configuracionActiva.etiqueta.charAt(0).toUpperCase() + configuracionActiva.etiqueta.slice(1)} #${idABuscar} ${verboOperacion}${configuracionActiva.genero} a [${nuevoEstado}] con éxito.`;
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

                            tablaIntegridad.getRangeBetweenHeaderAndTotal().setValues(matrizSeg);
                        }
                    }
                }
            }
            
            // Consolidación de Errores UX
            if (listaErroresValidacion.length > 0) {
                resultadoOperacion.success = false;
                resultadoOperacion.message = "⚠️ " + listaErroresValidacion.join(" | ");
                resultadoOperacion.logLevel = 'WARN';
            }
        }
    } catch (e) {
        resultadoOperacion.success = false;
        resultadoOperacion.logLevel = 'ERROR';
        resultadoOperacion.message = `❌ Fallo Crítico de Infraestructura: ${String(e)}`;
    } finally {
        // --- 7. PROTOCOLO DE CIERRE SEGURO ---
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
            try { hoja.getProtection().protect({ allowAutoFilter: true }, pass); }
            catch (e) { 
                res.success = false;
                res.logLevel = 'ERROR';
                res.message += ` | Falla protegiendo: ${hoja.getName()}`; 
            }
        }
    }

    function auxiliarLimpiarFormulario(hoja: ExcelScript.Worksheet, matriz: ValorCelda[][], filaInicio: number, campoId: string): void {
        matriz.forEach((fila: ValorCelda[], i: number) => {
            const claveCampo: string = String(fila[0]).trim().toUpperCase().replace("*", "").replace(/\s/g, "_");
            if (claveCampo !== "" && claveCampo !== campoId) {
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