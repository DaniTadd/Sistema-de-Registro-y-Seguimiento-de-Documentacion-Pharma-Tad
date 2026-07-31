/**
 * SCRIPT: UI_BUSCAR_UNIVERSAL
 * OBJETIVO: Recuperar un registro único mediante Query-By-Example (QBE) con feedback visual de estado (UX).
 * GARANTÍA: Estructura SESE estricta. Actualización de UI y sellado de seguridad garantizado en bloque finally.
 */

// --- INTERFACES Y TIPOS ---
type NivelLog = 'INFO' | 'SUCCESS' | 'WARNING' | 'ERROR';

interface ResultadoAccion {
    success: boolean;
    message: string;
    logLevel: NivelLog;
}

interface MapaColoresUX {
    [key: string]: { fondo: string; texto: string };
}

function main(workbook: ExcelScript.Workbook) {
    // --- 1. CONFIGURACIÓN DE IDENTIDAD Y CONSTANTES ---
    const hojaActivaWS: ExcelScript.Worksheet = workbook.getActiveWorksheet();
    const nombreHojaActiva: string = hojaActivaWS.getName().trim();
    const NOMBRE_ITEM_CLAVE: string = "SISTEMA_CLAVE";

    const CONFIGURACION_ENTIDADES: { [key: string]: { tabla: string } } = {
        "INP_DES": { tabla: "TablaDesvios" },
        "INP_CAPAS": { tabla: "TablaCapas" },
        "INP_AFECT": { tabla: "TablaAfectacion" },
        "INP_EQ": { tabla: "TablaEquipos" }
    };

    const PALETA_COLORES_UX: MapaColoresUX = {
        'INFO': { fondo: "#D9E1F2", texto: "#1F4E78" },     // Azul tenue
        'SUCCESS': { fondo: "#E2EFDA", texto: "#375623" },  // Verde
        'WARNING': { fondo: "#FFF2CC", texto: "#7F6000" },  // Amarillo
        'ERROR': { fondo: "#FCE4D6", texto: "#C00000" }     // Rojo
    };

    let resultadoOperacion: ResultadoAccion = { success: true, message: "Inicio de transacción de Búsqueda QBE.", logLevel: 'INFO' };
    let ejecucionHabilitada: boolean = true;
    let claveSeguridadSistema: string = "";

    // --- I. VALIDACIÓN DE ENTORNO E INFRAESTRUCTURA ---
    if (!nombreHojaActiva.startsWith("INP_") || !CONFIGURACION_ENTIDADES[nombreHojaActiva]) {
        resultadoOperacion = { success: false, message: `Error: La hoja '${nombreHojaActiva}' no es un formulario válido.`, logLevel: 'ERROR' };
        ejecucionHabilitada = false;
    } else {
        const rangoItemClave = workbook.getNamedItem(NOMBRE_ITEM_CLAVE)?.getRange();
        if (rangoItemClave) {
            claveSeguridadSistema = String(rangoItemClave.getText()).trim();
        } else {
            resultadoOperacion = { success: false, message: `Error de Seguridad: Falta credencial '${NOMBRE_ITEM_CLAVE}'.`, logLevel: 'ERROR' };
            ejecucionHabilitada = false;
        }
    }

    // --- II. EJECUCIÓN DEL MOTOR DE BÚSQUEDA (SESE) ---
    if (ejecucionHabilitada && hojaActivaWS) {
        try {
            const configActiva = CONFIGURACION_ENTIDADES[nombreHojaActiva];
            const tablaObjetivo: ExcelScript.Table | undefined = workbook.getTable(configActiva.tabla);

            if (!tablaObjetivo) {
                throw new Error(`No se encontró la BBDD '${configActiva.tabla}'.`);
            }

            const rangoUsado = hojaActivaWS.getRange("B:C").getUsedRange();
            
            if (rangoUsado) {
                const matrizInterfaz: (string | number | boolean)[][] = rangoUsado.getValues();
                const numFilasEncabezado: number = 1;

                let etiquetaCriterioBusqueda: string = "";
                let valorCriterioBusqueda: string = "";
                let camposPobladosCount: number = 0;

                let idxInterfaz: number = numFilasEncabezado;
                while (idxInterfaz < matrizInterfaz.length) {
                    const etiquetaRaw: string = String(matrizInterfaz[idxInterfaz][0]).trim();
                    const valorRaw: string = String(matrizInterfaz[idxInterfaz][1]).trim();

                    if (etiquetaRaw !== "" && valorRaw !== "" && valorRaw !== "N/A") {
                        etiquetaCriterioBusqueda = etiquetaRaw.replace(/\*/g, "").toUpperCase().replace(/\s/g, "_");
                        valorCriterioBusqueda = valorRaw.toUpperCase();
                        camposPobladosCount++;
                    }
                    idxInterfaz++;
                }

                if (camposPobladosCount === 0) {
                    resultadoOperacion = { success: false, message: "Validación: Formulario vacío. Ingrese un valor para buscar.", logLevel: 'WARNING' };
                } else if (camposPobladosCount > 1) {
                    resultadoOperacion = { success: false, message: "Validación: Riesgo de ambigüedad. Utilice solo UN campo para buscar.", logLevel: 'WARNING' };
                } else {
                    
                    const encabezadosBD: string[] = tablaObjetivo.getHeaderRowRange().getValues()[0].map(h => String(h).trim().toUpperCase());
                    const datosBD: (string | number | boolean)[][] = tablaObjetivo.getRangeBetweenHeaderAndTotal().getValues();
                    const indiceColumnaBusqueda: number = encabezadosBD.indexOf(etiquetaCriterioBusqueda);

                    if (indiceColumnaBusqueda === -1) {
                        resultadoOperacion = { success: false, message: `Error Estructural: El campo '${etiquetaCriterioBusqueda}' no existe en la BBDD.`, logLevel: 'ERROR' };
                    } else {
                        let registrosCoincidentes: (string | number | boolean)[][] = [];
                        let idxBD: number = 0;

                        while (idxBD < datosBD.length) {
                            const valorFilaActual: string = String(datosBD[idxBD][indiceColumnaBusqueda]).trim().toUpperCase();
                            if (valorFilaActual === valorCriterioBusqueda) {
                                registrosCoincidentes.push(datosBD[idxBD]);
                            }
                            idxBD++;
                        }

                        if (registrosCoincidentes.length === 0) {
                            resultadoOperacion = { success: false, message: `Búsqueda sin resultados para '${valorCriterioBusqueda}'.`, logLevel: 'WARNING' };
                        } else if (registrosCoincidentes.length > 1) {
                            resultadoOperacion = { success: false, message: `Bloqueo ALCOA+: Se hallaron ${registrosCoincidentes.length} registros. El criterio debe ser único.`, logLevel: 'ERROR' };
                        } else {
                            
                            const filaRecuperada = registrosCoincidentes[0];
                            const direccionesEscritura: string[] = [];
                            const valoresEscritura: (string | number | boolean)[][] = [];

                            let idxRestauracion: number = numFilasEncabezado;
                            while (idxRestauracion < matrizInterfaz.length) {
                                const etiquetaRaw: string = String(matrizInterfaz[idxRestauracion][0]).trim();
                                if (etiquetaRaw !== "") {
                                    const etiquetaNormalizada: string = etiquetaRaw.replace(/\*/g, "").toUpperCase().replace(/\s/g, "_");
                                    const indiceEnBD: number = encabezadosBD.indexOf(etiquetaNormalizada);
                                    
                                    if (indiceEnBD !== -1) {
                                        const valorARestaurar = filaRecuperada[indiceEnBD];
                                        const filaExcelFisica: number = rangoUsado.getRowIndex() + idxRestauracion + 1;
                                        direccionesEscritura.push(`C${filaExcelFisica}`);
                                        valoresEscritura.push([valorARestaurar !== "" ? valorARestaurar : "N/A"]);
                                    }
                                }
                                idxRestauracion++;
                            }

                            if (direccionesEscritura.length > 0) {
                                hojaActivaWS.getProtection().unprotect(claveSeguridadSistema);
                                
                                for (let i = 0; i < direccionesEscritura.length; i++) {
                                    hojaActivaWS.getRange(direccionesEscritura[i]).setValues([valoresEscritura[i]]);
                                }
                                resultadoOperacion = { success: true, message: `Registro recuperado con éxito.`, logLevel: 'SUCCESS' };
                            }
                        }
                    }
                }
            } else {
                throw new Error("No se pudo leer el rango de la interfaz.");
            }
        } catch (errorEjecucion) {
            resultadoOperacion = { success: false, message: `Falla crítica: ${(errorEjecucion as Error).message}`, logLevel: 'ERROR' };
        } finally {
            // Cierre estandarizado: Notificación visual al analista y re-sellado
            auxiliarActualizarInterfazUX(hojaActivaWS, resultadoOperacion, PALETA_COLORES_UX, claveSeguridadSistema);
            auxiliarProtegerHoja(hojaActivaWS, claveSeguridadSistema);
            console.log(`[SYS_AUDIT] ${resultadoOperacion.message}`);
        }
    } else {
        console.log(`[SYS_ABORT] ${resultadoOperacion.message}`);
    }

    // --- FUNCIONES AUXILIARES (SESE COMPLIANT) ---

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
                console.log("Falla al aplicar UX. La hoja podría estar bloqueada con otra clave: ", e);
            }
        } else {
            console.log("Falla de Infraestructura: No se encontró el ítem nombrado 'UI_FEEDBACK'.");
        }
    }

    function auxiliarProtegerHoja(hojaAProteger: ExcelScript.Worksheet, clave: string): void {
        if (hojaAProteger) {
            try {
                hojaAProteger.getProtection().protect({ allowAutoFilter: true }, clave);
            } catch (e) {
                console.log(`⛔ CRÍTICO: Falla al reproteger hoja activa. Superficie vulnerable.`);
            }
        }
    }
}