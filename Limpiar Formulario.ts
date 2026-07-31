/**
 * SCRIPT: UI_LIMPIAR_FORMULARIO
 * OBJETIVO: Purgar los valores ingresados en el formulario activo (Columna C) para preparar una nueva transacción o búsqueda.
 * GARANTÍA: Operación SESE, preserva formatos/validaciones (solo borra contenido) y sella la seguridad al finalizar.
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

    const PALETA_COLORES_UX: MapaColoresUX = {
        'INFO': { fondo: "#D9E1F2", texto: "#1F4E78" },     // Azul tenue
        'SUCCESS': { fondo: "#E2EFDA", texto: "#375623" },  // Verde
        'WARNING': { fondo: "#FFF2CC", texto: "#7F6000" },  // Amarillo
        'ERROR': { fondo: "#FCE4D6", texto: "#C00000" }     // Rojo
    };

    let resultadoOperacion: ResultadoAccion = { success: true, message: "Inicio de rutina de limpieza.", logLevel: 'INFO' };
    let ejecucionHabilitada: boolean = true;
    let claveSeguridadSistema: string = "";

    // --- I. VALIDACIÓN DE ENTORNO E INFRAESTRUCTURA ---
    if (!nombreHojaActiva.startsWith("INP_")) {
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

    // --- II. EJECUCIÓN DEL MOTOR DE LIMPIEZA (SESE) ---
    if (ejecucionHabilitada && hojaActivaWS) {
        try {
            const rangoUsado = hojaActivaWS.getUsedRange();
            
            if (rangoUsado) {
                // Cálculo de la última fila real para no barrer un millón de filas en vano
                const ultimaFilaFisica: number = rangoUsado.getLastCell().getRowIndex() + 1; // +1 por notación base 1 de Excel
                
                if (ultimaFilaFisica >= 2) {
                    // Abstracción del rango objetivo: Columna C (Datos), desde fila 2 (omitiendo encabezado)
                    const rangoLimpieza = hojaActivaWS.getRange(`C2:C${ultimaFilaFisica}`);
                    
                    // Apertura de Bóveda
                    hojaActivaWS.getProtection().unprotect(claveSeguridadSistema);
                    
                    // Mutación destructiva controlada (Solo contenido, preserva formato y validación de datos)
                    rangoLimpieza.clear(ExcelScript.ClearApplyTo.contents);
                    
                    resultadoOperacion = { success: true, message: "Formulario limpio y listo.", logLevel: 'SUCCESS' };
                } else {
                    resultadoOperacion = { success: true, message: "El formulario ya estaba vacío.", logLevel: 'INFO' };
                }
            } else {
                resultadoOperacion = { success: true, message: "El formulario está vacío.", logLevel: 'INFO' };
            }

        } catch (errorEjecucion) {
            resultadoOperacion = { success: false, message: `Falla crítica de infraestructura: ${(errorEjecucion as Error).message}`, logLevel: 'ERROR' };
        } finally {
            // Cierre estandarizado: Notificación visual al analista y re-sellado garantizado
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
            console.log("Falla de Infraestructura evitada: No se encontró el ítem nombrado 'UI_FEEDBACK'.");
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