/**
 * SCRIPT: SISTEMA_PURGAR_OUTBOX_BATCH
 * OBJETIVO: Microservicio idempotente con preservación de Viewport (evita saltos de hoja).
 */
interface ResultadoPurgado { success: boolean; message: string; }

async function main(
    workbook: ExcelScript.Workbook,
    idsMensajesCsv: string
): Promise<ResultadoPurgado> {
    const resultado: ResultadoPurgado = { success: true, message: "Inicio de purgado batch." };
    let hojaOutbox: ExcelScript.Worksheet | undefined;
    let claveProteccion: string = "";
    
    // 1. Captura anticipada de la hoja activa (Viewport State Preservation)
    const hojaActivaOriginal: ExcelScript.Worksheet = workbook.getActiveWorksheet();

    try {
        if (idsMensajesCsv === undefined || idsMensajesCsv === null || idsMensajesCsv.trim() === "") {
            throw new Error("Contrato API violado: Se requiere una cadena separada por comas con los IDs a purgar.");
        }

        const itemClaveSistema = workbook.getNamedItem("SISTEMA_CLAVE");
        const tablaOutbox = workbook.getTable("TablaNotificaciones_Outbox");

        if (!itemClaveSistema || !tablaOutbox) {
            throw new Error("Fallo de infraestructura: Tabla de Outbox o Clave de Sistema no encontrada.");
        }

        claveProteccion = String(itemClaveSistema.getRange().getValue());
        hojaOutbox = tablaOutbox.getWorksheet();
        
        // 2. Apertura Atómica (Operación en background)
        hojaOutbox.getProtection().unprotect(claveProteccion);

        // 3. Extracción y Normalización de Datos (Zero Trust Sanitization)
        const patronMSG = /MSG-[A-Za-z0-9\-]+/g;
        const extraccionCruda = idsMensajesCsv.match(patronMSG) || [];
        const listaIdsRequeridos = Array.from(new Set(extraccionCruda));
        
        if (listaIdsRequeridos.length === 0) {
            throw new Error("Validación de entrada: No se detectaron patrones de ID válidos (MSG-XXX) en el payload.");
        }

        const matrizOutbox = tablaOutbox.getRangeBetweenHeaderAndTotal().getValues() as string[][];

        // 4. Búsqueda Lineal en RAM (SESE Compliance)
        const indicesAEliminar: number[] = [];
        let indiceBusqueda = 0;

        while (indiceBusqueda < matrizOutbox.length) {
            const idFilaActiva = String(matrizOutbox[indiceBusqueda][0]).trim();
            let indiceComparacion = 0;
            let coincidenciaEncontrada = false;

            while (indiceComparacion < listaIdsRequeridos.length && !coincidenciaEncontrada) {
                if (idFilaActiva === listaIdsRequeridos[indiceComparacion]) {
                    indicesAEliminar.push(indiceBusqueda);
                    coincidenciaEncontrada = true;
                }
                indiceComparacion++;
            }
            indiceBusqueda++;
        }

        // 5. Ejecución de Borrado Batch (De abajo hacia arriba)
        if (indicesAEliminar.length > 0) {
            const indicesOrdenadosDesc = indicesAEliminar.sort((a, b) => b - a);
            let indiceBorrado = 0;
            
            while (indiceBorrado < indicesOrdenadosDesc.length) {
                tablaOutbox.getRangeBetweenHeaderAndTotal().getRow(indicesOrdenadosDesc[indiceBorrado]).delete(ExcelScript.DeleteShiftDirection.up);
                indiceBorrado++;
            }
            resultado.message = `Purgado batch exitoso. Registros eliminados: ${indicesOrdenadosDesc.length}.`;
        } else {
            resultado.success = false;
            resultado.message = `Idempotencia: Ninguno de los IDs provistos fue encontrado en la cola de transacciones.`;
        }

    } catch (e) {
        resultado.success = false;
        resultado.message = `Error crítico en purgado: ${String(e)}`;
    } finally {
        // 6. Re-sellado Hermético
        if (hojaOutbox) {
            try {
                hojaOutbox.getProtection().protect({ allowAutoFilter: true }, claveProteccion);
            } catch (errorProteccion) {
                resultado.success = false;
                resultado.message += " | ADVERTENCIA: Fallo al re-proteger la hoja Outbox.";
            }
        }

        // 7. Restauración Garantizada del Viewport Original
        try {
            if (hojaActivaOriginal) {
                hojaActivaOriginal.activate();
            }
        } catch (errorViewport) {
            // Silencioso por seguridad de infraestructura si el contexto visual expira
        }
    }

    return resultado;
}