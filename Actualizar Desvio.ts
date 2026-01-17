function main(workbook: ExcelScript.Workbook) {
    // --- BLOQUE 1: CONFIGURACIÓN INICIAL ---
    const hojaInput = workbook.getWorksheet("INPUT_DESVIOS");
    const hojaMaestros = workbook.getWorksheet("MAESTROS");
    let CLAVE_SEGURIDAD = "";

    // Validación de Dependencia
    if (hojaMaestros) {
        CLAVE_SEGURIDAD = hojaMaestros.getRange("XFD1").getText();
    } else {
        throw new Error("ERROR CRÍTICO: No se encuentra la hoja 'MAESTROS'.");
    }

    // Desproteger Input
    try {
        hojaInput.getProtection().unprotect(CLAVE_SEGURIDAD);
    } catch (e) { }

    // Configuración UI Mensaje (Merge & Clear)
    const celdaMensaje = hojaInput.getRange("E4:H6");
    celdaMensaje.merge(false);
    celdaMensaje.clear(ExcelScript.ClearApplyTo.contents);
    celdaMensaje.getFormat().getFill().clear();
    celdaMensaje.getFormat().getFont().setColor("Black");
    celdaMensaje.getFormat().getFont().setBold(false);
    celdaMensaje.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    celdaMensaje.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    celdaMensaje.getFormat().setWrapText(true);

    // Mapeo
    const mapaCeldas: { [key: string]: string } = {
        "Fecha Suceso": "C4",
        "Fecha Registro": "C5",
        "Fecha QA": "C6",
        "Planta": "C8",
        "Tercerista": "C10",
        "Descripción": "C12",
        "Etapa Ocurrencia": "C14",
        "Etapa Detección": "C16",
        "Clasificación": "C18",
        "Impacto": "C20",
        "Observaciones": "C22",
        "Usuario": "C24"
    };

    const celdaIdBuscar = "C2";
    const celdaMotivo = "C26";
    const celdaTestigo = "Z1";

    try {
        // --- BLOQUE 2: VALIDACIONES DE INTEGRIDAD ---

        // 1. Validar ID Visual vs. ID Testigo (Anti-Cambio)
        const idVisual = hojaInput.getRange(celdaIdBuscar).getValue() as number;
        const idTestigo = hojaInput.getRange(celdaTestigo).getValue() as number;

        if (!idVisual) throw new Error("Debe indicar un ID en la celda C2.");

        if (idVisual !== idTestigo) {
            throw new Error(`⛔ ERROR DE INTEGRIDAD:\nEl ID en pantalla (${idVisual}) no coincide con el ID buscado originalmente (${idTestigo}).\n\nPresione 'BUSCAR DESVÍO' para recargar los datos correctos.`);
        }

        // 2. Anti-Vacío (Anti-Borrado)
        const checkDescripcion = hojaInput.getRange(mapaCeldas["Descripción"]).getText();
        if (checkDescripcion === "") {
            throw new Error("⛔ PELIGRO:\nEl formulario parece estar vacío. Ejecute 'BUSCAR DESVÍO' antes de intentar actualizar.");
        }

        // Nota: NO validamos el motivo todavía. Primero vemos si vale la pena.

        // --- BLOQUE 3: COMPARACIÓN (DELTA LOGGING) ---

        const hojaBD = workbook.getWorksheet("BD_DESVIOS");
        if (!hojaBD) throw new Error("Falta BD_DESVIOS.");
        const tablaDesvios = hojaBD.getTable("TablaDesvios");
        if (!tablaDesvios) throw new Error("Falta TablaDesvios.");

        // Buscar Fila
        const columnaID = tablaDesvios.getColumnByName("ID");
        const valoresID = columnaID.getRangeBetweenHeaderAndTotal().getValues();
        let indiceFila = -1;
        for (let i = 0; i < valoresID.length; i++) {
            if (valoresID[i][0] == idVisual) {
                indiceFila = i;
                break;
            }
        }
        if (indiceFila === -1) throw new Error(`El Desvío #${idVisual} no existe en BD.`);

        // Comparar
        const rangoEncabezados = tablaDesvios.getHeaderRowRange();
        const encabezados = rangoEncabezados.getValues()[0] as string[];
        const filaVieja = tablaDesvios.getRangeBetweenHeaderAndTotal().getRow(indiceFila).getValues()[0];

        let listaCambios: string[] = [];
        let nuevaFilaValores: (string | number | boolean)[] = [];
        let huboCambios = false;

        for (let i = 0; i < encabezados.length; i++) {
            let nombreCol = encabezados[i];
            let valorViejo = filaVieja[i];
            let valorNuevo = valorViejo;

            if (mapaCeldas[nombreCol]) {
                let valorInput = hojaInput.getRange(mapaCeldas[nombreCol]).getValue();
                if (valorInput.toString() != valorViejo.toString()) {
                    listaCambios.push(`[${nombreCol}: ${valorViejo} -> ${valorInput}]`);
                    valorNuevo = valorInput;
                    huboCambios = true;
                } else {
                    valorNuevo = valorInput;
                }
            }
            if (nombreCol === "Audit Trail" && huboCambios) valorNuevo = new Date().toLocaleString();
            nuevaFilaValores.push(valorNuevo);
        }

        // --- BLOQUE DE DECISIÓN: ¿HUBO CAMBIOS? ---

        if (!huboCambios) {
            // CASO A: NO CAMBIÓ NADA
            // Avisamos amablemente y terminamos. No pedimos motivo.
            celdaMensaje.setValue("ℹ️ SIN CAMBIOS:\nLos datos en el formulario son idénticos a los de la base de datos. No se realizó ninguna acción.");
            celdaMensaje.getFormat().getFill().setColor("#F2F2F2"); // Gris claro
            celdaMensaje.getFormat().getFont().setColor("#595959"); // Gris oscuro

        } else {
            // CASO B: SÍ HUBO CAMBIOS
            // Ahora sí, VALIDAMOS EL MOTIVO (Obligatorio GMP)
            const motivoCambio = hojaInput.getRange(celdaMotivo).getText();

            if (motivoCambio === "") {
                throw new Error("⚠️ MOTIVO REQUERIDO:\nSe han detectado cambios en los datos.\nEs obligatorio indicar el 'Motivo del Cambio' para proceder.");
            }

            // --- BLOQUE 4: ESCRITURA BLINDADA ---
            const hojaHistorial = workbook.getWorksheet("HISTORIAL_DESVIOS");
            if (!hojaHistorial) throw new Error("Falta HISTORIAL_DESVIOS.");

            // 🔓 Desproteger BDs
            hojaBD.getProtection().unprotect(CLAVE_SEGURIDAD);
            hojaHistorial.getProtection().unprotect(CLAVE_SEGURIDAD);

            // 1. Update BD
            tablaDesvios.getRangeBetweenHeaderAndTotal().getRow(indiceFila).setValues([nuevaFilaValores]);

            // 2. Insert Historial
            const tablaHistorial = hojaHistorial.getTable("TablaHistorial");
            let idEvento = 1;
            if (tablaHistorial.getRowCount() > 0) {
                let idsH = tablaHistorial.getColumnByName("ID_EVENTO").getRangeBetweenHeaderAndTotal().getValues();
                let maxH = Math.max(...idsH.map(f => Number(f[0])));
                idEvento = maxH + 1;
            }

            let filaHistorial = [
                idEvento, idVisual, new Date().toLocaleString(),
                hojaInput.getRange(mapaCeldas["Usuario"]).getText(),
                motivoCambio, listaCambios.join("; ")
            ];
            tablaHistorial.addRow(-1, filaHistorial);

            // Formato
            tablaDesvios.getRange().getFormat().autofitColumns();
            tablaHistorial.getRange().getFormat().autofitColumns();

            // 🔒 Reproteger BDs
            const opts = { allowInsertRows: false, allowDeleteRows: false, allowFormatCells: false, allowAutoFilter: true, allowSort: true };
            hojaBD.getProtection().protect(opts, CLAVE_SEGURIDAD);
            hojaHistorial.getProtection().protect(opts, CLAVE_SEGURIDAD);

            // Limpieza
            hojaInput.getRange("C4:C30").clear(ExcelScript.ClearApplyTo.contents);
            hojaInput.getRange(celdaIdBuscar).clear(ExcelScript.ClearApplyTo.contents);
            hojaInput.getRange(celdaTestigo).clear(ExcelScript.ClearApplyTo.contents);

            celdaMensaje.setValue(`✅ Actualización Exitosa.\nCambios guardados: ${listaCambios.length}`);
            celdaMensaje.getFormat().getFill().setColor("#DFF6DD");
            celdaMensaje.getFormat().getFont().setColor("#006600");
            celdaMensaje.getFormat().getFont().setBold(true);
        }

    } catch (error) {
        console.log(error);
        // Fail-safe reprotection
        try { workbook.getWorksheet("BD_DESVIOS").getProtection().protect({ allowAutoFilter: true }, CLAVE_SEGURIDAD); } catch (e) { }
        try { workbook.getWorksheet("HISTORIAL_DESVIOS").getProtection().protect({ allowAutoFilter: true }, CLAVE_SEGURIDAD); } catch (e) { }

        celdaMensaje.setValue("⛔ ERROR:\n" + error.message);
        celdaMensaje.getFormat().getFill().setColor("#FFDDDD");
        celdaMensaje.getFormat().getFont().setColor("Red");

    } finally {
        // CIERRE FINAL: Reproteger Input
        try {
            hojaInput.getProtection().protect({
                allowSelectLockedCells: true, allowSelectUnlockedCells: true, allowAutoFilter: false
            }, CLAVE_SEGURIDAD);
        } catch (e) { }
    }
}