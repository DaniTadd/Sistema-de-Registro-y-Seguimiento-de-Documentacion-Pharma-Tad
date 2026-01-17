function main(workbook: ExcelScript.Workbook) {
    // --- BLOQUE 1: CONFIGURACIÓN INICIAL ---
    const hojaInput = workbook.getWorksheet("INPUT_DESVIOS");
    const hojaMaestros = workbook.getWorksheet("MAESTROS");

    // Variable para la clave. La inicializamos vacía.
    let CLAVE_SEGURIDAD = "";

    // VALIDACIÓN DE DEPENDENCIA: ¿Existe la hoja de configuración?
    if (hojaMaestros) {
        // Si existe, leemos la clave
        CLAVE_SEGURIDAD = hojaMaestros.getRange("XFD1").getText();
    } else {
        // Si no existe, es un error crítico de infraestructura.
        // No usamos return. Lanzamos una excepción para avisar al sistema.
        throw new Error("ERROR CRÍTICO: No se encuentra la hoja 'MAESTROS' necesaria para la seguridad.");
    }

    // --- BLOQUE 2: PREPARACIÓN DE LA HOJA (UI) ---

    // Intentamos desproteger el Input usando la clave recuperada
    try {
        hojaInput.getProtection().unprotect(CLAVE_SEGURIDAD);
    } catch (e) {
        // Si la clave falla o ya está desprotegida, continuamos (no bloqueante para lectura)
    }

    // Configuración del Mensaje (Merge & Clear)
    const celdaMensaje = hojaInput.getRange("E4:H6");
    celdaMensaje.merge(false);
    celdaMensaje.clear(ExcelScript.ClearApplyTo.contents);
    celdaMensaje.getFormat().getFill().clear();
    celdaMensaje.getFormat().getFont().setColor("Black");
    celdaMensaje.getFormat().getFont().setBold(false);
    celdaMensaje.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    celdaMensaje.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    celdaMensaje.getFormat().setWrapText(true);

    // Limpieza de campos y testigo
    hojaInput.getRange("C4:C24").clear(ExcelScript.ClearApplyTo.contents);
    hojaInput.getRange("Z1").clear(ExcelScript.ClearApplyTo.contents);

    try {
        // --- BLOQUE 3: BÚSQUEDA ---

        // 1. Validar ID ingresado
        const idBuscado = hojaInput.getRange("C2").getValue() as number;

        if (!idBuscado) {
            // Error de Usuario
            celdaMensaje.setValue("⚠️ Por favor, ingrese un ID numérico en la celda C2.");
            celdaMensaje.getFormat().getFill().setColor("#FFFFCC");

        } else {
            // Si hay ID, procedemos a buscar
            const hojaBD = workbook.getWorksheet("BD_DESVIOS");

            if (hojaBD) {
                const tablaDesvios = hojaBD.getTable("TablaDesvios");

                if (tablaDesvios) {
                    const columnaID = tablaDesvios.getColumnByName("ID");
                    const valoresID = columnaID.getRangeBetweenHeaderAndTotal().getValues();

                    let indiceFila = -1;
                    // Bucle de búsqueda
                    for (let i = 0; i < valoresID.length; i++) {
                        if (valoresID[i][0] == idBuscado) {
                            indiceFila = i;
                            break;
                        }
                    }

                    if (indiceFila !== -1) {
                        // --- BLOQUE 4: LECTURA Y CARGA (Si encontramos la fila) ---

                        // Obtenemos datos y encabezados
                        const filaDatos = tablaDesvios.getRangeBetweenHeaderAndTotal().getRow(indiceFila).getValues()[0];
                        const encabezados = tablaDesvios.getHeaderRowRange().getValues()[0] as string[];

                        // Helper para mapear por nombre
                        const getValor = (nombreColumna: string) => {
                            const index = encabezados.indexOf(nombreColumna);
                            // Operador ternario: Si existe el índice devuelve el dato, sino vacío.
                            return index > -1 ? filaDatos[index] : "";
                        };

                        // Mapeo a celdas del formulario
                        hojaInput.getRange("C4").setValue(getValor("Fecha Suceso"));
                        hojaInput.getRange("C5").setValue(getValor("Fecha Registro"));
                        hojaInput.getRange("C6").setValue(getValor("Fecha QA"));
                        hojaInput.getRange("C8").setValue(getValor("Planta"));
                        hojaInput.getRange("C10").setValue(getValor("Tercerista"));
                        hojaInput.getRange("C12").setValue(getValor("Descripción"));
                        hojaInput.getRange("C14").setValue(getValor("Etapa Ocurrencia"));
                        hojaInput.getRange("C16").setValue(getValor("Etapa Detección"));
                        hojaInput.getRange("C18").setValue(getValor("Clasificación"));
                        hojaInput.getRange("C20").setValue(getValor("Impacto"));
                        hojaInput.getRange("C22").setValue(getValor("Observaciones"));
                        hojaInput.getRange("C24").setValue(getValor("Usuario"));

                        // Limpiamos motivo anterior
                        hojaInput.getRange("C26").clear(ExcelScript.ClearApplyTo.contents);

                        // 🔑 TESTIGO: Guardamos el ID encontrado en Z1 (oculta)
                        hojaInput.getRange("Z1").setValue(idBuscado);

                        celdaMensaje.setValue(`✅ Desvío #${idBuscado} cargado. Listo para editar.`);
                        celdaMensaje.getFormat().getFill().setColor("#DFF6DD");
                        celdaMensaje.getFormat().getFont().setColor("#006600");
                        celdaMensaje.getFormat().getFont().setBold(true);

                    } else {
                        // ID no encontrado en la Tabla
                        celdaMensaje.setValue(`⛔ No se encontró el Desvío #${idBuscado} en la base de datos.`);
                        celdaMensaje.getFormat().getFill().setColor("#FFDDDD");
                        celdaMensaje.getFormat().getFont().setColor("Red");
                    }
                } else {
                    throw new Error("Falta la TablaDesvios.");
                }
            } else {
                throw new Error("Falta la hoja BD_DESVIOS.");
            }
        }

    } catch (error) {
        // Manejo de errores técnicos (Nombre de hojas, tablas, etc.)
        console.log(error);
        celdaMensaje.setValue("ERROR DE SISTEMA:\n" + error.message);
        celdaMensaje.getFormat().getFill().setColor("#FFDDDD");
        celdaMensaje.getFormat().getFont().setColor("Red");

    } finally {
        // --- CIERRE FINAL: RE-PROTEGER INPUT ---
        // Esto se ejecuta SIEMPRE, haya éxito o error.
        try {
            hojaInput.getProtection().protect({
                allowSelectLockedCells: true,
                allowSelectUnlockedCells: true,
                allowAutoFilter: false
            }, CLAVE_SEGURIDAD);
        } catch (e) {
            // Si falla la reprotección (raro), no podemos hacer mucho más.
            console.log("No se pudo reproteger Input.");
        }
    }
}