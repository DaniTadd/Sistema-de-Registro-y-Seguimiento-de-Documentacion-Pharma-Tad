function main(workbook: ExcelScript.Workbook) {
    const hoja = workbook.getWorksheet("BD_DESVIOS");

    // 1. Primero desprotegemos por si acaso (ajusta pass si tienes)
    // hoja.getProtection().unprotect("1234"); 

    // 2. Definimos tus rangos editables
    const rangosADesbloquear = [""C2", "C4:C6", "C8", "C10", "C12", "C14", "C16","C18", "C20", "C22", "C24", "C26""];

    // 3. Recorremos y desbloqueamos
    rangosADesbloquear.forEach(direccion => {
        let rango = hoja.getRange(direccion);
        let formato = rango.getFormat();

        // Aquí está la magia: false = Desbloqueado (Editable)
        formato.getProtection().setLocked(false);
    });

    console.log("Celdas desbloqueadas correctamente.");
}