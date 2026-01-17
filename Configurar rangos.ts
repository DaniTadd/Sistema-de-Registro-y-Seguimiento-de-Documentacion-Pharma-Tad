function main(workbook: ExcelScript.Workbook) {
  // 1. CONFIGURACIÓN (Contexto V5)
  const SHEET_INPUT = "INPUT_DESVIOS";
  const SHEET_MAESTROS = "MAESTROS";
  const CELL_CLAVE = "XFD1";

  const wsInput = workbook.getWorksheet(SHEET_INPUT)!;
  const wsMaestros = workbook.getWorksheet(SHEET_MAESTROS)!;

  if (!wsInput) {
    console.log("Error: No se encuentra la hoja INPUT_DESVIOS");
    return;
  }

  // 2. OBTENER CLAVE Y DESPROTEGER (Para poder cambiar formatos)
  let clave = "";
  try {
    if (wsMaestros) {
      clave = wsMaestros.getRange(CELL_CLAVE).getValue() as string;
    }
    // Usamos la API corregida
    wsInput.getProtection().unprotect(clave);
    console.log("Hoja desprotegida temporalmente para configuración.");
  } catch (e) {
    console.log("La hoja ya estaba desprotegida o hubo un error de clave.");
  }

  // 3. DEFINIR TODOS LOS RANGOS DE INTERACCIÓN (V5)
  // Incluye Inputs Principales, Motivo y los nuevos módulos de la derecha
  const rangosInteractivos = [
    // A) Cabecera y Buscador
    "C2",           // ID Principal (Buscador)
    
    // B) Fechas
    "C6:C8",        // Suceso, Registro, QA
    
    // C) Selectores
    "C10", "C12",   // Planta, Tercerista
    
    // D) Textos Largos y Detalles
    "C14",          // Descripción
    "C16", "C18",   // Etapas
    "C20", "C22",   // Clasificación, Impacto
    "C24", "C26",   // Observaciones, Autor
    "C28",          // MOTIVO (Crítico para Actualizar)

    // E) NUEVO: Módulo Afectaciones (Lotes) - Columna F Arriba
    "F6:F12",       // Tipo, Orden, Material, Lote, Cantidad, Unidad, Disposición

    // F) NUEVO: Módulo Tareas (Acciones) - Columna F Abajo
    "F16:F18"       // Tarea, Responsable, Fecha Límite
  ];

  // 4. APLICAR DESBLOQUEO (Locked = false)
  // Primero, por seguridad, bloqueamos TODA la hoja (reset)
  wsInput.getRange().getFormat().getProtection().setLocked(true);
  
  // Luego desbloqueamos solo lo necesario
  rangosInteractivos.forEach(direccion => {
    // Obtenemos el rango y cambiamos su propiedad de protección
    wsInput.getRange(direccion).getFormat().getProtection().setLocked(false);
    console.log(`Rango ${direccion} desbloqueado.`);
  });

  // 5. REPROTECCIÓN FINAL (Modo Usuario)
  // Dejamos la hoja protegida, pero permitiendo seleccionar celdas desbloqueadas
  wsInput.getProtection().protect({
    allowSelectLockedCells: true,   // Permitir clic en celdas bloqueadas (útil para copiar)
    allowSelectUnlockedCells: true, // Permitir escribir en las desbloqueadas
    allowFormatCells: false,        // No dejar cambiar colores
    allowDeleteRows: false,
    allowInsertRows: false
  }, clave);

  console.log("✅ Configuración completada: INPUT_DESVIOS lista para el usuario.");
}