/**
 * SCRIPT: UI_CONFIGURAR_RANGOS_ENTRADA (UNIVERSAL)
 * OBJETIVO: Protocolo dinámico de bloqueo/desbloqueo de celdas basado en el contrato de interfaz (Etiqueta en B -> Input en C).
 * GARANTÍA: Aisla la superficie de captura de datos, protegiendo fórmulas, encabezados y Primary Keys en cualquier hoja INP_.
 */
function main(
  workbook: ExcelScript.Workbook,
  nombreHojaEntrada: string = "INP_DES", // Valor por defecto, reemplazable al invocar por botón
  nombreHojaMaestros: string = "MAESTROS" 
) {
  // --- 1. CONFIGURACIÓN DE IDENTIDAD Y CONSTANTES ---
  const NOMBRE_ITEM_CLAVE: string = "SISTEMA_CLAVE"; 

  const hojaEntradaWS: ExcelScript.Worksheet | undefined = workbook.getWorksheet(nombreHojaEntrada);
  const hojaMaestrosWS: ExcelScript.Worksheet | undefined = workbook.getWorksheet(nombreHojaMaestros);
  
  let claveSeguridadSistema: string = "";
  let registroMensajeLog: string = "";
  let ejecucionHabilitada: boolean = true;

  // --- I. VALIDACIÓN DE ENTORNO Y CONTRATO DE INTERFAZ ---
  if (!hojaEntradaWS || !hojaMaestrosWS) {
    registroMensajeLog = `⛔ Error de Infraestructura: Faltan hojas críticas (${nombreHojaEntrada} o ${nombreHojaMaestros}).`;
    ejecucionHabilitada = false;
  } else if (!nombreHojaEntrada.startsWith("INP_")) {
    registroMensajeLog = `⛔ Error de Negocio: La hoja '${nombreHojaEntrada}' rechazada. El protocolo exige el prefijo 'INP_'.`;
    ejecucionHabilitada = false;
  } else {
    const rangoItemClave = workbook.getNamedItem(NOMBRE_ITEM_CLAVE)?.getRange();
    if (rangoItemClave) {
      claveSeguridadSistema = String(rangoItemClave.getText());
      if (claveSeguridadSistema === "") {
        registroMensajeLog = "⚠️ Advertencia: El protocolo de seguridad detectó una clave de sistema vacía.";
      }
    } else {
      registroMensajeLog = `⛔ Error de Seguridad: No se encontró el ítem de protección estandarizado '${NOMBRE_ITEM_CLAVE}'.`;
      ejecucionHabilitada = false;
    }
  }
  
  // --- II. EJECUCIÓN DEL MOTOR DE DESBLOQUEO (SESE & BATCHING I/O) ---
  if (ejecucionHabilitada && hojaEntradaWS) {
    try {
      // A) RESET DE SEGURIDAD (Zero Trust): Bloqueo integral
      hojaEntradaWS.getProtection().unprotect(claveSeguridadSistema);
      
      const rangoUsadoTotal = hojaEntradaWS.getUsedRange();
      if (rangoUsadoTotal) {
          rangoUsadoTotal.getFormat().getProtection().setLocked(true); 
      }

      // B) MAPEO EN RAM (Contrato Columna B -> C)
      let camposContabilizados: number = 0;
      const FILA_MINIMA_DESBLOQUEO: number = 4; // Regla de Negocio: De C4 hacia arriba queda bloqueado
      
      const rangoEtiquetasUsadas = hojaEntradaWS.getRange("B:B").getUsedRange();
      
      if (rangoEtiquetasUsadas) {
        const matrizEtiquetas: (string | number | boolean)[][] = rangoEtiquetasUsadas.getValues();
        const indiceFilaInicial: number = rangoEtiquetasUsadas.getRowIndex();
        
        // Recolector numérico de filas objetivo
        const filasAUnlockear: number[] = []; 

        matrizEtiquetas.forEach((fila: (string | number | boolean)[], indiceMatriz: number) => {
          const valorEtiqueta: string = String(fila[0]).trim();
          const filaExcelFisica: number = indiceFilaInicial + indiceMatriz + 1; 

          // Evaluación de Reglas de Negocio Estrictas
          const esEtiquetaValida = valorEtiqueta !== "";
          const estaFueraDeZonaProtegida = filaExcelFisica >= FILA_MINIMA_DESBLOQUEO;
          const noEsIdentificador = !valorEtiqueta.toUpperCase().startsWith("ID_");

          if (esEtiquetaValida && estaFueraDeZonaProtegida && noEsIdentificador) { 
            filasAUnlockear.push(filaExcelFisica);
            camposContabilizados++;
          }
        });

        // C) ALGORITMO DE AGRUPACIÓN (Contiguous Range Batching)
        // Transforma [5,6,7,9,10] en ["C5:C7", "C9:C10"]
        if (filasAUnlockear.length > 0) {
            let bloquesContiguos: string[] = [];
            let filaInicio = filasAUnlockear[0];
            let filaFin = filasAUnlockear[0];

            let indiceIteracion = 1;
            while (indiceIteracion < filasAUnlockear.length) {
                if (filasAUnlockear[indiceIteracion] === filaFin + 1) {
                    filaFin = filasAUnlockear[indiceIteracion];
                } else {
                    bloquesContiguos.push(filaInicio === filaFin ? `C${filaInicio}` : `C${filaInicio}:C${filaFin}`);
                    filaInicio = filasAUnlockear[indiceIteracion];
                    filaFin = filasAUnlockear[indiceIteracion];
                }
                indiceIteracion++;
            }
            bloquesContiguos.push(filaInicio === filaFin ? `C${filaInicio}` : `C${filaInicio}:C${filaFin}`);

            // COMMIT I/O: Peticiones mínimas con sintaxis estricta API
            bloquesContiguos.forEach((bloqueRango: string) => {
                hojaEntradaWS.getRange(bloqueRango).getFormat().getProtection().setLocked(false);
            });
        }
      }

      registroMensajeLog = camposContabilizados > 0 
        ? `✅ Configuración universal exitosa. ${camposContabilizados} inputs habilitados en ${nombreHojaEntrada}.` 
        : `⚠️ Proceso finalizado: No se detectaron etiquetas operativas habilitables en ${nombreHojaEntrada}.`;

    } catch (errorEjecucion) {
      registroMensajeLog = `❌ Error de Ejecución en ${nombreHojaEntrada}: ${(errorEjecucion as Error).message}`;
    } finally {
      // D) PROTOCOLO DE CIERRE SEGURO (Zero Silent Failures)
      try {
        hojaEntradaWS.getProtection().protect({ allowAutoFilter: true }, claveSeguridadSistema);
      } catch (errorProteccion) {
        registroMensajeLog += `\n⛔ CRÍTICO: Falla al reproteger ${nombreHojaEntrada}. Superficie de datos vulnerable.`;
      }
    }
  }

  // Auditoría en consola de administración
  console.log(registroMensajeLog);
}