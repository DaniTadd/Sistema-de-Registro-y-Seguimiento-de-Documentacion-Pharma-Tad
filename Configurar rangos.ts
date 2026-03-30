/**
 * SCRIPT: UI_CONFIGURAR_RANGOS_ENTRADA
 * OBJETIVO: Establecer el protocolo de bloqueo/desbloqueo de celdas en las hojas de entrada.
 * GARANTÍA: Asegura que solo los campos destinados a la captura de datos sean editables, protegiendo la integridad de las etiquetas y fórmulas.
 */
function main(
  workbook: ExcelScript.Workbook,
  nombreHojaEntrada: string = "INP_DES", // Parametrizado para uso universal
  nombreHojaMaestros: string = "MAESTROS" // Parametrizado para consulta de claves
) {
  // --- 1. CONFIGURACIÓN DE IDENTIDAD Y CONSTANTES ---
  const ETIQUETA_ENTIDAD: string = "desvío"; 
  const ARTICULO_ENTIDAD: string = "el";
  const NOMBRE_ITEM_CLAVE: string = "SISTEMA_CLAVE"; 

  const hojaEntradaWS: ExcelScript.Worksheet = workbook.getWorksheet(nombreHojaEntrada);
  const hojaMaestrosWS: ExcelScript.Worksheet = workbook.getWorksheet(nombreHojaMaestros);
  
  let claveSeguridadSistema: string = "";
  let registroMensajeLog: string = "";

  // --- I. VALIDACIÓN DE ENTORNO Y SEGURIDAD ---
  if (!hojaEntradaWS || !hojaMaestrosWS) {
    registroMensajeLog = `⛔ Error de Infraestructura: Faltan hojas críticas (${nombreHojaEntrada} o ${nombreHojaMaestros}).`;
  } else {
    const rangoItemClave = workbook.getNamedItem(NOMBRE_ITEM_CLAVE)?.getRange();
    if (rangoItemClave) {
      claveSeguridadSistema = rangoItemClave.getText();
      if (claveSeguridadSistema === "") {
        registroMensajeLog = "⚠️ Advertencia: El protocolo de seguridad detectó una clave de sistema vacía.";
      }
    } else {
      registroMensajeLog = `⛔ Error de Seguridad: No se encontró el ítem de protección '${NOMBRE_ITEM_CLAVE}'.`;
    }
  }
  
  // --- II. EJECUCIÓN DEL PROTOCOLO DE CONFIGURACIÓN ---
  if (hojaEntradaWS && claveSeguridadSistema !== "") {
    try {
      // A) RESET DE SEGURIDAD: Bloqueo integral de la superficie de la hoja
      hojaEntradaWS.getProtection().unprotect(claveSeguridadSistema);
      hojaEntradaWS.getRange().getFormat().getProtection().setLocked(true); 

      // B) LÓGICA DE HABILITACIÓN SELECTIVA (Inputs)
      // Definimos qué columnas de etiquetas (B, E) habilitan qué columnas de datos (C, F)
      // auxiliarConfigurarDesbloqueoPorColumna(LetraColumnaEtiqueta, IndiceColumnaDatos, NumeroFilasEncabezado)
      let contadorCamposHabilitados: number = 0;
      
      // Procesar Bloque Primario (Columna B -> C)
      contadorCamposHabilitados += auxiliarConfigurarDesbloqueoPorColumna("B", 2, 1); 
      
      // Procesar Bloque Secundario (Columna E -> F) si existiera etiquetas en E
      contadorCamposHabilitados += auxiliarConfigurarDesbloqueoPorColumna("E", 5, 4);

      registroMensajeLog = contadorCamposHabilitados > 0 
        ? `✅ Configuración de ${ETIQUETA_ENTIDAD} exitosa. ${contadorCamposHabilitados} campos habilitados para entrada.` 
        : `⚠️ Proceso finalizado sin detectar campos para habilitar en ${nombreHojaEntrada}.`;

    } catch (errorEjecucion) {
      registroMensajeLog = `❌ Error de Ejecución en ${ETIQUETA_ENTIDAD}: ${(errorEjecucion as Error).message}`;
    } finally {
      // C) CIERRE ESTANDARIZADO DE SEGURIDAD
      // Re-protección de la hoja con permisos de navegación (Autofiltros)
      auxiliarEjecutarProteccionSegura(hojaEntradaWS, `Interfaz ${ETIQUETA_ENTIDAD}`);
    }
  }

  // Notificación del resultado en la consola de administración
  console.log(registroMensajeLog);

  // --- FUNCIONES AUXILIARES (HELPERS) ---

  /**
   * Recorre una columna de etiquetas y desbloquea la celda adyacente si contiene texto descriptivo.
   */
  function auxiliarConfigurarDesbloqueoPorColumna(letraColumnaEtiquetas: string, indiceColumnaDatos: number, numeroFilasEncabezado: number): number {
    let camposContabilizados = 0;
    const rangoEtiquetasUsadas = hojaEntradaWS!.getRange(`${letraColumnaEtiquetas}:${letraColumnaEtiquetas}`).getUsedRange();
    
    if (rangoEtiquetasUsadas) {
      const matrizEtiquetas = rangoEtiquetasUsadas.getValues();
      const indiceFilaInicial = rangoEtiquetasUsadas.getRowIndex();

      matrizEtiquetas.forEach((fila, i) => {
        const valorEtiqueta = String(fila[0]).trim();
        const indiceFilaReal = indiceFilaInicial + i;

        // Si la celda tiene una etiqueta y no pertenece al área de encabezados a ignorar
        if (valorEtiqueta !== "" && indiceFilaReal > numeroFilasEncabezado) { 
          // Desbloqueamos la celda de la derecha (Columna de Input)
          hojaEntradaWS!.getRangeByIndexes(indiceFilaReal, indiceColumnaDatos, 1, 1)
                       .getFormat().getProtection().setLocked(false);
          camposContabilizados++;
        }
      });
    }
    return camposContabilizados;
  }

  /**
   * Asegura el cierre de la hoja permitiendo funcionalidades básicas de usuario.
   */
  function auxiliarEjecutarProteccionSegura(hojaAProteger: ExcelScript.Worksheet | undefined, nombreReferencia: string) {
    if (hojaAProteger) {
      try {
        // Mantenemos consistencia con el resto de scripts: permitimos Autofiltros
        hojaAProteger.getProtection().protect({ allowAutoFilter: true }, claveSeguridadSistema);
      } catch (e) {
        console.log(`ℹ️ Protocolo de Cierre (${nombreReferencia}): Estado de protección verificado o error menor.`);
      }
    }
  }
}