# 📚 Manual de Lógica y Memoria Técnica: Sistema Universal de Gestión (SGC)

**Versión:** 2.0 (Edición Auditable)  
**Tecnología:** Office Scripts (TypeScript) + Power Automate + Python Bridge  
**Estándar:** GMP / ALCOA+ (Integridad de Datos)  
**Arquitectura:** Motor Agnóstico de Alto Desempeño (SESE)

---

## 1. Filosofía del Sistema (Arquitectura)

El sistema se basa en un diseño agnóstico y descriptivo que garantiza flexibilidad, cumplimiento normativo y transparencia total en la ejecución.

### A. Mapeo Semántico Dinámico (Abstracción)
El código no utiliza coordenadas fijas (ej. "C5"). La navegación se basa en la correspondencia de texto:
* **Lógica:** El script recorre las etiquetas de la Columna B (o E), normaliza el texto (MAYÚSCULAS y `GUION_BAJO`) y busca la coincidencia exacta en los encabezados de la Base de Datos.
* **Nomenclatura Técnica:** Se utiliza una nomenclatura descriptiva para que el flujo de datos sea autoexplicativo:
    * `matrizEtiquetas`: Rango capturado de la interfaz de usuario.
    * `objetoDatosFormulario`: Estructura de datos procesados listos para validación.
    * `encabezadosTabla`: Mapeo de la estructura física de la base de datos.
    * `matrizValoresDB`: Datos crudos para cálculos y comparaciones.

### B. Salvaguarda de Integridad Estricta (Mismatch Check)
El sistema implementa una barrera de seguridad activa para evitar la sobreescritura accidental o desincronizada:
* **Validación de Identidad:** El motor compara el `idDesdePanel` (parámetro externo) contra el `idEnHoja` (dato local en el formulario).
* **Condición de Aborto:** Si existe discrepancia entre ambos Identificadores, el sistema lanza una excepción (`throw Error`) y detiene la transacción antes de afectar la persistencia de los datos.

### C. Lógica de Cierre Seguro (SafeProtect)
El flujo de ejecución garantiza la seguridad del entorno mediante el bloque `finally`:
* **Protocolo auxiliarProtegerHoja:** Al finalizar cada proceso, el sistema reaplica la protección con permisos de navegación (Autofiltro) y edición restringida.
* **Transaccionalidad:** La integridad de la operación principal es prioritaria, pero el estado de la protección final siempre se verifica y se registra en el log técnico.

### D. Seguridad por "Puente" (Bridge)
La seguridad está desacoplada del código fuente mediante el uso de **Nombres Definidos**:
1. No se almacenan credenciales dentro de los scripts.
2. La clave de protección se recupera dinámicamente desde el ítem `SISTEMA_CLAVE`.

---

## 2. Estructura de Datos (Compliance ALCOA+)

El sistema distingue entre datos de negocio (flexibles) y metadatos de auditoría (rígidos).

### 2.1 Metadatos de Auditoría (Trazabilidad Inmutable)
Campos obligatorios gestionados exclusivamente por el motor lógico:

| Campo | Función | Comportamiento |
| :--- | :--- | :--- |
| **ID** | Identificador único | Prefijo dinámico + Máximo correlativo numérico + 1. |
| **ESTADO** | Ciclo de vida | Controlado por el sistema (ABIERTO / CERRADO / ANULADO). |
| **AUDIT_TRAIL** | Timestamp | Fecha/Hora de la operación generada en el servidor (ART). |
| **USUARIO** | Firma Digital | Email del usuario responsable de la acción. |
| **MOTIVO** | Justificación | Parámetro obligatorio para cualquier edición o cambio de estado. |
| **CAMBIOS** | Log de diferencias | Detalle automático de campos modificados: `[Campo]: [V. Anterior] -> [V. Nuevo]`. |

### 2.2 Gestión de Persistencia
* **Inyección de NULL:** En el registro de nuevos datos, las columnas que no figuran en el formulario reciben un valor `null`, permitiendo que Excel ejecute el **autorrelleno de fórmulas** de forma nativa.
* **Commit Quirúrgico:** Las actualizaciones solo sobrescriben las celdas modificadas en el formulario, preservando la integridad de las columnas de cálculo existentes en la tabla.

---

## 3. Motor de Reglas y Validación Jerárquica

La validación lógica se parametriza desde la `TablaReglas` en la hoja `MAESTROS`.

* **EXISTE_EN:** Verifica integridad referencial en tablas maestras.
* **ESTA_ABIERTO (Validación Jerárquica):** Permite validar el estado de una entidad "Madre" desde un proceso de una entidad "Hija".
    * **Sintaxis:** `TablaMadre[ColumnaID];[ColumnaEstado];ValorEsperado`
* **Operadores Comparativos:** Soporte para validaciones cronológicas (`<`, `>`, `<=`, `>=`) entre campos del formulario.

---

## 4. Gestión de Interfaz y Sincronización

* **Sincronizador Python:** Herramienta externa que garantiza la paridad absoluta entre el repositorio de código local (Git) y los archivos operativos en la nube (.osts).
* **Configuración de Rangos:** Proceso automático de bloqueo de etiquetas y desbloqueo de celdas de entrada para minimizar el error humano.
* **Seguridad Visual:** La función de búsqueda escribe el ID localizado directamente en el formulario para confirmar la identidad del registro activo.

---

## 5. Matriz de Solución de Problemas (Troubleshooting)

| Mensaje de Error | Canal | Causa Probable |
| :--- | :---: | :--- |
| **"ERROR: Mismatch de ID"** | ⚠️ Feedback | El ID del panel no coincide con el de la hoja. Se requiere re-sincronizar. |
| **"AccessDenied"** | ⛔ Sistema | Error en el ítem `SISTEMA_CLAVE` o protección manual de hoja. |
| **"Faltan columnas..."** | ⛔ Sistema | Inconsistencia entre etiquetas de UI y encabezados de Base de Datos. |
| **"ID Requerido"** | ⚠️ Feedback | Intento de operación de edición sin un registro cargado. |
| **"Registro ANULADO"** | ⚠️ Feedback | Intento de modificar un estado inmutable definido por protocolo. |

---