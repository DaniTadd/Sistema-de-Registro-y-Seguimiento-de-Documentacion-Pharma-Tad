# Manual de Lógica, Configuración y Memoria Técnica: Motor de Registro y Seguimiento de Documentación

**Versión:** 1.0
**Tecnología:** TypeScript (Office Scripts)
**Estándar:** GMP / ALCOA+
**Propósito:** Motor agnóstico para gestionar flujos de documentación de calidad (Desvíos, Reclamos, CC) mediante configuración en Excel, asegurando integridad transaccional.

---

## 1. Filosofía del Sistema (Arquitectura)

El sistema se basa en cuatro pilares que garantizan flexibilidad, robustez y mantenibilidad:

### A. Mapeo Dinámico (Abstracción)
El código no contiene referencias fijas a celdas de datos (ej. C5, F8).
* **Funcionamiento:** El script lee las etiquetas de la Columna B (o E), identifica qué dato se pide y busca su coincidencia exacta en los encabezados de la Base de Datos.
* **Ventaja:** Permite agregar filas, mover campos o replicar el sistema sin tocar el código.

### B. Estructura SESE (Single Entry, Single Exit) & SafeProtect
Para garantizar la seguridad de los datos, el flujo de ejecución es lineal y el cierre es a prueba de fallos.
* **Patrón "Check-Before-Act" (Éxito Silencioso):** Al finalizar cualquier script (en el bloque `finally`), el sistema invoca una función interna `safeProtect`.
* **Lógica:** Intenta proteger la hoja. Si recibe un error `InvalidOperation` (significa que *ya estaba protegida*), lo ignora deliberadamente. Esto evita falsos positivos de error y asegura que la hoja termine siempre bloqueada, sin importar el estado inicial.

### C. Seguridad por "Puente" (Bridge)
El código es público y no contiene credenciales hardcodeadas. Utiliza un **Nombre Definido** (`SISTEMA_CLAVE`) en Excel que apunta a la celda real de la contraseña, desacoplando el código de la configuración de seguridad.

### D. Arquitectura "Clean Code" (Estándar de Desarrollo)
Todos los scripts siguen una estructura periodística estricta para facilitar la auditoría y evitar errores de duplicidad en el entorno de Excel Online:
1.  **Configuración:** Constantes y lecturas iniciales al tope del archivo.
2.  **Lógica de Negocio:** Validaciones y transacciones en el cuerpo central.
3.  **Helpers Encapsulados:** Las funciones auxiliares (`reportarError`, `safeProtect`) se definen **dentro** de la función `main`, al final del archivo. Esto permite que compartan el *scope* (acceso a variables como `clave` o `UX`) sin necesidad de pasarlos como argumentos repetidamente.

---

## 2. Estructura de Datos Híbrida (Compliance)

El sistema distingue entre datos flexibles del usuario y metadatos rígidos de auditoría.

### 2.1 Datos de Negocio (Dinámicos)
Cualquier campo definido en el Input (ej. "Lote", "Máquina").
* **Comportamiento:** Si la etiqueta existe en BD, se guarda. Si se borra del Input, el sistema rellena con **"N/A"** (No Aplica) en lugar de dejar vacíos.

### 2.2 Metadatos de Auditoría (Estáticos)
Columnas obligatorias para cumplir con ALCOA+. Sus nombres son fijos en el código.

| Campo | Función | Comportamiento |
| :--- | :--- | :--- |
| **ID** | Identificador único | Autonumérico gestionado por el sistema. |
| **ESTADO** | Ciclo de vida | Controlado por scripts (Abierto/Cerrado/Anulado). |
| **AUDIT TRAIL** | Timestamp | Fecha/Hora inmutable de creación/modificación. |
| **USUARIO** | Firma | Obligatorio para cualquier cambio de estado o edición. |
| **MOTIVO** | Justificación | Obligatorio para auditoría de cambios. |
| **CAMBIOS** | Log de diferencias | Generado automáticamente: `[Campo: Valor A -> Valor B]`. |

---

## 3. Lógica de Flujos Específicos

### 3.1 Actualización de Desvíos ("Cajero Amable")
El script de actualización prioriza la validación de datos sobre la burocracia de la firma.
1.  **Validación de Datos:** Primero verifica que todos los campos del formulario cumplan con las reglas (obligatorios, tipos de datos, lógica de negocio).
2.  **Detección de Cambios:** Verifica si el usuario realmente modificó algún dato respecto a la BD.
3.  **Solicitud de Firma:** Solo si los datos son válidos Y existen cambios reales, el sistema exige completar **Usuario** y **Motivo**.
    * *Ventaja:* Evita frustrar al usuario pidiendo firma cuando el formulario aún tiene errores de carga.

### 3.2 Anulación (Acción Destructiva)
La anulación es lógica, no física. El registro permanece en la BD pero su estado cambia a "ANULADO". Esta acción es irreversible mediante los scripts estándar y requiere firma obligatoria.

---

## 4. Configuración del Formulario (Hoja INPUT)

**Requisito:** Para realizar configuraciones estructurales, el usuario debe contar con la clave de desbloqueo.

### 4.1 Crear nuevos campos
Para agregar un dato nuevo al formulario:
1.  Desproteja la hoja.
2.  Escriba el nombre del nuevo campo en la **Columna B** (ej. "TIPO DE FALLA").
3.  Asegúrese de que exista una columna con **exactamente el mismo nombre** en la tabla de la hoja de base de datos (`BD_DESVIOS`).
4.  Ejecute el script **"Configurar Rangos"** para desbloquear la nueva celda.

### 4.2 Campos Obligatorios y Normalización
El sistema maneja la integridad de los datos según la configuración de la etiqueta:
* **Obligatorio (`*`):** Agregue un asterisco al final de la etiqueta (ej. `LOTE*`).
* **Opcional (Sin `*`):** Si se deja vacío, se guarda como "N/A".

---

## 5. Configuración de Reglas de Negocio (Hoja MAESTROS)

La validación lógica se controla desde la `TablaReglas`.
* **Lógica de Validación:** El sistema crea una "Fila Hipotética" (mezclando datos actuales de BD + nuevos datos del Input) y valida las reglas sobre ese resultado final antes de guardar.

---

## 6. Matriz de Solución de Problemas

| Síntoma / Mensaje | Tipo | Causa Probable y Solución |
| :--- | :---: | :--- |
| **"Error Configuración..."** | ⛔ | **Falta Nombre Definido.** Verifique que exista `SISTEMA_CLAVE` en el Excel. |
| **"AccessDenied"** | ⛔ | **Clave Incorrecta.** La contraseña en la celda apuntada no coincide con la de la hoja. |
| **"Faltan columnas..."** | ⛔ | **Estructura Rota.** Se borró una columna crítica (`ID`, `ESTADO`). Restaúrela. |
| **"Falta el argumento..."** | ⛔ | **Desincronización de Claves.** La hoja (BD o Input) tiene una contraseña diferente a la de `SISTEMA_CLAVE`, o está protegida sin contraseña. Unifique las claves manualmente. |
| **"InvalidOperation"** (Consola) | ℹ️ | **Aviso de Seguridad.** El sistema intentó proteger una hoja que ya estaba protegida. Es un comportamiento esperado (SafeProtect) y no afecta al usuario. |
| **Datos quedan como "N/A"** | ⚠️ | **Error de Mapeo.** Diferencia de escritura entre Input y BD (ej. espacios o tildes). |
| **Mensaje Naranja** | ⚠️ | **Alerta de Negocio.** Faltan firmas (Usuario/Motivo) o acción destructiva. No es un error técnico. |