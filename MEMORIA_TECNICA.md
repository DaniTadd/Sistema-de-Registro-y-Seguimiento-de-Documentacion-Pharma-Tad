# 📚 Manual de Lógica y Memoria Técnica: Sistema Universal de Gestión (SGC)

**Versión:** 1.0  
**Tecnología:** Office Scripts (TypeScript)  
**Estándar:** GMP / ALCOA+ (Integridad de Datos)  
**Arquitectura:** Motor Agnóstico de Alto Desempeño (SESE)

---

## 1. Filosofía del Sistema (Arquitectura)

El sistema se basa en cuatro pilares que garantizan flexibilidad y cumplimiento normativo:

### A. Mapeo Dinámico (Abstracción)
El código no contiene referencias fijas a celdas (ej. "C5"). 
* **Lógica:** El script lee las etiquetas de la Columna B (o E), normaliza el texto (MAYÚSCULAS y `GUION_BAJO`) y busca la coincidencia exacta en los encabezados de la Base de Datos.
* **Normalización:** Los asteriscos (`*`) se utilizan para identificar campos **Obligatorios** en la UI, pero se remueven durante el mapeo para encontrar la columna correspondiente.

### B Estructura SESE & Lógica de Cierre Seguro
Para garantizar la integridad, el flujo de ejecución es lineal y el cierre es a prueba de fallos mediante el bloque `finally`:

* **Patrón "Check-Before-Act" (Éxito Silencioso):** Al finalizar cualquier script, el sistema intenta reaplicar la protección. Si la hoja ya está protegida o el comando falla, el script captura el error (`catch`) para evitar un "crash" del sistema.
* **Transaccionalidad:** El éxito de la operación principal (ej. registrar) no depende del éxito de la protección final. El aviso `[⚠️ Seguridad]` se adjunta al registro interno del script (consola) para diagnóstico técnico sin interrumpir la experiencia del usuario.

### C. Seguridad por "Puente" (Bridge)
El sistema utiliza un **Nombre Definido** (`SISTEMA_CLAVE`) que apunta a la celda que contiene la contraseña. Esto permite:
1. Desacoplar la seguridad del código (no hay contraseñas hardcodeadas).
2. Actualizar la clave global desde un solo punto sin editar los scripts.

### D. Arquitectura "Clean Code"
Los scripts están diseñados para ser autocontenidos debido a que Office Scripts no permite llamadas externas:
1. **Configuración de Identidad:** Variables `ENT` (entidad), `ART` (artículo) y `GEN` (género) al inicio para personalizar mensajes.
2. **Helpers Encapsulados:** Funciones como `protect`, `updateUI` y `parseDateToNum` se definen dentro de `main` para compartir el *scope* de variables críticas.

### E Gestión de Errores y Excepciones
El sistema categoriza los fallos según su impacto en la integridad y la necesidad de intervención:

1.  **Errores de Sistema (Excepciones):** Se gestionan mediante `throw`. Son fallos críticos (ej. tablas faltantes o falta de clave) que detienen la ejecución inmediatamente para proteger la base de datos.
2.  **Errores de Negocio (Validaciones):** No detienen el script. Se informan al usuario en la **celda de feedback** (ej. "Falta Fecha") para que pueda corregirlos sin que el motor de ejecución "explote".
3.  **Advertencias de Mantenimiento (Silenciosas):** Se registran únicamente en la **consola de desarrollador**. Incluyen conflictos de protección de hoja que no afectan el éxito de la transacción principal.
---

#### 🛠️ Jerarquía de Visibilidad
Esta distinción asegura que el usuario solo vea lo que puede corregir, mientras que los detalles técnicos quedan para auditoría:

| Síntoma | Canal de Aviso | Gravedad | Explicación |
| :--- | :--- | :---: | :--- |
| **Cartel Rojo de Excel** | UI de Office Scripts | ⛔ Crítico | Fallo estructural (el código no pudo ni empezar). |
| **Mensaje Gris/Naranja** | Celda de Feedback | ⚠️ Advertencia | Error del usuario (faltan datos o reglas de negocio). |
| **Log en Consola** | Panel de Editor | ℹ️ Info | Aviso técnico (SafeProtect, tiempos de ejecución). |
---

## 2. Estructura de Datos (Compliance ALCOA+)

El sistema distingue entre datos de negocio (flexibles) y metadatos de auditoría (rígidos).

### 2.1 Metadatos de Auditoría (Estáticos)
Columnas obligatorias cuyos nombres están fijos en la lógica del motor:

| Campo | Función | Comportamiento |
| :--- | :--- | :--- |
| **ID** | Identificador único | Prefijo dinámico (ej: `D-`) + Máximo correlativo + 1. |
| **ESTADO** | Ciclo de vida | Controlado por scripts (ABIERTO / CERRADO / ANULADO). |
| **AUDIT_TRAIL** | Timestamp | Fecha/Hora inmutable de la operación (Huso Horario ART). |
| **USUARIO** | Firma Digital | Email del usuario que ejecutó la acción. |
| **MOTIVO** | Justificación | Obligatorio para cualquier modificación o anulación. |
| **CAMBIOS** | Log de diferencias | Generado en Actualizar: `[Campo: Valor A -> Valor B]`. |

### 2.2 Protección de Fórmulas y "N/A"
* **Registrar:** Si una columna de la tabla no está en el formulario, el script envía un valor `null`. Esto permite que Excel dispare el **autorrelleno automático de fórmulas**.
* **Actualizar:** Utiliza un **"Commit Quirúrgico"**; solo se sobrescriben las celdas que el usuario modificó en el formulario, protegiendo las fórmulas existentes en otras columnas de la fila.
* **Campos Opcionales:** Si un campo sin asterisco se deja vacío, el sistema guarda **"N/A"** para evitar celdas nulas involuntarias.

---

## 3. Motor de Reglas y Validación

La validación lógica se controla desde la `TablaReglas` en la hoja `MAESTROS`.

* **Lógica de Validación:** El sistema utiliza un objeto puente (`valFuente`) para unificar los datos del formulario y validarlos contra las reglas antes de escribir en la BD.
* **Operadores Soportados:** * `<` / `>` / `<=` / `>=`: Comparaciones lógicas (principalmente fechas).
    * `EXISTE_EN`: Verifica que el dato ingresado exista en una tabla maestra externa (ej: `TablaProductos[Codigo]`).

---

## 4. Gestión de Filtros e Interfaz

* **Limpieza Automática:** Los scripts de **Registrar** y **Buscar** limpian los filtros de la tabla al inicio. Esto garantiza que el nuevo registro o el registro buscado sean siempre visibles para el usuario.
* **Tratamiento de Fechas:** Para evitar desfasajes por zona horaria, el script de **Buscar** recupera el valor serial de la fecha y fuerza el formato local `dd/mm/yyyy` en el formulario.

---

## 5. Matriz de Solución de Problemas

| Síntoma / Mensaje | Tipo | Causa Probable y Solución |
| :--- | :---: | :--- |
| **"AccessDenied"** | ⛔ | **Clave Incorrecta.** La contraseña en `SISTEMA_CLAVE` no coincide con la de la hoja. |
| **"Faltan columnas..."** | ⛔ | **Estructura Rota.** Se borró o renombró una columna crítica (`ID`, `ESTADO`). Restaure el encabezado exacto. |
| **"ID Requerido"** | ⚠️ | **Falta ID.** El campo ID está vacío o tiene "N/A" en una operación de Actualizar/Anular. |
| **"Fecha Inválida"** | ⛔ | **Formato Incorrecto.** Se ingresó un texto que no puede convertirse a fecha (`dd/mm/yyyy`). |
| **Datos quedan como "N/A"** | ⚠️ | **Error de Mapeo.** Diferencia de escritura (espacios, tildes) entre la etiqueta del Input y el encabezado de la BD. |
| **Fórmulas Borradas** | ⛔ | **Error de Configuración.** Se omitió la lógica de envío de `null` para columnas de cálculo en el script. |
| **`[⚠️ Seguridad]`** | ℹ️ | **Aviso de Protección.** El script terminó con éxito pero no pudo reaplicar la protección (hoja ya bloqueada). |