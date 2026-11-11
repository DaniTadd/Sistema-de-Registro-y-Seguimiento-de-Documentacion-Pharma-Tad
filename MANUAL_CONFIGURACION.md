# ⚙️ Manual de Configuración y Arquitectura del Sistema

Este documento explica paso a paso cómo estructurar un libro de Excel Online para que funcione con un motor de **Office Scripts (TypeScript)** orientado a la gestión de registros de calidad (cumplimiento ALCOA+). 

Está pensado tanto para administradores que quieran replicar o adaptar la planilla a nuevos procesos, como para desarrolladores o modelos de IA que necesiten comprender la relación entre la estructura de las hojas y el código.

---

## ⚖️ Alcance, Ventajas y Filosofía de Diseño

Este sistema nace como una **prueba de concepto (ensayo técnico)** para resolver la trazabilidad documental en entornos regulados utilizando herramientas estándar de Microsoft 365, demostrando que es posible inyectar rigor transaccional y criptográfico en una plataforma de ofimática genérica.

### Ventajas
- **Integridad de datos (ALCOA+):** Generación automática de correlativos, firmas electrónicas con *salt* y hash SHA-256 (UTF-8), bloqueo de registros cerrados/anulados y pista de auditoría (*Audit Trail*) automática por diferencias (lógica Delta).
- **Diseño guiado por tablas (Data-Driven):** La creación de nuevos campos o reglas de validación relacional se resuelve agregando filas o columnas en las tablas de Excel, sin necesidad de modificar el código TypeScript.
- **Cero costos extra:** Funciona íntegramente sobre el ecosistema base de Excel Online, Office Scripts y flujos de Power Automate.

### Limitaciones y Criterio de Seguridad
- **Debilidad intrínseca de la ofimática:** Excel no es un motor de base de datos cliente-servidor real. Quien tiene acceso al archivo para correr un script, técnicamente posee permisos de lectura/escritura sobre el documento subyacente en SharePoint/OneDrive.
- **Mitigación mediante integridad:** El sistema no busca resistir ataques maliciosos de ciberseguridad, sino **garantizar la integridad operativa y evitar errores humanos involuntarios** dentro de un grupo reducido y validado (ej. un equipo de Aseguramiento de la Calidad). Las barreras incluyen hojas de datos ocultas (`veryHidden`), protección con contraseña administrada por el script en bloques `finally`, y una verificación criptográfica continua (*hash sealing*) de las tablas maestras mediante `TablaIntegridad`.
- **Concurrencia:** Debido a los tiempos de I/O de Excel Online, el diseño soporta colisión simultánea moderada mediante un Control de Concurrencia Optimista (OCC) basado en sellos de tiempo (`AUDIT_TRAIL`), evitando que dos operadores sobrescriban datos mutuamente.

---

## 🗂️ 1. Estructura de Hojas del Libro

El libro se organiza en cuatro tipos de hojas. La convención de prefijos es obligatoria para que los scripts identifiquen qué proteger, qué ocultar y dónde operar:

1. **`INP_` (Hojas de Formulario / Interfaz de Usuario):**
   * Ejemplos: `INP_MADRE`, `INP_HIJA`, `INP_MAESTRO`.
   * Son las únicas hojas visibles de forma permanente para el operador. Contienen las etiquetas en la **Columna B**, las celdas de carga de datos en la **Columna C**, el cuadro de mensajes (`UI_FEEDBACK`) y los botones asignados a cada script.
2. **`BD_` (Hojas de Base de Datos Activa):**
   * Ejemplos: `BD_ENTIDAD_MADRE`, `BD_ENTIDAD_HIJA`.
   * Contienen las tablas estructuradas de Excel (`TablaEntidadMadre`, `TablaEntidadHija`) con el estado actual de cada registro. Permanecen ocultas y protegidas.
3. **`HIST_` o `_HISTORIAL` (Hojas de Audit Trail):**
   * Ejemplos: `HIST_ENTIDAD_MADRE`, `HIST_ENTIDAD_HIJA`.
   * Contienen las tablas de historial (`TablaEntidadMadreHistorial`, etc.) donde se inserta una fila inmutable (Append-Only) por cada alta, modificación de campo, cambio de estado o anulación.
4. **`MAESTROS` / Configuración Interna:**
   * Contienen las tablas de control del motor: `TablaUsuarios`, `TablaPermisos`, `TablaReglas`, `TablaIntegridad`, `TablaNotificaciones_Outbox` y tablas auxiliares para listas desplegables (ej. `TablaMateriales`).

### Nombres Definidos (Named Ranges) Globales Obligatorios
El libro debe contar con los siguientes nombres definidos a nivel Libro (*Fórmulas -> Administrador de nombres*):
- **`SISTEMA_CLAVE`**: Contraseña con la que los scripts bloquean y desbloquean las hojas.
- **`SISTEMA_SALT`**: Cadena de texto secreta concatenada a las contraseñas de usuarios y datos de tablas antes de calcular los hashes SHA-256.
- **`SISTEMA_HASH_INICIO`**: Palabra clave de inicialización para nuevos sellos (valor estándar: `INICIAR`).
- **`UI_FEEDBACK`**: Rango o celda combinada en las hojas `INP_` donde el sistema imprime el resultado de cada operación (éxito, error de validación o fallo crítico).

---

## 📝 2. Configuración de Hojas de Entrada (`INP_`)

El motor utiliza **mapeo semántico dinámico**: lee la etiqueta de la **Columna B**, la normaliza (pasa a mayúsculas, quita el asterisco `*` y reemplaza espacios por guiones bajos `_`) y busca una columna con ese nombre exacto en la base de datos.

> ⚠️ **Cuidado con las tildes:** Si la columna en la tabla se llama `FECHA_EJECUCION` (sin tilde), la etiqueta en la Columna B del formulario debe ser `FECHA EJECUCION` (sin tilde). Para el motor relacional, `Ó` y `O` son caracteres distintos.

### Reglas de la Columna B (Etiquetas) y Columna C (Valores)
- **Campos Obligatorios (`*`):** Si una etiqueta en la Columna B termina con asterisco (ej. `DESCRIPCION*`), el script impedirá registrar o actualizar si la celda contigua está vacía. Si el campo no aplica, el usuario debe escribir conscientemente `N/A`.
- **Campos Opcionales (sin `*`):** Si se dejan vacíos, el script los autocompleta con `"N/A"` al guardar en la base de datos.
- **Campos de Fecha (`FECHA`):** Cualquier etiqueta que contenga la palabra `FECHA` es evaluada automáticamente por el parser UTC del sistema, exigiendo formato `dd/mm/aaaa` y fecha real de calendario.
- **Celdas Superiores de Testigo (`ESTADO` e `ID_<ENTIDAD>`):** Ubicadas en la parte superior del formulario. Al ejecutar el script auxiliar `Configurar Rangos.ts`, toda fila cuya etiqueta sea `ESTADO` o empiece con `ID_` queda **bloqueada** para escritura manual.
- **Celdas de Búsqueda (`BUSQUEDA_<CAMPO>`):** Permiten recuperar registros (ej. `BUSQUEDA_ID_ENTIDAD`). Quedan desbloqueadas para el usuario.
- **Puente para Claves Foráneas (`FILTRO_<ID_MADRE>`):** En formularios de entidades hijas, el usuario necesita indicar a qué entidad madre pertenece el registro. Como cualquier etiqueta que empieza con `ID_` es bloqueada por el sistema, se utiliza el prefijo `FILTRO_` (ej. `FILTRO_ID_MADRE*`). El script quita este prefijo internamente y guarda el dato en la columna `ID_MADRE` de la tabla hija.

### Sugerencia de Listas Desplegables y Botones
- **Validación de Datos:** Se recomienda configurar listas desplegables nativas de Excel en la Columna C apuntando a tablas maestras (ej. `"Código - Descripción"` para ítems de catálogo). Pueden armarse tablas filtradas dinámicamente con fórmulas de matriz (`FILTRAR`) para que el desplegable de `FILTRO_ID_MADRE*` muestre solo entidades en estado `ABIERTO`.
- **Botones de Acción:** Insertar formas rectangulares en la hoja `INP_` asignadas a los scripts operativos: `Buscar`, `Limpiar Formulario`, `Registrar`, `Actualizar`, `Cerrar / Reabrir` (solo en entidades con ciclo de estado) y `Anular`.
- **Botones de Administración (`Ocultar / Mostrar`):** Se recomienda tener botones (generalmente en una hoja de inicio o maestros) asignados a los scripts de visualización para que el administrador pueda mostrar u ocultar rápidamente las hojas `BD_` y `HIST_` durante tareas de mantenimiento, manteniendo la propiedad `veryHidden` el resto del tiempo.

---

## 🗄️ 3. Estructura de Tablas de Base de Datos e Historial

### Tablas de Base de Datos (`TablaEntidadMadre`, `TablaEntidadHija`, etc.)
Cada tabla puede tener todas las columnas operativas que el proceso requiera (siempre en mayúsculas y con guiones bajos), pero debe respetar estas columnas de sistema:
1. **`ID_<ENTIDAD>` (Primera columna):** Clave primaria (ej. `ID_MADRE`, `ID_HIJA`). El script calcula el correlativo numérico y le antepone el prefijo configurado.
2. **`ID_<MADRE>` (Solo en tablas hijas):** Columna de clave foránea.
3. **`ESTADO`:** Estado del ciclo de vida del registro (`ABIERTO`, `CERRADO`, `ANULADO`, `VIGENTE`).
4. **`USUARIO`:** Identificador del analista que realizó la última acción.
5. **`AUDIT_TRAIL`:** Timestamp de la última operación (Control de Concurrencia Optimista).

### Tablas de Historial (`TablaEntidadHistorial`)
Todas las tablas de historial tienen exactamente **7 columnas fijas en este orden**:
1. `ID_HISTORIAL` (Correlativo automático)
2. `ID_REGISTRO` (Clave primaria de la entidad afectada)
3. `FECHA_HORA` (Timestamp exacto)
4. `USUARIO` (Usuario firmante)
5. `ACCION` (`ALTA DE REGISTRO`, `MODIFICACION`, `CAMBIO DE ESTADO`, `ANULACION`)
6. `DETALLE_CAMBIOS` (Formato Delta: `CAMPO: [Valor Anterior] -> [Valor Nuevo]`)
7. `MOTIVO` (Justificación ingresada al firmar)

### Tablas Especiales de Eventos
Si se utiliza la regla relacional `LOG_EVENTO` para trazar cambios específicos (ej. cambios de disposición de un ítem), la tabla destino debe tener las columnas: `ID_EVENTO`, `ID_<CLAVE_FORANEA>`, `<NOMBRE_DEL_CAMPO_A_TRAZAR>`, `MOTIVO`, `USUARIO` y `FECHA_EVENTO`.

---

## 💻 4. Diccionario `CONFIGURACION_ENTIDADES` en el Código

Al inicio de los scripts principales existe la constante `CONFIGURACION_ENTIDADES`. Es un diccionario TypeScript donde la **clave** es el nombre de la hoja `INP_` y el **valor** define el comportamiento de esa entidad:

```typescript
const CONFIGURACION_ENTIDADES: { [key: string]: ConfiguracionEntidad } = {
    "INP_ENT": {
        tabla: "TablaEntidad",               // Nombre de la tabla BD
        historial: "TablaEntidadHistorial",  // Nombre de la tabla Historial
        prefijo: "ENT-",                     // Prefijo para el ID autogenerado
        etiqueta: "entidad",                 // Nombre legible para mensajes
        genero: "a",                         // "o" (registrado) o "a" (registrada)
        articulo: "la",                      // "el" o "la" para armar frases
        estadoInicial: "ABIERTO",            // Estado al crear y al reabrir
        estadosBloqueantes: ["CERRADO", "ANULADO"] // Impiden editar el registro
    }
};
```

> **Nota Arquitectónica:** En este diccionario, todas las entidades se declaran por igual. Lo que convierte a una entidad en "Madre" o "Hija" son las relaciones y restricciones configuradas en `TablaReglas`.

## 🧠 5. Motor de Reglas de Negocio (`TablaReglas`)

`TablaReglas` permite definir validaciones cronológicas, relacionales y de ciclo de vida sin tocar el código. Posee 5 columnas:

| ENTIDAD | CAMPO_A | OPERADOR | CAMPO_B | MENSAJE |
| :--- | :--- | :--- | :--- | :--- |

### Activación y Desactivación (`INACTIVA_`)
El código busca coincidencia exacta con el nombre del `OPERADOR`. Para desactivar una regla sin borrar la fila, basta con anteponer `INACTIVA_` (ej. `INACTIVA_DEPENDENCIAS_CERRADAS`). Si una regla está activa pero la tabla declarada en `CAMPO_B` no existe, el sistema lanzará un fallo preventivo por seguridad.

### Sintaxis de `CAMPO_B`
1. **`OTRO_CAMPO`**: Para comparar dos fechas del mismo formulario (ej. `FECHA_INICIO`).
2. **`NombreTabla`**: Para reglas unarias o de log (ej. buscar unicidad dentro de `TablaMaestra`).
3. **`TablaDestino[COLUMNA_CLAVE]`**: Indica tabla y clave foránea (ej. `TablaMadre[ID_MADRE]`).
4. **`TablaDestino[COLUMNA_CLAVE];[COLUMNA_ESTADO];VALOR_ESTADO`**: Sintaxis compuesta para evaluar el estado de una entidad relacionada (ej. `TablaMadre[ID_MADRE];[ESTADO];ABIERTO`).
5. **`N/A`**: Marcador visual para operadores unarios que solo evalúan `CAMPO_A` consigo mismo.

### Catálogo de Operadores Programados

| Operador | Aplicación | Ejemplo de CAMPO_B | Función principal |
| :--- | :--- | :--- | :--- |
| `<= / >=` | Registrar, Actualizar | `FECHA_CIERRE` | Compara cronología. En `>=`, si se carga A, exige que B no sea "N/A". |
| `EXISTE_EN` | Registrar, Actualizar | `TablaMadre[ID_MADRE]` | Verifica que el código foráneo exista en la tabla maestra. |
| `MADRE_ABIERTA_AL_CREAR` | Registrar, Actualizar* | `TablaMadre[ID_MADRE];[ESTADO];ABIERTO` | Al Registrar: Exige que la entidad madre esté abierta. Al Actualizar: Solo se evalúa si el usuario intenta reasignar el registro a otra entidad madre. |
| `ESTADO_DISTINTO_A` | Registrar, Actualizar | `TablaMadre[ID_MADRE];[ESTADO];ANULADO` | Bloquea la edición si la entidad relacionada posee el estado prohibido indicado. |
| `ES_UNICO_ALFANUMERICO` | Registrar, Actualizar | `TablaMaestra` | Limpia espacios/guiones y verifica unicidad de códigos o series. |
| `LOG_EVENTO` | Registrar, Actualizar | `TablaEventos` | Si el valor de `CAMPO_A` muta, inserta una fila de trazabilidad en la tabla destino. |
| `DEPENDENCIAS_CERRADAS` | Cambio de Estado | `TablaHija[ID_MADRE]` | Bloquea el cierre de una entidad si posee registros dependientes abiertos. |
| `REQUERIDO_PARA_CERRAR` | Cambio de Estado | `N/A` | Bloquea el paso a estado de cierre si `CAMPO_A` está vacío o es "N/A". |
| `ANULAR_EN_CASCADA` | Anular | `TablaHija[ID_MADRE]` | Al anular un registro madre, anula automáticamente todos sus dependientes. |

### Cómo crear un operador nuevo en TypeScript
Para agregar un operador (ej. `MAYOR_A_CERO`):
1. Agregar la fila en `TablaReglas` con el nuevo operador.
2. En los scripts `.ts`, dentro de la Fase 2 (Evaluación en RAM), agregar el bloque lógico:

```typescript
else if (operador === "MAYOR_A_CERO") {
    if (valorAValidar && valorAValidar !== "N/A") {
        if (Number(valorAValidar) <= 0) {
            listaErroresValidacion.push(mensajeErrorRegla);
        }
    }
}
```

## 🔐 6. Seguridad, Integridad (`TablaIntegridad`), Usuarios y Notificaciones

### 1. `TablaIntegridad` (Sellado Criptográfico SHA-256)
Posee dos columnas: `TABLA` y `HASH_SHA256`.
- **Mecanismo:** El script calcula en tiempo real el hash SHA-256 de las tablas maestras y de datos. Si un usuario altera una celda manualmente burlando la interfaz, los hashes diferirán y el sistema abortará operaciones futuras. Al finalizar una transacción válida por UI, el script sella la tabla con el nuevo hash.
- **Inicialización (`INICIAR`):** Al modificar configuraciones (reglas, usuarios, permisos), el administrador debe escribir la palabra `INICIAR` en el hash correspondiente. El sistema adoptará y sellará el nuevo estado en la siguiente ejecución.
- **Desactivación modular (`INACTIVA_`):** Anteponiendo este prefijo al nombre de la tabla en `TablaIntegridad` (ej. `INACTIVA_TablaEntidad`), el motor omite la validación de hash para esa tabla específica.

### 2. `TablaUsuarios` y `TablaPermisos`
- **Gestión de Claves:** Para crear o blanquear la clave de un usuario, se coloca su ID y la palabra `INICIAR` como hash. En su primera transacción, la clave ingresada será salada (`SISTEMA_SALT`), hasheada en SHA-256 y guardada definitivamente.
- **Control de Acceso Relacional (ACL):** `TablaPermisos` vincula cada `USUARIO` con una `ENTIDAD` (o `TODAS`) y la `ACCION` permitida (`ALTA`, `MODIFICACION`, `CAMBIO_ESTADO`, `ANULACION` o, todas las acciones con a`*`).

### 3. Integración RPA (`TablaNotificaciones_Outbox`)
- Cada operación exitosa encola una transacción con estado `PENDIENTE`.
- Un flujo externo (Power Automate) puede programarse para consumir esta tabla mediante intervalos por lotes (*Batch Processing*), generando un Digest diario en PDF/Email, para luego invocar al script de limpieza pasándole los IDs procesados.
- **Desactivación:** Cambiando el nombre a `INACTIVA_TablaNotificaciones_Outbox` en el administrador de Excel, el motor omite el encolado sin requerir modificaciones en el código.

---

## 🚀 7. Segregación de Entornos y Control de Cambios (DEV vs. PROD)

En cumplimiento de las buenas prácticas de validación de sistemas computarizados (GAMP 5 / ALCOA+), el sistema implementa una **estricta segregación** entre el desarrollo/ensayo de nuevas funcionalidades y la operación productiva en planta.

### Arquitectura de Punteros en Office Scripts
Los botones de acción en Excel Online no incrustan el código fuente, sino que guardan un puntero hacia un contenedor `.osts` específico. Para evitar que la edición de scripts afecte la operación en vivo, se definen dos circuitos independientes:

1. **Entorno de Desarrollo (Sandbox Local / OneDrive):**
   * **Archivos fuente:** Nombrados con el sufijo `_DEV.ts` (ej. `Registrar Universal_DEV.ts`).
   * **Contenedores `.osts`:** Alojados en la carpeta personal de OneDrive (`RUTA_ONEDRIVE_DEV`), bajo la denominación `_DEV.osts`.
   * **Libro de pruebas:** Sus botones están vinculados exclusivamente a los scripts `_DEV`.
2. **Entorno de Producción (SharePoint Corporativo de Calidad):**
   * **Contenedores `.osts`:** Alojados en la biblioteca de documentos del sitio de SharePoint (`RUTA_SHAREPOINT_PROD`), con nombres limpios (ej. `Registrar Universal.osts`).
   * **Libro productivo oficial:** Sus botones apuntan a los scripts limpios de SharePoint, quedando completamente aislados de cualquier prueba o refactorización en curso.

### Automatización del Despliegue (`puente.py`)
El script de automatización elimina la manipulación manual de archivos JSON/OSTS y garantiza la consistencia del pase a producción:
* **`python puente.py` (Modo DEV):** Sincroniza el código de trabajo contra los contenedores `_DEV.osts` de OneDrive de forma instantánea.
* **`python puente.py --prod` (Pase a Producción):** Despoja automáticamente el sufijo `_DEV` del nombre y refresca los contenedores oficiales en SharePoint, asegurando que el código liberado sea idéntico al probado.
* **`python puente.py --ambos`:** Sincroniza ambos destinos en una única ejecución cuando se requiere alineación inmediata.
