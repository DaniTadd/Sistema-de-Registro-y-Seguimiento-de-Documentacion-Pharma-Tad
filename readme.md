# 🚀 SGC-Engine v3.1: High-Integrity Excel-Centric Architecture

![Version](https://img.shields.io/badge/version-4.0.0-blue.svg?style=for-the-badge)
![Tech](https://img.shields.io/badge/Office_Scripts-TypeScript-3178C6.svg?style=for-the-badge)
![Security](https://img.shields.io/badge/Security-SHA--256_Encryption-red.svg?style=for-the-badge)
![Philosophy](https://img.shields.io/badge/UX-Excel--Centric-green.svg?style=for-the-badge)

## 1. Fundamentos de la Arquitectura

La versión 3.0 del SGC-Engine establece una migración desde un modelo basado en flujos en la nube hacia una arquitectura de procesamiento local en Excel (Office Scripts). 

Esta decisión arquitectónica responde a la necesidad de mitigar la latencia y eliminar la dependencia de servidores externos. Al ejecutar la lógica de validación de forma nativa en la hoja de cálculo, se asegura que el registro de la información sea estrictamente contemporáneo al momento de la validación, garantizando la trazabilidad bajo normativas GxP.

---

## ⚖️ ¿Por qué abandonar Power Automate? El Business Case de la "Fricción Cero"

La automatización de versiones anteriores dependía de puentes externos. Esta versión rompe ese esquema basándose en tres pilares de eficiencia:

### A. Eliminación de Latencia (Real-Time Integrity)
Al ejecutar los scripts directamente en el motor de Excel, la respuesta es instantánea. Se elimina el tiempo de "disparo" y "procesamiento" de Power Automate, garantizando que el dato se registre en el momento exacto en que se valida (Contemporaneidad ALCOA+).

### B. Adopción con Fricción Cero y Eficiencia Operativa

El diseño de la herramienta prioriza la adaptabilidad mediante la utilización de recursos preexistentes, sin incurrir en costos de licenciamiento de software de terceros:

* **Adopción Orgánica (Native UX):** La interfaz de captura se mantiene en el entorno habitual (Excel). Si el analista comprende el uso de una hoja de cálculo, comprende el uso del sistema sin salir de su zona de confort.

* **Procesamiento I/O Optimizado:** Para evitar el colapso de memoria frente a las limitaciones de las APIs de Excel, el sistema implementa lectura y escritura por lotes (*Batching*). Las consultas y validaciones se resuelven íntegramente en memoria RAM, reduciendo las llamadas de red y garantizando una respuesta instantánea.

### C. Sin Costos de Licenciamiento
Se elimina la dependencia de conectores premium o flujos de Power Automate que consumen cuotas de ejecución, maximizando el retorno de inversión (ROI) sobre la licencia base de Microsoft 365.

---

## 2. 🛡️ Pilares de Seguridad y Criptografía (Zero-Trust)

Para compensar la eliminación del login de Power Automate, se implementó una capa de seguridad dentro de la planilla:

*   **Identidad Digital SHA-256:** Las contraseñas de los usuarios ya no existen en texto plano. Se almacenan como Hashes criptográficos en la `TablaUsuarios`. Ni siquiera un administrador del libro puede ver la clave real de un operario.
*   **Protocolo de Firma "INICIAR":** Sistema de auto-gestión de credenciales. Los nuevos usuarios o aquellos con claves reseteadas inician su seguridad mediante un "flag" maestro que los obliga a establecer su propia firma digital en el primer uso.
*   **Matriz de Integridad ALCOA+ Expandida:** La tabla de sellos digitales ahora no solo vigila los datos (BBDD e Historial), sino también la integridad de la propia infraestructura de seguridad (`TablaUsuarios`). Si alguien altera manualmente una clave en la tabla, el sistema se bloquea automáticamente.
*   **Double-Check Intencional (Anti-Error):** En transacciones críticas (como **Anular**), el sistema exige el tipeo manual del ID. Esta "fricción positiva" garantiza que la acción sea deliberada y no un clic accidental.

---

## 3. Desacoplamiento de Interfaz y Base de Datos (Ingeniería Creativa)

El principal desafío operativo de utilizar Excel como sistema relacional es la rigidez estructural. Para lograr versatilidad frente a distintos tipos de registros (Desvíos, CAPAs, Equipos), se implementó un motor de abstracción semántica en memoria. 

### 3 A. 🛠️ Innovaciones en la captura de los datos mediante prefijos como banderas (flags):

* **Convención de Búsqueda Dinámica (`BUSQUEDA_`):** Se separó la acción de consultar de la acción de editar. Los campos con este prefijo permiten buscar registros por múltiples criterios (ID, TAG) sin interferir con las columnas reales de la base de datos.
* **Puente de Normalización de Claves Foráneas (`FILTRO_`):** Para vincular entidades (ej. asignar una CAPA a un Desvío), el sistema utiliza menús desplegables amigables en la interfaz. El motor intercepta este prefijo en RAM y lo transforma en la Clave Foránea estricta que exige la base de datos, garantizando la integridad relacional sin exponer códigos complejos al usuario.
* **Inmutabilidad de Identidad:** Una vez que un registro es capturado, su Clave Primaria (ID) queda bloqueada lógicamente. El motor detecta cualquier alteración manual de la identidad y bloquea la transacción.
* **Campos Obligatorios Dinámicos:** Mediante uso del símbolo`*` se indica la obligatoriedad de un campo a la vez que el motor discrimina inteligentemente los campos operativos de los campos de búsqueda, evitando exigir un ID cuando el objetivo de la transacción es, precisamente, crear un registro nuevo.
* **Independencia Estructural:** El uso de prefijos para discriminar la funcionalidad de campos en formularios y BBDD permite insertar columnas auxiliares de soporte (como `COD_ANTERIOR` en un caso de "migración de planillas") sin quebrar la lógica transaccional.

## 4. Estabilidad Transaccional y Concurrencia

* **Control de Concurrencia Optimista (OCC):** El sistema mitiga el riesgo de *Lost Update* (sobreescritura accidental entre múltiples analistas). Al actualizar un registro, el motor compara los *timestamps* en tiempo real; si detecta que otro operador modificó el archivo en el ínterin, la transacción se aborta de forma segura.


## 🛠️ Innovaciones en el Diseño de Datos

### 🧬 Validación de Dependencias Dinámicas
El motor ahora incluye el operador `DEPENDENCIAS_CERRADAS`. Esto permite, por ejemplo, impedir el cierre de un Desvío si existen CAPAs asociadas que aún permanecen abiertas, consultando las relaciones en tiempo real sin necesidad de columnas auxiliares con fórmulas.

### 🧹 Limpieza Quirúrgica (Preservación de Inteligencia)
El motor de limpieza diferencia entre datos transaccionales y datos de sesión. Al finalizar un registro, limpia el formulario pero preserva el **ID** y el **Usuario** por persistencia visual, enviando valores controlados a la base de datos para no corromper columnas calculadas o fórmulas nativas de Excel.

---

## 🚀 Roadmap de Evolución

*   **v3.1 (Actual):** Arquitectura Excel-Centric, Hashes de seguridad, validación de dependencias campos de `BUSQUEDA_` y `FILTRO_`
*   **v4.0 (Visión):** Dashboard nativo de consumo de los datos volcados en esta planilla con datos productivos.

---

## 👨‍💻 El Rigor detrás del Código

Este sistema sigue los principios de la **Rigurosidad Pragmática**:
*   **SESE (Single Entry, Single Exit):** Código predecible y auditable.
*   **Manejo Híbrido de Errores:** Los errores de negocio informan al usuario; los de infraestructura protegen el sistema.
*   **Zero Silent Failures:** Si algo falla, el sistema se bloquea y avisa. No hay lugar para la ambigüedad.

---
*Desarrollado bajo estándares **GMP** (Good Manufacturing Practices) para garantizar que el dato sea el activo más seguro de la compañía.*

