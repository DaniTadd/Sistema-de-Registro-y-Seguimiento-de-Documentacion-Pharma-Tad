# 🚀 SGC-Engine v3.0: High-Integrity Excel-Centric Architecture

![Version](https://img.shields.io/badge/version-3.0.0-blue.svg?style=for-the-badge)
![Tech](https://img.shields.io/badge/Office_Scripts-TypeScript-3178C6.svg?style=for-the-badge)
![Security](https://img.shields.io/badge/Security-SHA--256_Encryption-red.svg?style=for-the-badge)
![Philosophy](https://img.shields.io/badge/UX-Excel--Centric-green.svg?style=for-the-badge)

## 💎 Visión de Ingeniería: Del Cloud-Flow al Local-Power

La versión 3.0 del **SGC-Engine** representa un salto evolutivo en la arquitectura de sistemas de calidad. Hemos migrado de un modelo híbrido basado en flujos de nube (Power Automate) a una arquitectura **Excel-Centric** pura. 

¿Por qué? Porque en la industria farmacéutica y en entornos de alta criticidad, la **latencia** y la **fricción** son los enemigos de la integridad de datos. Al centralizar la lógica en Office Scripts, eliminamos los tiempos de espera del servidor y devolvemos el control total al usuario en su entorno habitual.

---

## ⚖️ ¿Por qué abandonar Power Automate? El Business Case de la "Fricción Cero"

Históricamente, la automatización dependía de puentes externos. Esta versión rompe ese esquema basándose en tres pilares de eficiencia:

### 1. Eliminación de Latencia (Real-Time Integrity)
Al ejecutar los scripts directamente en el motor de Excel, la respuesta es instantánea. Eliminamos el tiempo de "disparo" y "procesamiento" de Power Automate, garantizando que el dato se registre en el momento exacto en que se valida (Contemporaneidad ALCOA+).

### 2. Adopción con Resistencia Cero (Native UX)
No obligamos al usuario a salir de su zona de confort. El sistema utiliza los paneles nativos de parámetros de Excel para pedir confirmaciones e IDs. Si el usuario sabe usar Excel, ya está capacitado para usar el sistema. **Capacitación = Cero costo adicional.**

### 3. Reducción de Costos de Licenciamiento
Se elimina la dependencia de conectores premium o flujos de Power Automate que consumen cuotas de ejecución, maximizando el retorno de inversión (ROI) sobre la licencia base de Microsoft 365.

---

## 🛡️ Pilares de Seguridad y Criptografía (Zero-Trust)

Para compensar la eliminación del login de Power Apps, hemos implementado una capa de seguridad de grado bancario dentro de la planilla:

*   **Identidad Digital SHA-256:** Las contraseñas de los usuarios ya no existen en texto plano. Se almacenan como Hashes criptográficos en la `TablaUsuarios`. Ni siquiera un administrador del libro puede ver la clave real de un operario.
*   **Protocolo de Firma "INICIAR":** Sistema de auto-gestión de credenciales. Los nuevos usuarios o aquellos con claves reseteadas inician su seguridad mediante un "flag" maestro que los obliga a establecer su propia firma digital en el primer uso.
*   **Matriz de Integridad ALCOA+ Expandida:** La tabla de sellos digitales ahora no solo vigila los datos (BBDD e Historial), sino también la integridad de la propia infraestructura de seguridad (`TablaUsuarios`). Si alguien altera manualmente una clave en la tabla, el sistema se bloquea automáticamente.
*   **Double-Check Intencional (Anti-Error):** En transacciones críticas (como **Anular**), el sistema exige el tipeo manual del ID. Esta "fricción positiva" garantiza que la acción sea deliberada y no un clic accidental.

---

## 🛠️ Innovaciones en el Diseño de Datos

### 🧬 Validación de Dependencias Dinámicas
El motor ahora incluye el operador `DEPENDENCIAS_CERRADAS`. Esto permite, por ejemplo, impedir el cierre de un Desvío si existen CAPAs asociadas que aún permanecen abiertas, consultando las relaciones en tiempo real sin necesidad de columnas auxiliares con fórmulas.

### 🧹 Limpieza Quirúrgica (Preservación de Inteligencia)
El motor de limpieza diferencia entre datos transaccionales y datos de sesión. Al finalizar un registro, limpia el formulario pero preserva el **ID** y el **Usuario** por persistencia visual, enviando valores controlados a la base de datos para no corromper columnas calculadas o fórmulas nativas de Excel.

---

## 🚀 Roadmap de Evolución

*   **v3.0 (Actual):** Arquitectura Excel-Centric, Hashes de seguridad y validación de dependencias.
*   **v3.1 (Próximo):** Script administrativo de reseteo de claves con re-sellado automático de integridad.
*   **v4.0 (Visión):** Dashboard nativo de monitoreo de salud del sistema (Integrity Audit) consumiendo los logs SHA-256.

---

## 👨‍💻 El Rigor detrás del Código

Este sistema sigue los principios de la **Rigurosidad Pragmática**:
*   **SESE (Single Entry, Single Exit):** Código predecible y auditable.
*   **Manejo Híbrido de Errores:** Los errores de negocio informan al usuario; los de infraestructura protegen el sistema.
*   **Zero Silent Failures:** Si algo falla, el sistema se bloquea y avisa. No hay lugar para la ambigüedad.

---
*Desarrollado bajo estándares **GMP** (Good Manufacturing Practices) para garantizar que el dato sea el activo más seguro de la compañía.*