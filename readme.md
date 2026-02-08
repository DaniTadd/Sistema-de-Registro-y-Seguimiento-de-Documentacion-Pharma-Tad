# 🏭 SGC-Engine: Motor Universal de Gestión de Calidad (GMP)

![Version](https://img.shields.io/badge/version-1.0.0-blue.svg?style=for-the-badge)
![Tech](https://img.shields.io/badge/Office_Scripts-TypeScript-3178C6.svg?style=for-the-badge)
![Compliance](https://img.shields.io/badge/Compliance-ALCOA%2B%20%2F%20GMP-orange.svg?style=for-the-badge)

## 💎 Visión Estratégica: El Dato como Activo

> **Cultura de la Información:** Del registro de calidad como tarea administrativa a un pilar estratégico. En este sistema, la integridad de datos trasciende el cumplimiento normativo para convertirse en un motor de la inteligencia y la toma de decisiones de la compañía.

1. **Escalabilidad Garantizada:** El diseño de tablas normalizadas facilita una transición fluida hacia **Power Apps / Dataverse**, funcionando como un prototipo funcional de alta fidelidad.
2. **Data-Ready para BI:** La estructura de "Esquema en Estrella" asegura que la información sea consumible de inmediato por herramientas de **Business Intelligence (Power BI)** sin necesidad de limpieza previa.
3. **Análisis de Valor (RCA):** La trazabilidad entre entidades (Madre-Hija) permite realizar análisis de causa raíz y tendencias con rigor estadístico, transformando el cumplimiento normativo en inteligencia de negocio.

---

## ⚖️ ¿Por qué Excel Online + Office Scripts? (Business Case)

### 1. Infraestructura Zero ($0 Inversión)
Si su organización ya posee licencias de **Microsoft 365**, el costo de infraestructura es **cero**. Se eliminan gastos de servidores SQL, hosting web o licencias de software de nicho ni consultoría especializada en IT.

### 2. Adopción con Resistencia Cero
Los usuarios no necesitan aprender a usar un software nuevo. La interfaz es Excel, un entorno que ya dominan. Esto reduce reduce drásticamente el tiempo de capacitación, la resistencia al cambio y los errores de carga comparado con la implementación de un nuevo software propietario.

### 3. Stack Moderno y Co-autoría
A diferencia de las macros VBA antiguas, este motor corre en la nube. Permite que múltiples usuarios operen el sistema simultáneamente (Co-autoría) desde cualquier dispositivo, garantizando seguridad y disponibilidad 24/7.


---

## 🛠️ Fortalezas del Sistema

### Stack Moderno (Cloud vs. Local)
A diferencia de las macros tradicionales (VBA), **Office Scripts** se ejecuta en la nube de Microsoft. Esto aporta ventajas críticas para el entorno corporativo:
* **Co-autoría Real:** Varios usuarios pueden editar el archivo simultáneamente mientras los scripts se ejecutan, sin bloqueos de lectura/escritura.
* **Multiplataforma:** El sistema funciona en Excel Online desde cualquier navegador o dispositivo (PC, Tablet), eliminando la dependencia de instalaciones locales.
* **Seguridad:** Al no utilizar archivos `.xlsm`, se mitigan los riesgos de virus por macro y se facilita la distribución segura del libro.



### 🌟 Diferenciales de Diseño
* **Flexibilidad Total:** Gracias a un sistema de **"Mapeo Dinámico"**, es posible agregar nuevos campos (ej. "Turno", "Temperatura") directamente en la hoja de Excel sin necesidad de modificar una sola línea de código.
* **Arquitectura de Datos (Star Schema):** El sistema utiliza un enfoque de **Esquema en Estrella** donde la Base de Datos centraliza los hechos (registros), mientras que las hojas de Maestros y Reglas actúan como dimensiones. Esta organización garantiza que los datos sean robustos, normalizados y fáciles de exportar a herramientas de Power BI.

### 🛡️ Integridad de Datos (ALCOA+)
Diseñado bajo principios de cumplimiento normativo:
* **Audit Trail:** Registro inmutable de *Quién* cambio *Qué* y *Cuándo* lo hizo.
* **Firmas Electrónicas:** Captura de usuario y motivo de cambio obligatorios.
* **Seguridad:** Bloqueo automático de registros cerrados o anulados.

### 🧬 Arquitectura "Madre-Hija"
El sistema está diseñado como un motor genérico. La "Entidad Madre" (Ej: Desvíos) provee el ADN funcional que puede ser replicado instantáneamente para otras entidades (Ej: Controles de cambios) y la "Entidad hija" (Ej: CAPAs, Afectaciones, etc.), para los registros dependientes del principal (ejemplo, CAPAs de un Desvío, Acciones de un Control de Cambios, etc.).

### 💉 Compromiso Quirúrgico de Datos
A diferencia de otros scripts, el **SGC-Engine** protege sus fórmulas nativas. 
* Solo sobrescribe las celdas que el usuario modifica explícitamente.
* Respeta las columnas de cálculo automático, permitiendo indicadores en tiempo real dentro de la base de datos sin riesgo de borrado accidental.


### ⚖️ Motor de Reglas Dinámico
Permite configurar validaciones de negocio (ej: "La fecha de cierre no puede ser menor a la de apertura") directamente desde una tabla en Excel, sin tocar una sola línea de código.



---


## 🚀 Roadmap de Evolución

* **v1.0 (Actual):** Módulo de Desvíos consolidado. Registro, Búsqueda, Actualización y Audit Trail.
* **v2.0 (En Desarrollo):** * **Entidades Hijas:** Lanzamiento de módulos de **Afectaciones** y **CAPAs** vinculados a la entidad madre.
    * **Integridad Cruzada:** Validación de códigos de producto/lote contra maestros globales.
* **v3.0 (Visión):**
    * **Analytics:** Tableros de control nativos en Power BI consumiendo la data estructurada del motor.
    * **Cloud Forms:** Captura remota de datos desde dispositivos móviles.

---

## ⚙️ Instalación y Requisitos

### Requisitos Técnicos
* Cuenta de **Microsoft 365 Business** (Basic, Standard o Premium).
* Excel Online habilitado para **Office Scripts**.


## 🛠️ Instalación y Despliegue

### Pasos Rápidos para Implementación
1.  **Esquema de Hojas:** Crear las hojas `INPUT_MADRE`, `BD_MADRE` y `HIST_MADRE`.
2.  **Esquema de Tablas:** Crear las tablas con los siguientes nombres exactos:
    * `TablaMadre` (en la hoja `BD_MADRE`).
    * `TablaHistorialMadre` (en la hoja `HIST_MADRE`).
    * `TablaReglas` (en la hoja `MAESTROS`).
3.  **Seguridad:** Definir un **Nombre de Rango** llamado `SISTEMA_CLAVE` que apunte a la celda que contiene la contraseña de protección de hojas.
4.  **Carga de Scripts:** Copiar el contenido de la carpeta `/src` al editor de Office Scripts en Excel Online.
5.  **Configurar Rangos:** Ejecutar el script `Configurar Rangos` para mapear automáticamente el formulario.

* Para más instrucciones, visitar la [**Memoria Técnica**](./MEMORIA_TECNICA.md).o.

--

## ⚖️ Limitaciones y Transparencia
* **Volumen de Datos:** El sistema es ideal para registros de hasta **10,000 filas**. Esta limitación responde a los tiempos de ejecución (Timeout) de 120 segundos de Office Scripts y al impacto en el rendimiento de Excel Online al procesar grandes volúmenes de datos en memoria. Para escalas mayores, el motor está diseñado para facilitar una transición futura hacia Power Apps/Dataverse.
* **Entorno:** Diseñado exclusivamente para el ecosistema Microsoft 365 (Web).

---
*Desarrollado con foco en **GMP** (Good Manufacturing Practices) y **Data Integrity**.*