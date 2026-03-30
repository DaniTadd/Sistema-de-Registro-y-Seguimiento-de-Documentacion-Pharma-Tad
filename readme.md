# 🏭 SGC-Engine: Enterprise-Grade Quality Management on a Spreadsheet

![Version](https://img.shields.io/badge/version-2.0.0-blue.svg?style=for-the-badge)
![Tech](https://img.shields.io/badge/Office_Scripts-TypeScript-3178C6.svg?style=for-the-badge)
![Compliance](https://img.shields.io/badge/Compliance-ALCOA%2B%20%2F%20GMP-orange.svg?style=for-the-badge)
![Cost](https://img.shields.io/badge/Infrastructure-Zero_Cost-green.svg?style=for-the-badge)

## 💎 Visión Estratégica: El Dato como Activo Crítico

El **SGC-Engine** no es solo una hoja de cálculo; es un **Motor Universal de Gestión** diseñado para transformar el registro de calidad de una tarea administrativa pesada en un pilar estratégico de inteligencia de negocio. 

En entornos regulados (GMP), la integridad de datos suele verse como una carga burocrática. Este sistema rompe ese paradigma basándose en tres ejes:

1.  **Escalabilidad Nativa:** Diseñado con lógica de base de datos relacional, permitiendo una transición transparente hacia **Power Apps o Dataverse** cuando el negocio lo requiera.
2.  **Data-Ready (BI):** Estructura de "Esquema en Estrella" que permite conectar **Power BI** en segundos, sin necesidad de limpiezas previas.
3.  **Rigor Normativo:** Implementación técnica de los principios **ALCOA+**, garantizando que cada dato sea Atribuible, Legible, Contemporáneo, Original y Exacto.

---

## ⚖️ El Business Case: ¿Por qué Excel + Office Scripts?

Muchas organizaciones invierten miles de dólares en software de nicho que termina subutilizado por su complejidad. El SGC-Engine propone un camino disruptivo:

### 1. Inversión Cero ($0 USD)
Si tu organización ya paga licencias de **Microsoft 365**, ya tenés todo lo que necesitás. Aprovechamos el stack tecnológico existente para eliminar costos de servidores, hosting o licencias adicionales de software propietario.

### 2. Adopción con Resistencia Cero (User-Centric)
El mayor costo de un sistema nuevo es la capacitación. Al utilizar la interfaz de Excel —un entorno que los usuarios ya dominan y en el que confían— eliminamos la fricción del cambio y reducimos drásticamente los errores de carga.

### 3. Stack Moderno (Cloud-Native)
A diferencia de las antiguas macros VBA (locales y monousuario), este motor utiliza **Office Scripts (TypeScript)**:
* **Co-autoría:** Varios usuarios pueden operar el sistema simultáneamente en la nube.
* **Seguridad:** Ejecución segura en los servidores de Microsoft, eliminando riesgos de virus por macro y facilitando el acceso desde cualquier dispositivo (PC, Tablet, Web).

---

## 🛠️ Fortalezas y Diferenciales de Diseño

### 🧬 Arquitectura "Madre-Hija" (Modularidad)
El sistema funciona como un ADN funcional. La "Entidad Madre" (ej. Desvíos) provee la lógica que se replica instantáneamente para "Entidades Hijas" (ej. CAPAs, Afectaciones), permitiendo una trazabilidad total y análisis de causa raíz (RCA) con rigor estadístico.

### 🛡️ Integridad Quirúrgica de Datos
* **Salvaguarda de Identidad:** El sistema incluye un motor de validación que impide actualizar registros si existe un "mismatch" entre el panel de control y la base de datos.
* **Protección de Fórmulas:** El script detecta y preserva las columnas de cálculo automático de Excel, enviando valores `null` donde sea necesario para no romper la inteligencia nativa de la hoja.
* **Audit Trail Real:** Registro inmutable de *Quién*, *Qué*, *Cuándo* y *Por qué* (Motivo obligatorio de cambio).

### ⚖️ Motor de Reglas Dinámico
Permite al administrador configurar validaciones complejas (ej: "Fecha de cierre no puede ser anterior a apertura") directamente desde una tabla de Excel, sin necesidad de escribir una sola línea de código adicional.

---

## 🚀 Roadmap de Evolución

* **v1.0 (Consolidado):** Gestión completa de la Entidad Madre (Registro, Búsqueda, Actualización).
* **v2.0 (Actual):** * Lanzamiento de **Entidades Hijas** (CAPAs/Afectaciones) con integridad cruzada.
    * Regla `ESTA_ABIERTO` para bloqueo de jerarquías.
* **v3.0 (Visión):**
    * **Analytics Avanzado:** Tableros nativos en Power BI consumiendo el motor.
    * **Captura Móvil:** Integración directa con Power Automate Forms.

---

## ⚙️ Instalación y Despliegue Rápido

1.  **Infraestructura:** Crear las hojas de proceso (Input, DB, Historial) y la hoja de Maestros.
2.  **Seguridad:** Definir el rango `SISTEMA_CLAVE` con tu contraseña de protección.
3.  **Carga:** Copiar los archivos de `/src` al editor de Office Scripts en Excel Online.
4.  **Sincronización:** (Opcional) Usar el puente de Python incluido para mantener tu repositorio Git sincronizado con OneDrive.

> Para un análisis profundo de la arquitectura, consultar la [**Memoria Técnica**](./MEMORIA_TECNICA.md).

---

## ⚖️ Transparencia y Limitaciones
* **Escalabilidad:** Optimizado para registros de hasta **10,000 filas** (debido a los tiempos de ejecución de Office Scripts).
* **Ecosistema:** Requiere Microsoft 365 (Web) para habilitar todas las funciones de co-autoría y automatización.

---

## 👨‍💻 Detrás del Proyecto: El Camino del Aprendizaje

Este sistema nació como un desafío personal de ingeniería: **¿Es posible llevar a Excel al límite de sus capacidades para que se comporte como un ERP profesional?**

**Lo que me llevo de este desarrollo:**
* **Arquitectura de Software en la Nube:** Dominio de Office Scripts y su interacción con el ecosistema Power Platform.
* **Mentalidad Data Integrity:** Aplicación práctica de normativas ALCOA+ en el diseño de bases de datos.
* **Resolución de Problemas:** Crear un puente en Python para gestionar versiones de Office Scripts (que nativamente no tiene integración con Git) fue un hito clave en mi aprendizaje de automatización de flujos de trabajo.

Este proyecto demuestra que con la arquitectura correcta, las herramientas que ya tenemos en el escritorio pueden ser el motor de transformación digital de una compañía.

---
*Desarrollado con foco en **GMP** (Good Manufacturing Practices) e **Integridad de Datos**.*