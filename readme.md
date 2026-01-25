# 🏭 Sistema de Registro y Seguimiento de Documentación (GMP)

![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)
![Tech](https://img.shields.io/badge/Office_Scripts-TypeScript-3178C6.svg)
![Standard](https://img.shields.io/badge/Compliance-GMP%20%2F%20ALCOA%2B-green.svg)

Sistema automatizado para el registro, seguimiento y auditoría de documentación de calidad (Desvíos, Reclamos, CC) en entornos regulados. Desarrollado sobre **Excel Online** para garantizar la integridad de datos y trazabilidad completa (Audit Trail) sin costos de infraestructura.

---

## 📋 Propósito del Sistema
Este motor transforma una planilla estándar en una aplicación segura, permitiendo a la industria farmacéutica y alimentaria gestionar incidentes de calidad cumpliendo con normativas de integridad de datos, pero con la facilidad de uso de Excel.

### Funcionalidades Clave
* ✅ **Formularios Inteligentes:** Validación automática de datos y campos obligatorios.
* ✅ **Ciclo de Vida:** Flujo de estados controlado (Abierto → Cerrado ↔ Reabierto → Anulado).
* ✅ **Audit Trail Inmutable:** Registro automático de *quién, cuándo y qué* se modificó.
* ✅ **Seguridad Robusta:** El sistema se autoprotege ante errores, garantizando que las hojas nunca queden expuestas.

---

## ⚖️ ¿Por qué Excel + TypeScript?

Elegimos esta combinación para reemplazar tecnologías obsoletas (como Access o VBA local) y evitar la complejidad de servidores dedicados.

### 1. Adopción Inmediata (UI Familiar)
El usuario trabaja en un entorno que ya domina (Excel), eliminando la resistencia al cambio y la necesidad de capacitaciones costosas sobre nuevas interfaces.

### 2. Infraestructura Zero (Sin Gastos Adicionales)
Eliminamos la necesidad de contratar servidores, pagar licencias de bases de datos (SQL) o adquirir software de terceros. El sistema utiliza los recursos **ya incluidos** en cualquier licencia comercial estándar de **Microsoft 365**.
> **Impacto Económico:** Si la organización ya cuenta con Office 365, el costo de infraestructura para implementar y mantener este sistema es **$0**.

### 3. Stack Moderno (Cloud vs. Local)
A diferencia de las macros viejas (VBA), **Office Scripts** corre en la nube. Esto permite ejecutar el sistema desde cualquier navegador o dispositivo (PC, Tablet) sin bloquear los archivos y sin riesgos de virus de macro locales.

---

## 🌟 Diferenciales de Diseño

* **Flexibilidad Total:** Gracias a un sistema de "Mapeo Dinámico", se pueden agregar nuevos campos (ej. "Turno", "Temperatura") directamente en la hoja de Excel sin necesidad de tocar el código.
* **Código Seguro:** Las credenciales y contraseñas no están en el código. Utilizamos un sistema de punteros internos (`Named Items`) para mantener la seguridad incluso si se comparte el repositorio.

---

## 🚀 Roadmap

* **v1.0 (Actual):** Registro, Búsqueda, Actualización, Audit Trail y Gestión de Estados.
* **v2.0 (Próximamente):**
    * **Módulo CAPA:** Gestión de Tareas y acciones correctivas.
    * **Impacto y Afectaciones:** Vinculación de Lotes, Equipos y Materias Primas.
    * **Contexto Analítico:** Captura de atributos extendidos para facilitar el análisis de causa raíz (RCA).
* **v3.0 y 4.0 (Futuro):**
    * **Ecosistema Integrado:** Identidad de usuario vía Azure AD y captura remota con Microsoft Forms.
    * **Inteligencia de Datos:** Tableros de control avanzados en Power BI.

---

## 🛠️ Instalación y Despliegue

Este sistema (en su estado actual) requiere una estructura específica en el libro de Excel para funcionar.

### Paso 1: Preparación del Libro (Schema)
Antes de cargar los scripts, el archivo Excel debe tener la siguiente estructura:

1.  **Hojas Requeridas:** Crear 4 hojas llamadas exactamente: `INPUT_DESVIOS`, `BD_DESVIOS`, `HISTORIAL_DESVIOS`, `MAESTROS`.
2.  **Tablas de Datos:**
    * En `BD_DESVIOS`: Insertar una tabla llamada **`TablaDesvios`**.
    * En `HISTORIAL_DESVIOS`: Insertar una tabla llamada **`TablaHistorialDesvios`**.
    * En `MAESTROS`: Insertar una tabla llamada **`TablaReglas`**.
3.  **Configuración de Seguridad:**
    * Crear un **Nombre Definido** (Fórmulas > Administrador de Nombres) llamado `SISTEMA_CLAVE` que apunte a una celda con la contraseña maestra.

### Paso 2: Carga de la Lógica (Scripts)
* **Requisito:** Licencia Microsoft 365 Business (Basic o superior).
* **Opción Manual:** En la pestaña **Automatizar** de Excel, crear un **Nuevo Script** para cada archivo de la carpeta `/src`, pegar el código y guardarlo con el nombre exacto (ej. `Registrar Desvio`).
* **Opción Dev:** Ejecutar el script `tools/puente.py` para sincronizar automáticamente los archivos locales con la carpeta de Office Scripts en OneDrive (⚠️ requiere haber creado previamente en OneDrive los archivos `.osts` vacíos con el mismo nombre exacto).

> 📘 **Documentación Técnica:**
> Para el detalle exacto de las columnas requeridas en cada tabla y la lógica interna, consultar la [**Memoria Técnica**](./MEMORIA_TECNICA.md).

---

*Desarrollado con foco en GMP (Good Manufacturing Practices) y Data Integrity.*