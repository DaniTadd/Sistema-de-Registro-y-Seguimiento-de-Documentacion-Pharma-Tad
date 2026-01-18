# 🏭 Sistema de Gestión de Desvíos (GMP) - V1.0 Stable

Sistema automatizado para el registro, seguimiento y auditoría de desvíos en entornos regulados (GMP/BPF). Desarrollado sobre **Excel Online** utilizando **Office Scripts (TypeScript)** para garantizar la integridad de datos, seguridad de acceso y trazabilidad completa (Audit Trail).

## 📋 Descripción Técnica y Arquitectura
El sistema utiliza una arquitectura de "Frontend" controlado en Excel que se comunica con una base de datos protegida mediante scripts de servidor.

## 💡 ¿Por qué Excel Web + Office Scripts?
La elección de esta stack tecnológica se basa en tres pilares estratégicos:

1.  **Curva de Aprendizaje Nula:** Aprovechamos la familiaridad universal con la interfaz de Excel para que los usuarios finales interactúen con un entorno que ya dominan, facilitando la adopción del sistema.
2.  **Soberanía y Eficiencia de Costos:** El sistema utiliza herramientas estándar de Office/OneDrive, eliminando la necesidad de infraestructura dedicada (servidores o bases de datos externas) y costos operativos adicionales de mantenimiento.
3.  **Portabilidad y Despliegue Inmediato:** Al ser una solución basada en la nube, el sistema es accesible desde cualquier navegador, garantizando que la lógica de validación (Office Scripts) se ejecute de forma centralizada sin necesidad de instalaciones locales.

* **Mapeo Dinámico (Label-Matching):** A diferencia de scripts convencionales, este sistema localiza los datos mediante etiquetas en la columna B del formulario. Esto permite modificar el diseño visual del Excel sin romper la lógica del código.
* **Seguridad ALCOA+:** Implementación de principios de integridad de datos. No se permiten registros anónimos ni modificaciones sin justificación (Motivo de Cambio obligatorio).
* **Validación de Estados:** Blindaje lógico que impide la edición de registros con estado "CERRADO".

---

## 🚀 Roadmap de Desarrollo (Evolución del Sistema)

### ✅ Versión 1.0: El Núcleo (Core) - *ESTABILIZADO*
* **Registrar/Buscar/Actualizar:** Módulos base con validación ALCOA+ y mapeo dinámico.
* **Audit Trail:** Historial de cambios con Delta Logging y formato horario de 24hs.

### 🚧 Versión 1.1: Gestión de Impacto & Cierre (En Desarrollo)
* **Módulo de Cierre:** Script `Cerrar Desvio.ts` para transicionar el estado a "CERRADO", activando el bloqueo de edición GMP.
* **Módulo Acciones (CAPA):** Gestión de tareas correctivas/preventivas con seguimiento independiente.
* **Módulo Afectaciones (Lotes):** Vinculación N:1 para identificar materiales impactados.

### 📊 Versión 1.2: Contexto e Investigación (Analítica & BI)
* **Módulo RCA (Root Cause Analysis):** Tabla independiente de atributos (Equipo, Turno, Área, condiciones ambientales) vinculada por ID.
* **Preparación para Power BI:** Este diseño relacional permite el consumo directo desde herramientas de Business Intelligence para la detección de patrones críticos, análisis de Pareto y visualización de tendencias de causa raíz.

### 🔮 Versión 2.0: Seguridad & Automatización (QA Interno)
* **Identidad de Usuario:** Captura de identidad de Azure AD mediante Power Automate para firmas digitales auténticas.

### 📝 Versión 3.0: Ecosistema de Reporte en Planta (MS Forms)
* **Captura Externa:** Apertura a otros sectores para reportes rápidos desde planta.
* **Módulo de Triaje:** Revisión y validación de QA antes del ingreso formal a la base principal.

### 📂 Versión 4.0: Gestión de Evidencias (Alta Complejidad)
* **Módulo de Archivos:** Investigación de integración para la creación de carpetas automáticas y vinculación de sustento documental (Fotos/PDFs) a cada registro.

---

## 🛠️ Configuración y Seguridad

1.  **Puente de Sincronización:** El desarrollo se realiza localmente en VS Code y se sincroniza mediante un script de Python (`puente.py`) hacia OneDrive.
2.  **Protección de Datos:**
    * Las hojas de Base de Datos e Historial están protegidas por contraseña, gestionada de forma centralizada desde una celda oculta en la hoja `MAESTROS`.
    * Uso de bloques `try-catch-finally` para asegurar que las hojas se vuelvan a proteger automáticamente incluso si el script falla.

## 🔒 Notas de Privacidad y Seguridad
* **Protección de Rutas:** El archivo `config.py` está excluido del control de versiones (`.gitignore`) ya que contiene rutas de directorios locales.
* **Implementación:** Se provee un archivo `config.example.py` como plantilla. Para implementar el sistema, se debe renombrar a `config.py` y configurar la ruta local hacia la carpeta de sincronización de OneDrive.
