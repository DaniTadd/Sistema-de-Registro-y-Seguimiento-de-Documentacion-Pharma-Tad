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

## 🚀 Roadmap de Desarrollo

### ✅ Versión 1.0: El Núcleo (Core) - *ESTABILIZADO*
Estado actual del sistema enfocado en la integridad referencial y auditoría.

* **Registrar Desvío:**
    * Generación de ID incremental automático.
    * Validación de cronología lógica: $Fecha\ Suceso \le Fecha\ Registro \le Fecha\ QA$.
    * Sellado de tiempo (**Audit Trail**) forzado en formato 24hs (es-AR) para eliminar ambigüedad AM/PM.
* **Buscar Desvío:**
    * Carga dinámica de datos en el formulario mediante mapa de lectura.
    * **Firma Forzada:** El buscador limpia el campo "Usuario" intencionalmente para obligar al operador actual a identificarse antes de actualizar.
* **Actualizar Desvío & Historial:**
    * **Delta Logging:** El sistema compara el valor viejo vs. nuevo y genera un log detallado: `[Campo: Valor Viejo -> Valor Nuevo]`.
    * **Traducción de Fechas:** Conversión de formatos seriales de Excel a fechas legibles para humanos en el historial de cambios.
    * **Gestión de Opcionales:** Soporta campos opcionales como `FECHA QA` sin romper las reglas de integridad de otros campos obligatorios.

### 🚧 Próximos Pasos (Evolución del Sistema)

1. **Módulo Acciones (CAPA):** Desarrollo de la relación 1:N para gestionar tareas correctivas y preventivas con seguimiento de estados independientes.
2. **Módulo Afectaciones (Lotes/Productos):** Implementación de una tabla relacional para vincular múltiples materiales impactados a un único registro de desvío.
3. **Módulo de Contexto e Investigación (RCA):** * Creación de una tabla independiente de atributos contextuales (Equipo, Turno, Área, condiciones ambientales).
    * Este diseño permite la expansión de variables de investigación sin alterar la estructura de la base de datos principal, facilitando el análisis de tendencias y causa raíz.

---

## 🛠️ Configuración y Seguridad

1.  **Puente de Sincronización:** El desarrollo se realiza localmente en VS Code y se sincroniza mediante un script de Python (`puente.py`) hacia OneDrive.
2.  **Protección de Datos:**
    * Las hojas de Base de Datos e Historial están protegidas por contraseña, gestionada de forma centralizada desde una celda oculta en la hoja `MAESTROS`.
    * Uso de bloques `try-catch-finally` para asegurar que las hojas se vuelvan a proteger automáticamente incluso si el script falla.

## 🔒 Notas de Privacidad y Seguridad
* **Protección de Rutas:** El archivo `config.py` está excluido del control de versiones (`.gitignore`) ya que contiene rutas de directorios locales.
* **Implementación:** Se provee un archivo `config.example.py` como plantilla. Para implementar el sistema, se debe renombrar a `config.py` y configurar la ruta local hacia la carpeta de sincronización de OneDrive.
