# 🏭 Sistema de Gestión de Desvíos (GMP)

Sistema automatizado para el registro, seguimiento y auditoría de desvíos en entornos regulados (GMP/BPF). Desarrollado sobre **Excel Online** utilizando **Office Scripts (TypeScript)** para garantizar la integridad de datos, seguridad de acceso y trazabilidad completa (Audit Trail).

## 📋 Descripción Técnica
El sistema reemplaza la carga manual en planillas inseguras por una interfaz controlada (Frontend en hoja `INPUT`) que escribe en bases de datos protegidas (Backend en hojas `BD_`).

* **Stack:** Excel Online (Business) + Office Scripts (TypeScript).
* **Seguridad:** Bloqueo de celdas automático, gestión de contraseñas centralizada y validación estricta de tipos (`strict: true`).
* **Despliegue:** Sincronización local-nube mediante puente Python (`puente.py`).

---

## 🚀 Roadmap de Desarrollo

### ✅ Versión 1.0: El Núcleo (Core) - *ESTADO ACTUAL*
El objetivo de esta versión es garantizar la carga segura y la integridad referencial de los desvíos principales.

* **[x] Registrar Desvío:**
    * Validación lógica de fechas (Suceso vs Registro vs QA).
    * Control de campos obligatorios.
    * Generación automática de ID incremental (concurrencia básica).
    * Escritura en `BD_DESVIOS`.
* **[x] Buscar Desvío:**
    * Lectura en memoria (`getValues`) para optimizar rendimiento.
    * Carga de datos en formulario `INPUT` para visualización.
* **[x] Actualizar Desvío (Audit Trail):**
    * Sistema de **Delta Logging**: Solo se guardan los campos que cambiaron.
    * Obligatoriedad de "Motivo de Cambio" (GMP).
    * Registro histórico inmutable en `HISTORIAL_DESVIOS`.
* **[x] UX/UI:**
    * Auto-focus en mensajes de estado (Scroll automático).
    * Feedback visual con colores (Éxito/Error).
    * Limpieza automática de formulario.

### 🚧 Versión 1.1: Gestión de Impacto (En Progreso)
Expansión del núcleo para incluir el detalle granular de lotes afectados y acciones correctivas.

* **[ ] Módulo Afectación (Lotes):**
    * Script `Agregar Afectacion.ts` para vincular N lotes a 1 desvío.
    * Tablas dedicadas: `BD_AFECTACION` e `HISTORIAL_AFECTACION`.
* **[ ] Módulo Acciones (CAPA):**
    * Asignación de tareas correctivas/preventivas.
    * Seguimiento de responsables y fechas límite.
* **[ ] Pruebas Integrales:** Validación de flujo completo (Alta -> Afectación -> Acción -> Cierre).

### 🔮 Versión 2.0: Seguridad Empresarial & Automatización (Futuro)
Migración de la lógica de seguridad y notificaciones a la capa de infraestructura de Microsoft 365.

* **[ ] Identidad Infalsificable:**
    * Reemplazo de botones directos por flujos de **Power Automate**.
    * Captura del "Usuario de Ejecución" (Active Directory) para evitar suplantación en la celda de firma.
* **[ ] Control de Acceso (RBAC):**
    * Lista de usuarios permitidos en hoja `MAESTROS`.
    * Power Automate como "Portero" que valida permisos antes de ejecutar el script.
* **[ ] Notificaciones:**
    * Envío automático de mails a Calidad ante desvíos críticos.

---

## 🛠️ Configuración Local (Para Desarrolladores)

Este proyecto utiliza un puente de sincronización para permitir el desarrollo en VS Code local.

1.  **Requisitos:** Python 3.x, cuenta de OneDrive Business sincronizada.
2.  **Configuración:**
    * Renombrar `config.example.py` a `config.py`.
    * Establecer la `RUTA_ONEDRIVE_REAL` apuntando a la carpeta de Scripts de Excel en local.
3.  **Sincronización:**
    * Ejecutar `python puente.py`.
    * El script detectará cambios en los archivos `.ts` y actualizará los `.osts` en OneDrive automáticamente.

## 🔒 Notas de Seguridad
* **No subir `config.py`:** Contiene rutas locales.
* **No subir `.xlsx`:** Los datos de prueba deben permanecer locales.
* **Gestión de Claves:** La contraseña de protección se administra dinámicamente desde la hoja de configuración `MAESTROS` (evitando hardcoding de la contraseña real en los scripts).