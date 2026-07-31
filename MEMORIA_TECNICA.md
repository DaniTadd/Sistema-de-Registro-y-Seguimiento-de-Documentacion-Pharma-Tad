# 📚 Manual de Lógica y Memoria Técnica: SGC-Engine v3.0 (Excel-Centric Edition)

**Versión:** 3.0 (Arquitectura de Alta Integridad)
**Tecnología:** Office Scripts (TypeScript)
**Estándar:** Compliance GMP / ALCOA+ (Data Integrity)
**Arquitectura:** Motor Transaccional de Estado Sólido (SESE)

---

## 1. Filosofía de Ingeniería y Arquitectura

El sistema ha evolucionado hacia un modelo **Excel-Centric**, eliminando dependencias de flujos externos para reducir la latencia y maximizar el control de integridad en el punto de uso.

### A. Mapeo Semántico Dinámico
El motor opera mediante una capa de abstracción que desacopla la interfaz de usuario de la base de datos física:
*   **Navegación por Etiquetas:** El script no utiliza referencias fijas (ej. "B5"). Realiza un escaneo dinámico de etiquetas en el formulario, normaliza los textos y busca su paridad exacta en los encabezados de las tablas.
*   **Nomenclatura Técnica:** Se mantiene un estándar descriptivo para garantizar que el código sea auditable:
    *   `objetoDatosFormulario`: Buffer de datos capturados y sanitizados.
    *   `encabezadosTabla`: Estructura lógica de destino.

### B. Seguridad Criptográfica Zero-Trust
Se implementa una capa de seguridad basada en estándares industriales para la protección de la identidad:
*   **Firmas SHA-256:** Las contraseñas de usuario no se almacenan en texto plano. El sistema utiliza hashes criptográficos irreversibles para validar la identidad.
*   **Protocolo "INICIAR":** Las credenciales nuevas o reseteadas utilizan la bandera `INICIAR`, obligando al operario a establecer una firma digital única en su primera transacción.

### C. Validación de Integridad en Vivo (ALCOA+)
Antes de cualquier transacción (Escritura/Edición), el motor ejecuta un control de salud estructural:
*   **Live Hash Check:** El script calcula el SHA-256 de las tablas en tiempo real y lo compara contra el sello maestro en `TablaIntegridad`.
*   **Bloqueo Catastrófico:** Si se detecta una alteración manual de los datos (violación de integridad), el sistema bloquea cualquier operación y solicita la intervención de Calidad.

---

## 2. Estructura de Datos y Persistencia

### 2.1 Metadatos de Auditoría Inmutables
Campos gestionados estrictamente por el motor para garantizar la trazabilidad:

| Campo | Función | Comportamiento |
| :--- | :--- | :--- |
| **ID** | Identificador Único | Prefijo dinámico + Máximo correlativo + 1. |
| **ESTADO** | Ciclo de Vida | ABIERTO / CERRADO / ANULADO. |
| **AUDIT_TRAIL** | Timestamp | Registro temporal generado por el servidor (GMT-3). |
| **USUARIO** | Firma Digital | ID del usuario autenticado mediante Hash. |
| **MOTIVO** | Justificación | Parámetro obligatorio para cumplir con la intención del registro. |
| **CAMBIOS** | Log de Auditoría | Detalle de diferencias: `[Campo]: [V. Anterior] -> [V. Nuevo]`. |

### 2.2 Normalización de Fechas
Para evitar falsos positivos en el Audit Trail, el sistema implementa una normalización cronológica doble:
1.  **Input:** Acepta tanto números seriales de Excel como strings en formato `DD/MM/YYYY`.
2.  **Comparación:** Estandariza ambos valores a un formato común antes de evaluar si existe un cambio real de datos.

---

## 3. Protocolos de Operación Segura

### 3.1 Cambio de Estado y Dependencias
El sistema permite el cierre condicionado de registros mediante la regla `DEPENDENCIAS_CERRADAS`:
*   **Integridad Referencial:** Antes de cerrar una entidad "Madre", el motor escanea las tablas "Hijas" (ej. CAPAs) y bloquea la acción si detecta dependencias abiertas.

### 3.2 Double-Check de Anulación
La anulación es una acción terminal que requiere una confirmación de seguridad extendida:
*   **Validación de Mismatch:** El usuario debe ingresar manualmente el ID del registro que desea anular. El sistema aborta si el ID ingresado no coincide con el registro activo en pantalla.

---

## 4. Matriz de Solución de Problemas (Troubleshooting)

| Mensaje de Error | Canal | Causa y Acción Correctiva |
| :--- | :---: | :--- |
| **"ERROR: Integridad violada en..."** | ⛔ Sistema | Alteración manual de la base de datos detectada. El sistema se bloquea por seguridad ALCOA+. |
| **"Firma inválida: Credenciales..."** | ⚠️ Feedback | La clave ingresada no coincide con el Hash del usuario. |
| **"Mismatch de Seguridad: ID..."** | ⚠️ Feedback | Error en el Double-Check. El ID tipeado no coincide con el del formulario. |
| **"Integridad violada en TablaUsuarios"** | ⛔ Sistema | Alguien intentó modificar las claves de seguridad manualmente. |
| **"AccessDenied / SISTEMA_CLAVE"** | ⛔ Sistema | Fallo al recuperar la clave de protección del Administrador de Nombres. |

---

## 5. Gestión de Infraestructura (Administración)

*   **Reset de Claves:** El administrador puede resetear la contraseña de un operario escribiendo `INICIAR` en el campo de Hash de la `TablaUsuarios`. Esto debe realizarse mediante el script de administración para actualizar también la `TablaIntegridad`.
*   **Protección (SafeProtect):** Al finalizar cada script, el bloque `finally` garantiza que todas las hojas (Datos, Historial, Usuarios e Integridad) queden protegidas bajo el secreto de `SISTEMA_CLAVE`.

---
*Documento de carácter confidencial. Diseñado para asegurar el cumplimiento de la **CFR 21 Part 11** en entornos de hojas de cálculo.*