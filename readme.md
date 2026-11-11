# 🔬 QA-ALCOA-Engine: Framework Transaccional para Excel Online (Office Scripts)

Un motor de gestión de registros de calidad y trazabilidad documental (ALCOA+) construido enteramente sobre Microsoft Excel Online, Office Scripts (TypeScript) y Power Automate.

Este proyecto es una prueba de concepto que demuestra cómo es posible inyectar rigor transaccional, criptografía SHA-256 y reglas de negocio relacionales en una plataforma de ofimática genérica, utilizando herramientas nativas de Microsoft 365 sin costos de licencias adicionales.

---

## 💡 Motivación y Criterios de Elección

En la industria farmacéutica y entornos regulados (normas GMP / ISO), los sistemas de gestión de calidad (QMS) suelen ser monolíticos, costosos y lentos de adaptar. Por otro lado, el uso de planillas de cálculo (Excel) para llevar registros de calidad presenta debilidades intrínsecas: falta de control de cambios, riesgo de alteración de datos (voluntaria o involuntaria) y ausencia de integridad referencial.

**QA-ALCOA-Engine** nace para cerrar esa brecha. Se eligió Office Scripts (TypeScript) operando sobre Excel Online porque permite:
* **Ejecución del lado del servidor:** Los scripts corren en la nube de Microsoft, asegurando que todos los usuarios ejecuten la misma versión del código sin depender de macros VBA locales (las cuales suelen ser bloqueadas por políticas de seguridad IT corporativas).
* **Adopción inmediata:** Mantiene la interfaz de usuario familiar de Excel, reduciendo a cero la curva de aprendizaje para los analistas de Calidad.
* **Trazabilidad Inmutable:** Implementa un diseño *Append-Only* para historiales (*Audit Trails*) con firmas electrónicas por transacción.

---

## 🛠️ Arquitectura y Tecnologías

El framework funciona mediante un modelo de separación estricta entre la Interfaz de Usuario y la Base de Datos:
* **Hojas de Interfaz (`INP_`):** Formularios dinámicos donde el usuario ingresa datos. Poseen campos obligatorios, validación de formatos (ej. Fechas UTC) y celdas bloqueadas para identificadores primarios (`ID`).
* **Hojas de Base de Datos (`BD_` y `HIST_`):** Tablas estructuradas ocultas (`veryHidden`) y protegidas.
- **Controladores TypeScript:** Un conjunto de scripts transaccionales (`Registrar`, `Actualizar`, `Cambio de Estado`, `Anular`) y utilitarios de interfaz (`Buscar`, `Limpiar Formulario`, `Ocultar/Mostrar Hojas`) que aplican lógica de negocio leyendo una tabla de configuración (`TablaReglas`) en tiempo de ejecución.
* **Integridad Criptográfica:** Sellado en vivo mediante **SHA-256** (con *salt*) de las tablas maestras, para detectar y bloquear alteraciones manuales no autorizadas a nivel archivo.
* **RPA (Power Automate):** Cola de salida (*Outbox*) para notificaciones y reportes asincrónicos.

Para comprender a fondo la estructura de tablas, cómo crear reglas de negocio nuevas, o cómo inicializar el sellado criptográfico, consultá el [Manual de Configuración y Arquitectura](MANUAL_CONFIGURACION.md).

---

## 🚀 Despliegue y Desarrollo Local (`puente.py`)

Office Scripts no posee integración nativa directa con Git. Para desarrollar localmente usando un IDE (como Visual Studio Code) y mantener el control de versiones sin perder la sincronización con Excel Online, este repositorio incluye un script de automatización en Python (`puente.py`).

El script gestiona de manera transparente la **segregación entre el Entorno de Desarrollo y el Entorno de Producción**, permitiendo iterar en planillas de pruebas sin riesgo de alterar el código en uso por la planta.

### Configuración del entorno
Para proteger la privacidad de las rutas corporativas de cada desarrollador, la configuración física local está excluida del control de versiones mediante `.gitignore`. Para configurar el entorno:

1. Cloná este repositorio.
2. Buscá el archivo `config.example.py` en la raíz del proyecto.
3. Duplicalo y renombralo como `config.py` (este archivo ya está ignorado por Git).
4. Editá el nuevo archivo definiendo las rutas físicas hacia tu OneDrive personal y la biblioteca sincronizada de SharePoint:
   ```python
   from pathlib import Path

   # 1. RUTA DE DESARROLLO (OneDrive personal para pruebas con sufijo _DEV)
   RUTA_ONEDRIVE_DEV = Path(r"C:/Users/TuUsuario/OneDrive/Documentos/Office Scripts")

   # 2. RUTA DE PRODUCCIÓN (SharePoint oficial de Calidad sincronizado)
   RUTA_SHAREPOINT_PROD = Path(r"C:/Users/TuUsuario/Organizacion/Sitio - Documentos/Office Scripts")
   ```

### Sincronización y Modos de Ejecución

Los archivos de código en el repositorio se desarrollan con el sufijo `_DEV.ts` (ej. `Registrar Universal_DEV.ts`). `puente.py` maneja la inyección según el comando ejecutado:

* **Desarrollo (por defecto):**
  Actualiza únicamente los contenedores `_DEV.osts` en tu OneDrive local vinculados a tus planillas de ensayo.
  ```bash
  python puente.py
  ```
* **Pase a Producción (`--prod` o `-p`):**
  Toma los archivos `_DEV.ts`, les retira automáticamente el sufijo `_DEV` y refresca los `.osts` productivos oficiales en SharePoint (ej. `Registrar Universal.osts`).
  ```bash
  python puente.py --prod
  ```
* **Sincronización Total (`--ambos` o `-a`):**
  Actualiza ambos entornos en una sola operación.
  ```bash
  python puente.py --ambos
  ```

> ⚠️ **Nota:** OneDrive / SharePoint puede demorar entre 10 y 15 segundos en sincronizar los archivos físicos hacia los servidores de Microsoft. Aguardá ese tiempo antes de testear los botones en Excel Online.

## 📄 Licencia
Este proyecto se distribuye bajo la licencia MIT. Consultá el archivo `LICENSE` para más información.