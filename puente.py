import json
import os
import sys
import time
from typing import Dict, Any, List

"""
 * PROTOCOLO DE CONFIGURACIÓN DE SEGURIDAD
 * MOTIVO: Disociar las rutas físicas del código fuente para permitir la portabilidad 
 * y proteger la estructura de carpetas local tanto en DEV como en PROD.
"""

try:
    from config import RUTA_ONEDRIVE_DEV, RUTA_SHAREPOINT_PROD
    
except ImportError:
    # Estado de contingencia si el archivo de configuración no está presente
    RUTA_ONEDRIVE_DEV = "RUTA_NO_CONFIGURADA"
    RUTA_SHAREPOINT_PROD = "RUTA_NO_CONFIGURADA"
    

def ejecutar_sincronizacion_par_archivos(ruta_fuente_ts: str, ruta_contenedor_osts: str) -> bool:
    """
    Compara y sincroniza un par de archivos (.ts -> .osts).
    Garantiza que el 'body' del JSON en OneDrive coincida con el código en el Repositorio.
    Retorna True si se realizó una actualización física del archivo.
    """
    estado_cambio_detectado: bool = False

    try:
        # A. EXTRACCIÓN DE CÓDIGO FUENTE (.ts)
        # Se lee el archivo TypeScript que contiene la lógica refactorizada.
        with open(ruta_fuente_ts, 'r', encoding='utf-8') as archivo_fuente:
            contenido_codigo_nuevo = archivo_fuente.read()
        
        # B. LECTURA DE CONTENEDOR DESTINO (.osts)
        # Los Office Scripts son archivos JSON. Extraemos el objeto completo.
        estructura_json_osts: Dict[str, Any] = {}
        with open(ruta_contenedor_osts, 'r', encoding='utf-8') as archivo_destino:
            estructura_json_osts = json.load(archivo_destino)

        # C. VERIFICACIÓN DE INTEGRIDAD Y COMPARACIÓN
        # Comparamos el código actual en OneDrive contra el nuevo código del Repo.
        contenido_codigo_existente = str(estructura_json_osts.get('body', ''))
        estado_cambio_detectado = (contenido_codigo_nuevo != contenido_codigo_existente)

        # D. INYECCIÓN DE CÓDIGO Y PERSISTENCIA
        # Solo se sobreescribe el archivo si se detecta una diferencia (Optimización de Sync).
        if estado_cambio_detectado:
            estructura_json_osts['body'] = contenido_codigo_nuevo
            with open(ruta_contenedor_osts, 'w', encoding='utf-8') as archivo_destino:
                json.dump(estructura_json_osts, archivo_destino, ensure_ascii=False, indent=4)

    except Exception as excepcion_tecnica:
        print(f"    ❌ ERROR DE PROTOCOLO: {str(excepcion_tecnica)}")
        raise excepcion_tecnica
    
    return estado_cambio_detectado

def sincronizar_directorio(directorio_raiz: str, directorio_destino: str, etiqueta_entorno: str, es_produccion: bool) -> None:
    """
    Ejecuta el ciclo de sincronización contra un directorio destino (DEV o PROD).
    Si es_produccion es True, elimina el sufijo '_DEV' de los nombres de destino.
    """
    print(f"\n🚀 INICIANDO SINCRONIZACIÓN -> [{etiqueta_entorno}]")
    print(f"📁 Destino: {directorio_destino}")
    print("-" * 60)

    if directorio_destino == "RUTA_NO_CONFIGURADA" or not os.path.exists(directorio_destino):
        print(f"⚠️ ALERTA CRÍTICA: No se puede acceder a la ruta de [{etiqueta_entorno}].")
        print("    Verifique la ruta configurada en 'config.py'.")
        return

    contador_actualizados: int = 0
    contador_errores: int = 0
    lista_archivos: List[str] = os.listdir(directorio_raiz)

    for nombre_archivo in lista_archivos:
        if nombre_archivo.endswith(".ts") and not nombre_archivo.endswith(".d.ts"):
            nombre_base_script: str = nombre_archivo[:-3]  # Quita ".ts"

            # En Producción quitamos el sufijo '_DEV' del archivo destino .osts
            if es_produccion and nombre_base_script.endswith("_DEV"):
                nombre_base_destino = nombre_base_script[:-4]  # Quita '_DEV'
            else:
                nombre_base_destino = nombre_base_script

            nombre_archivo_osts: str = nombre_base_destino + ".osts"

            ruta_completa_origen: str = os.path.join(directorio_raiz, nombre_archivo)
            ruta_completa_destino = os.path.join(str(directorio_destino), nombre_archivo_osts)

            if os.path.exists(ruta_completa_destino):
                print(f"🔁 Verificando: {nombre_archivo} -> {nombre_archivo_osts} ...")
                try:
                    fue_actualizado = ejecutar_sincronizacion_par_archivos(ruta_completa_origen, ruta_completa_destino)
                    if fue_actualizado:
                        print("    ✅ STATUS: ACTUALIZADO -> El contenedor .osts ha sido refrescado.")
                        contador_actualizados += 1
                    else:
                        print("    💤 STATUS: SINCRONIZADO -> Sin cambios requeridos.")
                except Exception:
                    contador_errores += 1
            else:
                print(f"    ⚠️  OMITIDO: No se halló '{nombre_archivo_osts}' en el destino.")

    print(f"📊 RESUMEN [{etiqueta_entorno}]: {contador_actualizados} actualizados | {contador_errores} errores.")


def main() -> None:
    """
    Punto de entrada principal. Determina el entorno a sincronizar según los argumentos:
      python puente.py          -> Sincroniza DESARROLLO (por defecto)
      python puente.py --prod   -> Sincroniza PRODUCCIÓN (SharePoint)
      python puente.py --ambos  -> Sincroniza ambos entornos
    """
    # Determinamos la raíz del repositorio de forma absoluta
    directorio_raiz_repositorio = os.path.dirname(os.path.abspath(__file__))
    nombre_repo = os.path.basename(directorio_raiz_repositorio)

    print(f"🛠️  ORIGEN (Repositorio): .../{nombre_repo}")
    print("-" * 60)

    # Detección de argumentos de línea de comandos
    argumentos = [arg.lower() for arg in sys.argv[1:]]

    if "--ambos" in argumentos or "-a" in argumentos:
        sincronizar_directorio(directorio_raiz_repositorio, RUTA_ONEDRIVE_DEV, "DESARROLLO (OneDrive)", es_produccion=False)
        sincronizar_directorio(directorio_raiz_repositorio, RUTA_SHAREPOINT_PROD, "PRODUCCIÓN (SharePoint)", es_produccion=True)
    elif "--prod" in argumentos or "-p" in argumentos:
        sincronizar_directorio(directorio_raiz_repositorio, RUTA_SHAREPOINT_PROD, "PRODUCCIÓN (SharePoint)", es_produccion=True)
    else:
        # Modo por defecto: DESARROLLO
        sincronizar_directorio(directorio_raiz_repositorio, RUTA_ONEDRIVE_DEV, "DESARROLLO (OneDrive)", es_produccion=False)


if __name__ == "__main__":
    main()
    time.sleep(2)