import json
import os
import time
from typing import Dict, Any, List

/**
 * PROTOCOLO DE CONFIGURACIÓN DE SEGURIDAD
 * MOTIVO: Disociar las rutas físicas del código fuente para permitir la portabilidad 
 * y proteger la estructura de carpetas local.
 */
try:
    from config import RUTA_ONEDRIVE_REAL
except ImportError:
    # Estado de contingencia si el archivo de configuración no está presente
    RUTA_ONEDRIVE_REAL = "RUTA_NO_CONFIGURADA"

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

def main() -> None:
    """
    Punto de entrada principal para el proceso de sincronización masiva.
    """
    # 1. SETUP Y VALIDACIÓN DE ENTORNO
    # Determinamos las rutas basándonos en el contexto de ejecución actual.
    directorio_raiz_repositorio = os.getcwd()
    directorio_destino_onedrive = RUTA_ONEDRIVE_REAL
    nombre_identificador_repo = os.path.basename(directorio_raiz_repositorio)

    print(f"🛠️  ORIGEN (Repositorio):  .../{nombre_identificador_repo}")
    print(f"☁️  DESTINO (OneDrive):     [Ruta Protegida en config.py]")
    print("-" * 60)

    # Verificamos que la ruta de destino sea válida y accesible
    if directorio_destino_onedrive != "RUTA_NO_CONFIGURADA" and os.path.exists(directorio_destino_onedrive):
        contador_archivos_actualizados: int = 0
        contador_errores_encontrados: int = 0
        lista_archivos_en_repo: List[str] = os.listdir(directorio_raiz_repositorio)

        # 2. BUCLE DE PROCESAMIENTO SECUENCIAL
        # Recorremos el repositorio buscando archivos TypeScript candidatos a sincronizar.
        for nombre_archivo in lista_archivos_en_repo:
            # Filtramos solo archivos .ts (excluyendo archivos de definición .d.ts)
            if nombre_archivo.endswith(".ts") and not nombre_archivo.endswith(".d.ts"):
                nombre_base_script: str = nombre_archivo[:-3]
                nombre_archivo_osts: str = nombre_base_script + ".osts"

                # Construcción de rutas absolutas para el par de archivos
                ruta_completa_origen: str = os.path.join(directorio_raiz_repositorio, nombre_archivo)
                ruta_completa_destino = os.path.join(directorio_destino_onedrive, nombre_archivo_osts)

                # Verificación de existencia en el destino antes de proceder
                if os.path.exists(ruta_completa_destino):
                    print(f"🔁 Verificando Sincronía: {nombre_base_script} ...")
                    try:
                        fue_actualizado = ejecutar_sincronizacion_par_archivos(ruta_completa_origen, ruta_completa_destino)
                        if fue_actualizado:
                            print(f"    ✅ STATUS: ACTUALIZADO -> El contenedor .osts ha sido refrescado.")
                            contador_archivos_actualizados += 1
                        else:
                            print(f"    💤 STATUS: SINCRONIZADO -> No se requieren cambios.")
                    except Exception:
                        contador_errores_encontrados += 1
                else:
                    # Alerta de inconsistencia: Existe el código pero no el contenedor en OneDrive
                    print(f"    ⚠️  OMITIDO: No se halló '{nombre_archivo_osts}' en el destino.")

        # 3. RESUMEN FINAL DE OPERACIÓN (LOG)
        print("-" * 60)
        print(f"📊 RESUMEN: {contador_archivos_actualizados} actualizados | {contador_errores_encontrados} errores.")

    else:
        # FALLA DE SEGURIDAD O CONFIGURACIÓN
        print("⚠️ ALERTA CRÍTICA: Fallo en la conexión con OneDrive o falta 'config.py'.")
        print("    Acción Requerida: Verifique que RUTA_ONEDRIVE_REAL sea correcta.")

if __name__ == "__main__":
    main()
    # Pausa de cortesía para lectura de logs en consola
    time.sleep(3)