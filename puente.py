import json
import os
import time
from typing import Dict, Any, List

# Importación de la configuración local.
# MOTIVO DE SEGURIDAD: Esto permite que 'config.py' tenga los datos reales mientras que
# el script (público) sólo hace referencia a la variable, sin mostrar el valor.
try:
    from config import RUTA_ONEDRIVE_REAL
except ImportError:
    RUTA_ONEDRIVE_REAL = "RUTA_NO_CONFIGURADA"

def sincronizar_un_archivo(ruta_ts: str, ruta_osts: str) -> bool:
    """
    Compara y sincroniza un solo par de archivos (.ts -> .osts).
    Retorna True si hubo cambios, False si no.
    """
    hubo_cambios: bool = False

    try:
        # A. LEER CÓDIGO FUENTE (.ts)
        with open(ruta_ts, 'r', encoding='utf-8') as f_ts:
            codigo_nuevo = f_ts.read()
        
        # B. LEER CONTENEDOR DESTINO (.osts)
        data_json: Dict[str, Any] = {}
        with open(ruta_osts, 'r', encoding='utf-8') as f_osts:
            data_json = json.load(f_osts)

        # C. COMPARAR E INYECTAR
        codigo_viejo = str(data_json.get('body', ''))
        hubo_cambios = (codigo_nuevo != codigo_viejo)

        # D. INYECTAR Y GUARDAR (Solo si es necesario).
        if hubo_cambios:
            data_json['body'] = codigo_nuevo
            with open(ruta_osts, 'w', encoding='utf-8') as f_osts:
                json.dump(data_json, f_osts, ensure_ascii=False, indent=4)


    except Exception as e:
        print(f"   ❌ ERROR TÉCNICO: {str(e)}")
        raise e
    
    # Se devuelve una sola vez el estado calculado.
    return hubo_cambios

def main() -> None:
    # 1. SETUP DE RUTAS
    # MOTIVO DE SEGURIDAD: Uso de os.getcwd() para que el script se adapte a quien lo use
    # sin 'hardcodear' la estructura de carpetas en el código fuente.
    ruta_repo = os.getcwd()
    ruta_onedrive = RUTA_ONEDRIVE_REAL
    nombre_carpeta_repo = os.path.basename(ruta_repo)

    print(f"🛠️  Origen (Repo):     .../{nombre_carpeta_repo}")
    print(f"☁️  Destino (OneDrive): [Ruta oculta configurada en config.py]")
    print("-" * 60)

    if ruta_onedrive != "RUTA_NO_CONFIGURADA" and os.path.exists(ruta_onedrive):
        archivos_modificados: int = 0
        errores: int = 0
        archivos_repo: List[str] = os.listdir(ruta_repo)

        # 2. BUCLE PRINCIPAL

        for archivo in archivos_repo:
            if archivo.endswith(".ts") and not archivo.endswith(".d.ts"):
                nombre_base: str = archivo[:-3]
                nombre_osts: str = nombre_base + ".osts"

                # Rutas absolutas
                ruta_origen: str = os.path.join(ruta_repo, archivo)
                ruta_destino = os.path.join(ruta_onedrive, nombre_osts)

                # verificación de existencia
                if os.path.exists(ruta_destino):
                    print(f"🔁 Verificando: {nombre_base} ...")
                    try:
                        cambio_realizado = sincronizar_un_archivo(ruta_origen, ruta_destino)
                        if cambio_realizado:
                            print(f"   ✅ ACTUALIZADO -> OneDrive ha recibido el nuevo código.")
                            archivos_modificados += 1
                        else:
                            print(f"   zzz Sin cambios.")
                    except Exception:
                        errores +=1
                else:
                    # Si existe el .ts pero no el .osts en OneDrive, avisamos.
                    print(f"   ⚠️ OMITIDO: No existe '{nombre_osts}' en la carpeta de OneDrive.")

        # 3. RESUMEN
        print("-" * 60)
        print(f"📊 Resumen: {archivos_modificados} archivos actualizados | {errores} errores.")



    else:
        # SI FALLA LA CONFIGURACIÓN INICIAL
        print("⚠️ ALERTA: No se encuentra la carpeta de OneDrive o falta el archivo config.py")
        print("    Crea un archivo 'config.py' con la variable RUTA_ONEDRIVE_REAL")

if __name__ == "__main__":
    main()
    time.sleep(3)