import pandas as pd
import os

def imprimir_guia():
    print("=== Guía para principiantes: Validador de Filas Excel ===")
    print("Este programa te ayuda a eliminar filas duplicadas en tus archivos Excel.")
    print("\nCómo usar el programa:")
    print("1. Asegúrate de tener instaladas las bibliotecas pandas y openpyxl.")
    print("   Si no las tienes, instálalas con: pip install pandas openpyxl")
    print("2. Coloca todos los archivos Excel que quieras procesar en una carpeta.")
    print("3. Ejecuta este programa.")
    print("4. Cuando se te solicite, ingresa la ruta completa de la carpeta con tus archivos Excel.")
    print("5. El programa procesará todos los archivos Excel (.xlsx) en esa carpeta.")
    print("6. Por cada archivo, se creará una nueva versión sin filas duplicadas.")
    print("   El nuevo archivo tendrá el prefijo 'sin_duplicados_' en su nombre.")
    print("\nRecuerda: Este programa elimina filas completamente idénticas, manteniendo solo la primera ocurrencia.")
    print("¡Listo! Ahora puedes comenzar a usar el Validador de Filas Excel.\n")

def eliminar_duplicados(archivo_entrada, archivo_salida):
    # Leer el archivo Excel
    df = pd.read_excel(archivo_entrada)
    
    # Obtener el número de filas antes de eliminar duplicados
    filas_antes = len(df)
    
    # Eliminar filas duplicadas, manteniendo la primera ocurrencia
    df_sin_duplicados = df.drop_duplicates(keep='first')
    
    # Obtener el número de filas después de eliminar duplicados
    filas_despues = len(df_sin_duplicados)
    
    # Guardar el DataFrame sin duplicados en un nuevo archivo Excel
    df_sin_duplicados.to_excel(archivo_salida, index=False)
    
    # Calcular el número de filas eliminadas
    filas_eliminadas = filas_antes - filas_despues
    
    print(f"Proceso completado.")
    print(f"Filas en el archivo original: {filas_antes}")
    print(f"Filas en el archivo nuevo: {filas_despues}")
    print(f"Filas duplicadas eliminadas: {filas_eliminadas}")

# Programa principal
if __name__ == "__main__":
    print("Bienvenido al Validador de Filas Excel")
    imprimir_guia()
    
    directorio = input("Ingrese la ruta del directorio con los archivos Excel: ")
    
    archivos_procesados = 0
    for archivo in os.listdir(directorio):
        if archivo.endswith(".xlsx"):
            archivo_entrada = os.path.join(directorio, archivo)
            archivo_salida = os.path.join(directorio, f"sin_duplicados_{archivo}")
            
            print(f"\nProcesando archivo: {archivo}")
            eliminar_duplicados(archivo_entrada, archivo_salida)
            archivos_procesados += 1
    
    if archivos_procesados == 0:
        print("\nNo se encontraron archivos Excel (.xlsx) en el directorio especificado.")
    else:
        print(f"\nSe procesaron {archivos_procesados} archivos Excel.")
    
    print("\n¡Gracias por usar el Validador de Filas Excel!")