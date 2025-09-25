import pandas as pd
import matplotlib.pyplot as plt
import os
from pathlib import Path

plt.rcParams['font.family'] = 'DejaVu Sans'

def crear_carpeta_graficas():
    if not os.path.exists('graficas'):
        os.makedirs('graficas')

def leer_datos_excel(archivo_path):
    try:
        hojas_datos = pd.read_excel(archivo_path, sheet_name=None)
        return hojas_datos
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo {archivo_path}")
        return None
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        return None

def generar_grafica_barras(datos, nombre_hoja, indice):
    """Generar gráfica de barras"""
    plt.figure(figsize=(10, 6))
    
    # Asumir que la primera columna son las categorías y la segunda los valores
    columnas = datos.columns.tolist()
    if len(columnas) >= 2:
        categorias = datos.iloc[:, 0]
        valores = datos.iloc[:, 1]
        
        plt.bar(categorias, valores, color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FECA57'])
        plt.title(f'Gráfica de Barras - {nombre_hoja}', fontsize=16, fontweight='bold')
        plt.xlabel(columnas[0], fontsize=12)
        plt.ylabel(columnas[1], fontsize=12)
        plt.xticks(rotation=45, ha='right')
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        
        # Guardar la gráfica
        nombre_archivo = f'graficas/barras_{indice}_{nombre_hoja.lower()}.png'
        plt.savefig(nombre_archivo, dpi=300, bbox_inches='tight')
        plt.close()
        print(f"✓ Gráfica de barras guardada: {nombre_archivo}")
    else:
        print(f"Error: La hoja {nombre_hoja} no tiene suficientes columnas")

def generar_grafica_pastel(datos, nombre_hoja, indice):
    """Generar gráfica de pastel"""
    plt.figure(figsize=(8, 8))
    
    columnas = datos.columns.tolist()
    if len(columnas) >= 2:
        etiquetas = datos.iloc[:, 0]
        valores = datos.iloc[:, 1]
        
        colores = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FECA57', '#FD79A8']
        
        plt.pie(valores, labels=etiquetas, autopct='%1.1f%%', colors=colores, startangle=90)
        plt.title(f'Gráfica de Pastel - {nombre_hoja}', fontsize=16, fontweight='bold')
        plt.axis('equal')  # Para que el pastel sea circular
        
        nombre_archivo = f'graficas/pastel_{indice}_{nombre_hoja.lower()}.png'
        plt.savefig(nombre_archivo, dpi=300, bbox_inches='tight')
        plt.close()
        print(f"✓ Gráfica de pastel guardada: {nombre_archivo}")
    else:
        print(f"Error: La hoja {nombre_hoja} no tiene suficientes columnas")

def procesar_datos(hojas_datos):
    crear_carpeta_graficas()
    
    for nombre_hoja, datos in hojas_datos.items():
        print(f"\nProcesando hoja: {nombre_hoja}")
        print(f"Datos encontrados: {len(datos)} filas")
        
        if nombre_hoja.upper().startswith('BARRAS'):
            indice = nombre_hoja[-1]  # Obtener el número al final
            generar_grafica_barras(datos, nombre_hoja, indice)
        elif nombre_hoja.upper().startswith('PASTEL'):
            indice = nombre_hoja[-1]  # Obtener el número al final
            generar_grafica_pastel(datos, nombre_hoja, indice)
        else:
            print(f"Advertencia: No se reconoce el tipo de gráfica para {nombre_hoja}")

def main():
    print("SCRIPT DE GRAFICACIÓN - PARCIAL 3 OTM 106")
    
    # Ruta del archivo Excel
    archivo_excel = Path('datos/ejemplo.xlsx')
    
    if not archivo_excel.exists():
        print(f"Error: No se encontró el archivo {archivo_excel}")
        print("Asegúrese de que existe la carpeta 'datos' con el archivo 'ejemplo.xlsx'")
        return
    
    # Leer datos del Excel
    print(f"Leyendo datos de: {archivo_excel}")
    hojas_datos = leer_datos_excel(archivo_excel)
    
    if hojas_datos is None:
        return
    
    print(f"Hojas encontradas: {list(hojas_datos.keys())}")
    
    procesar_datos(hojas_datos)
    
    print("PROCESO COMPLETADO EXITOSAMENTE")
    print("Las gráficas se guardaron en la carpeta 'graficas/'")

if __name__ == "__main__":
    main()