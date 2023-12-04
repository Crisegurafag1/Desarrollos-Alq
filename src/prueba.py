import os
import time

# Ruta del directorio
RUTA = 'C:/Users/cristian.segura/Documents/Python/Proyecto2/Proyección de Ventas/VolumenDeProyección'
# Nombre de los archivos que esperamos encontrar
BDKILOSVAL = 'bdKilosVal.xlsx'
BDCATEGORIA = 'bdCategoria.xlsx'
BDCANAL = 'bdCanal.xlsx'

while True:
    # Lista los archivos en la ruta
    FILES = os.listdir(RUTA)

    # Verifica qué archivo está presente
    if BDKILOSVAL in FILES:
        print('\n---------------------------------\n')
        print(f" El archivo {BDKILOSVAL} está presente. Realizar función para Ventas.")
        break
    elif BDCATEGORIA in FILES:
        # Porcentaje Individual y homologación tabla dinámica
        print(f" El archivo {BDCATEGORIA} está presente. Se realizará la explosión en función para Categoría.")
        break
    else:
        print(" No hay archivos para realizar la explosión.")
        print('\n---------------------------------\n')
        
        # Solicitar al usuario que presione Enter para volver a verificar
        input("Presiona Enter cuando hayas agregado el archivo y quieras volver a verificar.")
