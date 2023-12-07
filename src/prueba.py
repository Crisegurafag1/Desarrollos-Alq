import pandas as pd
import numpy as np
# Crear un DataFrame de ejemplo
data = {
    'KILO_NETO': [-0.0000000000000001110223, -0.412000000000001, -7.9999999999999500000e-01, -0.0000000000000000277556],
    'VALORTOTAL_PRODREGULAR': [0, -3370, -43764.70588, -4140],
}

df_agrupadoVentas = pd.DataFrame(data)



# Constantes
UMBRAL_LIMPIEZA = 1e-6
# Limpiar y convertir los datos antes de realizar la división
df_agrupadoVentas['KILO_NETO'] = df_agrupadoVentas['KILO_NETO'].apply(lambda x: 0 if abs(x) < UMBRAL_LIMPIEZA else x)  # Ajusta el umbral según tus necesidades
#df_agrupadoVentas['PRODREGULAR X KILO'] = df_agrupadoVentas['VALORTOTAL_PRODREGULAR'] / df_agrupadoVentas['KILO_NETO']
df_agrupadoVentas['PRODREGULAR X KILO'] = np.where(df_agrupadoVentas['VALORTOTAL_PRODREGULAR'] != 0
                                                   , df_agrupadoVentas['VALORTOTAL_PRODREGULAR'] 
                                                   / df_agrupadoVentas['KILO_NETO'], 0)
df_agrupadoVentas['PRODREGULAR X KILO'].replace(-np.inf, 0, inplace=True)
# Exportar a un archivo Excel
df_agrupadoVentas.to_excel(
    'C:/Users/cristian.segura/Documents/Python/Proyecto2/Proyección de Ventas/4) Validación/Prueba.xlsx'
    , index=False
    , sheet_name='Hoja1'
    , engine='openpyxl')
print(df_agrupadoVentas)
      #df.to_excel('resultado.xlsx', index=False)
