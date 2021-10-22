# -*- coding: utf-8 -*-
"""
Ejemplo de procesamiento de excel con python 
Academia Yunus
Javiera Parada Epul
"""

# Importamos las librerias necesarias para procesar el excel
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pandas import ExcelWriter

#Generamos la opción de abrir el excel base_de_datos.xlsx
print("Abro el excel")
df=pd.read_excel("base_de _datos.xlsx")
#Obtenemos la cabecera
df.head()
df.sample(6)
df.tail()
#Generamos un array ordenando las cabeceras del excel
print("Ordeno las columnas")
df = df[["Nombre", "Apellido", "Departamento", "Sección", "Matrícula", "Salario", "Fecha ingreso"]]
df.head()
df.tail()
df.sample(6)
#Filtramos la seccion de copiadoras
print("Filtramos")
Sección=df[df.Sección == "Copiadoras"]
Sección
#Generamos un excel con el procesado resultante y se guarda con el nombre generacion_excel
writer = ExcelWriter('generacion_excel.xlsx')
df.to_excel(writer, 'Hoja de datos', index=False)
writer.save()
