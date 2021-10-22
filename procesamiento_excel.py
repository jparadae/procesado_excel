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
Fax=df[df.Sección=="Fax"]
Fax
print (Fax)

Copiadoras=df[df.Sección=="Copiadoras"]
Copiadoras
df=Copiadoras
print (Copiadoras)

Impresoras=df[df.Sección=="Impresoras"]
Impresoras
print (Impresoras)

SumaSalario=df.Salario.sum()
SumaSalario
print (SumaSalario)
#Generamos un excel con el procesado resultante y se guarda con el nombre generacion_excel
writer = ExcelWriter('generacion_excel.xlsx')
df.to_excel(writer, 'ETL', index=False)
writer.save()
print("Ejemplo de rama y modificacion de archivo")