import csv
from openpyxl import Workbook

nombres_y_apellidos = []
with open("alumnos.csv", newline="", encoding="utf-8") as file:
    reader = csv.reader(file)
    for row in reader:
        nombres_y_apellidos.append(row)

#print(array)

workbook = Workbook()
hoja = workbook.active

contador = 0 # Contador para saber cuantos personas
columna = 0
fila = 1 # Fila en la que comienza (El xlsx comienza en 1)
offset_fila = 0 
contador_parejas_por_persona = 0

for indice, nombre in enumerate(nombres_y_apellidos):
    fila = (indice // 5) + 1  # Calcula la fila actual (1-based)
    columna = (indice % 5) + 1  # Calcula la columna actual (1-based: 1-5)
    hoja.cell(row=fila, column=columna, value=nombre)


workbook.save("resultado.xlsx")