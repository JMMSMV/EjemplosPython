import pandas as pd
import os

print ("More") 
print ("---------Seleccion por Etiqueta----------------")
# DataFrame de ejemplo
data = {
    "Producto": ["Laptop", "Tablet", "Smartphone", "Monitor"],
    "Precio": [1200, 300, 800, 200],
    "Stock": [15, 40, 25, 10]
}
df = pd.DataFrame(data, index=["A1", "A2", "A3", "A4"])

# Selección por etiqueta
print(df.loc["A2"])  # Fila completa Tablet
print(df.loc[["A1", "A3"], ["Producto", "Precio"]])  # Subconjunto

print ("---------Seleccion por Posicion------------------")

# Selección por posición
print(df.iloc[0])  # Primera fila
print(df.iloc[:, 1])  # Segunda columna (Precio)
print(df.iloc[1:3, 0:2])  # Rango de filas y columnas

print ("--------------Filtrar Productos-------------")
# Filtrar productos con precio mayor a 500
print(df[df["Precio"] > 500])

# Filtrar productos con stock entre 10 y 30
print(df[(df["Stock"] >= 10) & (df["Stock"] <= 30)])


print ("---------MultiIndex------------------")
# Crear un DataFrame con MultiIndex
arrays = [
    ["Electrónica", "Electrónica", "Accesorios", "Accesorios"],
    ["Laptop", "Tablet", "Mouse", "Teclado"]
]
index = pd.MultiIndex.from_arrays(arrays, names=("Categoría", "Producto"))

df_multi = pd.DataFrame({"Precio": [1200, 300, 20, 50], "Stock": [15, 40, 100, 60]}, index=index)

# Selección por nivel
print(df_multi.loc["Electrónica"])  # Todos los productos electrónicos
print ("--------------")
print(df_multi.loc[("Accesorios", "Mouse")])  # Solo Mouse
