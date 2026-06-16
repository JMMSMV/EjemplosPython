import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from PyPDF2 import PdfMerger
from PIL import Image
import os
import sqlite3 

# --- 1. Variables de rutas ---
carpeta_base1 = "/storage/emulated/0/80 Documentos Legales/CodigosScriptPython/ProcesoImagenes"
carpeta_base2 = "/storage/emulated/0/80 Documentos Legales/DatosPrimarios"

ruta_base1 = os.path.join(carpeta_base1, "DetalleImagenes10.xlsx")
ruta_base2 = os.path.join(carpeta_base2, "DetalleImagenes.xlsx")

# --- 2. Cargar ambas bases ---
df_base1 = pd.read_excel(ruta_base1)
df_base2 = pd.read_excel(ruta_base2)

# --- 3. Normalizar claves ---
df_base1["Consecutivo"] = df_base1["Consecutivo"].astype(str)
df_base2["Consecutivo"] = df_base2["Consecutivo"].astype(str)

# --- 4. Definir nombres estándar (incluye CM1...CM19) ---
nombres_estandar = ["Consecutivo", "Archivo", "Camino", "Bloque"] + [f"CM{i}" for i in range(1,20)]

# --- 5. Sincronizar bases ---
df_sincronizada = pd.merge(
    df_base1,
    df_base2,
    on="Consecutivo",
    how="outer",
    suffixes=("_x", "_y")
)

# --- 6. Combinar columnas duplicadas automáticamente ---
for col in nombres_estandar:
    col_x, col_y = f"{col}_x", f"{col}_y"
    if col_x in df_sincronizada.columns and col_y in df_sincronizada.columns:
        df_sincronizada[col] = df_sincronizada[col_x].combine_first(df_sincronizada[col_y])
        df_sincronizada.drop([col_x, col_y], axis=1, inplace=True)
    elif col_x in df_sincronizada.columns:
        df_sincronizada.rename(columns={col_x: col}, inplace=True)
    elif col_y in df_sincronizada.columns:
        df_sincronizada.rename(columns={col_y: col}, inplace=True)

# --- 7. Filtrar columnas finales ---
columnas_existentes = [col for col in nombres_estandar if col in df_sincronizada.columns]
df_sincronizada = df_sincronizada[columnas_existentes]

# --- 8. Guardar resultado ---
ruta_destino = os.path.join(carpeta_base1, "DetalleImagenesSincronizada.xlsx")
df_sincronizada.to_excel(ruta_destino, index=False)
print("✅ Base sincronizada generada con columnas estándar:", columnas_existentes)

print (df_sincronizada)

ruta_maestro = "/storage/emulated/0/80 Documentos Legales/DatosPrimarios/BaseMaestra.xlsx"

# df_maestro = pd.read_excel(ruta_maestro, sheet_name="Table 1")

# --- 1. Conectar a SQLite ---
conn = sqlite3.connect("catalogo.db")

# --- 2. Cargar bases sincronizadas y maestra ---
df_sinc = pd.read_excel(ruta_destino)
df_maestra = pd.read_excel(ruta_maestro,sheet_name="Table 1")

# Guardar en SQLite
df_sinc.to_sql("DetalleImagenesSincronizada", conn, if_exists="replace", index=False)
df_maestra.to_sql("BaseMaestra", conn, if_exists="replace", index=False)

# --- 3. Dinamizar con SQL JOIN ---
query = """
SELECT m.N_Orden, m.NombreArticulo, m.PrecioUSD,
       d.Archivo, d.Camino, d.Bloque
FROM BaseMaestra m
JOIN DetalleImagenesSincronizada d
ON m.N_Orden = d.Consecutivo
"""
df_dinamizada = pd.read_sql_query(query, conn)

df_dinamizada.to_excel("DetalleImagenesDinamizada.xlsx", index=False)
print("✅ Base dinamizada generada: DetalleImagenesDinamizada.xlsx")
