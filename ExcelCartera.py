import pandas as pd
import numpy as np
import os
import re

print("=== INICIANDO LIMPIEZA BALANCE ANTIGÜEDAD ===")

INPUT_PATH = r"C:\Users\ASUS\Documents\EXCEL CARTERA\Pendientes x aplicar SOAT 30092025.xlsx"
SHEET_NAME = "Balance Antigüedad"

OUTPUT_PATH = r"C:\Users\ASUS\Documents\EXCEL CARTERA\Balance_Antiguedad_Limpio.xlsx"

# Validar que el archivo exista
print("Existe archivo:", os.path.exists(INPUT_PATH))

# Leer SOLO la hoja Balance Antigüedad
df = pd.read_excel(
    INPUT_PATH,
    sheet_name=SHEET_NAME,
    header=7
)


print("Hoja cargada correctamente:", SHEET_NAME)
print("Columnas detectadas:")
print(df.columns.tolist())
# ------------------------------
# NORMALIZACIÓN DE TEXTO
# ------------------------------
def normalizar(texto):
    if pd.isna(texto):
        return ""
    return str(texto).lower().strip()

df["SECTOR"] = df["SECTOR"].astype(str)
df["SUCURSAL"] = df["SUCURSAL"].astype(str)
df["ASEGURADO"] = df["ASEGURADO"].astype(str)

# ------------------------------
# FECHA EMISIÓN → SOLO DATE
# ------------------------------
if "FECHA EMISION" in df.columns:
    df["FECHA EMISION"] = pd.to_datetime(
        df["FECHA EMISION"], errors="coerce"
    ).dt.date

# ------------------------------
# REGLAS DE SECTOR
# ------------------------------

# Palabras clave que obligan a OFICIAL si está en PRIVADO
palabras_oficial = [
   "municipio",
    "hospital",
    "aguas",
    "energia",
    "policia",
    "asorrecio"
]

def ajustar_sector(row):
    sector = normalizar(row["SECTOR"])
    asegurado = normalizar(row["ASEGURADO"])
    sucursal = normalizar(row["SUCURSAL"])

    # 1. Sucursal estatal → OFICIAL
    if "estatal" in sucursal:
        return "OFICIAL"

    # 2. Sucursal virtual → PRIVADO
    if "virtual" in sucursal:
        return "PRIVADO"

   # 3. Asegurado con palabras clave → OFICIAL
    if sector == "privado" and any(p in asegurado for p in palabras_oficial):
        return "OFICIAL"

    # 4. Persona natural → PRIVADO
    # (no contiene palabras empresariales comunes)
    if not any(p in asegurado for p in [
        "ltda", "s.a", "sas", "municipio", "hospital",
        "empresa", "corporacion", "universidad"
    ]):
        return "PRIVADO"

    # Si no aplica ninguna regla, se queda igual
    return sector.upper()

df["SECTOR"] = df.apply(ajustar_sector, axis=1)

# ------------------------------
# LIMPIEZA FINAL
# ------------------------------
df["SECTOR"] = df["SECTOR"].str.upper().str.strip()

# ------------------------------
# GUARDAR NUEVO EXCEL
# ------------------------------
df.to_excel(OUTPUT_PATH, index=False)

print("=== PROCESO FINALIZADO ===")
print(f"Archivo generado en:\n{OUTPUT_PATH}")
