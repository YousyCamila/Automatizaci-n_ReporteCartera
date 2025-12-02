import pandas as pd
import pyodbc
import numpy as np
import os
import re

print("=== Iniciando automatización ===")

# ------------------------------
# CONFIGURACIONES
# ------------------------------

EXCEL_PATH = r"C:\Users\ASUS\Documents\EXCEL CARTERA\Pendientes x aplicar SOAT 30092025.xlsx"
SERVER = "NANOYOKI-06\\SQLEXPRESS"
DATABASE = "ReporteCartera"
USERNAME = "sa"
PASSWORD = "Sh@damy159"
TABLE_NAME = "ReporteCartera"

# ------------------------------
# Verificar archivo Excel
# ------------------------------

if not os.path.exists(EXCEL_PATH):
    print("ERROR: No se encontró el archivo Excel en la ruta:")
    print(EXCEL_PATH)
    exit()

print(" Archivo Excel encontrado.")

# ------------------------------
# Cargar Excel
# ------------------------------

print("=== Cargando Excel correctamente ===")

df = pd.read_excel(
    EXCEL_PATH,
    sheet_name="Balance Antigüedad",
    header=7
)

df.columns = df.columns.map(str)
df = df.loc[:, ~df.columns.str.contains("Unnamed")]
df = df.dropna(how="all")

print("Columnas detectadas en el Excel:")
print(df.columns.tolist())

# ------------------------------
# BUSCAR COLUMNA SECTOR
# ------------------------------

sector_cols = [c for c in df.columns if "sector" in c.lower()]

if sector_cols:
    sector_col = sector_cols[0]
    print(f" Columna SECTOR detectada: {sector_col}")
else:
    print("No se encontró columna SECTOR, se creará.")
    df["Sector"] = "PRIVADO"
    sector_col = "Sector"

# Sanear SECTOR
df["Sector"] = (
    df[sector_col]
    .astype(str)
    .str.upper()
    .str.replace(r"[^A-ZÁÉÍÓÚÑ ]", "", regex=True)
    .str.strip()
)

# Reemplazar valores inválidos
df["Sector"] = df["Sector"].replace(["NAN", "NONE", "NAT", ""], "PRIVADO")

# ------------------------------
# LIMPIEZA GENERAL: SOLO TEXTO VÁLIDO
# ------------------------------

cols_texto = ["Sector", "Sucursal", "Asegurado", "Intermediario", "Ramo"]

for col in cols_texto:
    if col in df.columns:
        df[col] = df[col].astype(str)
        df[col] = df[col].apply(lambda x: re.sub(r"[^A-Za-z0-9 ÁÉÍÓÚÑáéíóú.-]", "", x)).str.strip()

# ------------------------------
# NORMALIZAR COLUMNAS
# ------------------------------

df.columns = df.columns.str.strip().str.replace(" ", "_")

# ------------------------------
# FECHAS
# ------------------------------

date_map = {
    "FECHA_EMISION": "FechaEmision",
    "FECHA_DE_VIGENCIA_DESDE": "FechaVigenciaDesde",
    "FECHA_DE_VIGENCIA_HASTA": "FechaVigenciaHasta"
}

for col_old, col_new in date_map.items():
    if col_old in df.columns:
        df[col_new] = pd.to_datetime(df[col_old], errors='coerce').dt.date
    else:
        df[col_new] = None

# ------------------------------
# NUMÉRICOS
# ------------------------------

if "PRIMA_TOTAL" in df.columns:
    df["PrimaTotal"] = pd.to_numeric(df["PRIMA_TOTAL"], errors='coerce').fillna(0)
else:
    df["PrimaTotal"] = 0

# ------------------------------
# RENOMBRAR COLUMNAS A SQL
# ------------------------------

df = df.rename(columns={
    "COD_INTERMEDIARIO": "CodIntermediario",
    "INTERMEDIARIO": "Intermediario",
    "MONEDA": "Moneda",
    "SECTOR": "Sector",
    "SUCURSAL": "Sucursal",
    "ASEGURADO": "Asegurado",
    "RAMO": "Ramo",
    "NRO._FACTURA": "NumeroFactura",
    "POLIZA": "Poliza",
    "ENDOSO": "Endoso",
    "ANTIGUEDAD": "Antiguedad"
})

print("Columnas finales:")
print(df.columns.tolist())

# ------------------------------
# CONEXIÓN A SQL
# ------------------------------

try:
    print("Conectando a SQL Server...")
    conn = pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        f"UID={USERNAME};"
        f"PWD={PASSWORD}"
    )
    cursor = conn.cursor()
    print("Conexión exitosa.")
except Exception as e:
    print(" ERROR al conectar a SQL Server:")
    print(e)
    exit()

# ------------------------------
# INSERTAR DATOS
# ------------------------------

insert_query = f"""
INSERT INTO {TABLE_NAME} (
    CodIntermediario, Intermediario, Moneda, Sector, Sucursal, Asegurado, Ramo,
    NumeroFactura, Poliza, Endoso, FechaEmision, Antiguedad,
    FechaVigenciaDesde, FechaVigenciaHasta, PrimaTotal
)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""

count = 0
print("Iniciando inserciones...")

for index, row in df.iterrows():
    try:
        cursor.execute(insert_query,
            str(row.get("CodIntermediario")),
            str(row.get("Intermediario")),
            str(row.get("Moneda")),
            str(row.get("Sector")),
            str(row.get("Sucursal")),
            str(row.get("Asegurado")),
            str(row.get("Ramo")),
            str(row.get("NumeroFactura")),
            str(row.get("Poliza")),
            str(row.get("Endoso")),
            row.get("FechaEmision"),
            row.get("Antiguedad"),
            row.get("FechaVigenciaDesde"),
            row.get("FechaVigenciaHasta"),
            float(row.get("PrimaTotal"))
        )
        count += 1
    except Exception as e:
        print(f"Error al insertar fila {index}: {e}")

conn.commit()
cursor.close()
conn.close()

print(f" Inserción finalizada. Registros insertados: {count}")
print("=== Automatización finalizada ===")
