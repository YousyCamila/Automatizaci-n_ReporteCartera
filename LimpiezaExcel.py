import pandas as pd
import os
import pyodbc
import numpy as np

print("=== INICIANDO LIMPIEZA BALANCE ANTIGÜEDAD ===")

# ==============================
# RUTAS
# ==============================
INPUT_PATH = r"C:\Users\ASUS\Documents\EXCEL CARTERA\Pendientes x aplicar SOAT 30092025.xlsx"
SHEET_NAME = "Balance Antigüedad"
OUTPUT_PATH = r"C:\Users\ASUS\Documents\EXCEL CARTERA\Balance_Antiguedad_Limpio.xlsx"

# ==============================
# CONEXIÓN SQL SERVER
# ==============================
SERVER = "NANOYOKI-06\\SQLEXPRESS"
DATABASE = "ReporteCartera"
USERNAME = "sa"
PASSWORD = "Sh@damy159"

connection_string = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    f"UID={USERNAME};"
    f"PWD={PASSWORD};"
    "TrustServerCertificate=yes;"
)

# ==============================
# LEER EXCEL
# ==============================
print("Existe archivo:", os.path.exists(INPUT_PATH))

df = pd.read_excel(
    INPUT_PATH,
    sheet_name=SHEET_NAME,
    header=7
)

print("Columnas detectadas:")
print(df.columns.tolist())

# ==============================
# NORMALIZACIÓN
# ==============================
def normalizar(texto):
    if pd.isna(texto):
        return ""
    return str(texto).lower().strip()

for col in ["SECTOR", "SUCURSAL", "ASEGURADO"]:
    df[col] = df[col].astype(str)

# ==============================
# FECHAS → DATE
# ==============================
for col in ["FECHA EMISION", "FECHA DE VIGENCIA DESDE", "FECHA DE VIGENCIA HASTA"]:
    df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True).dt.date

# ==============================
# REGLAS DE SECTOR
# ==============================
def ajustar_sector(row):
    sector_original = normalizar(row["SECTOR"])
    asegurado = normalizar(row["ASEGURADO"])
    sucursal = normalizar(row["SUCURSAL"])

    if "estatal" in sucursal:
        return "OFICIAL"
    if "virtual" in sucursal:
        return "PRIVADO"

    if sector_original == "privado" and any(
        p in asegurado for p in [
            "municipio", "hospital", "aguas",
            "energia", "policia", "asorrecio"
        ]
    ):
        return "OFICIAL"

    palabras_empresa = [
        "ltda", "sas", "s.a", "s a", "e.s.p", "esp",
        "empresa", "corporacion", "universidad",
        "policia", "aguas", "energia"
    ]

    if not any(p in asegurado for p in palabras_empresa):
        return "PRIVADO"

    return sector_original.upper()

df["SECTOR"] = df.apply(ajustar_sector, axis=1)

# ==============================
# PRIMA TOTAL → NUMÉRICO
# ==============================
df["PRIMA TOTAL"] = (
    df["PRIMA TOTAL"]
    .astype(str)
    .str.replace(r"[^\d]", "", regex=True)
)

df["PRIMA TOTAL"] = pd.to_numeric(df["PRIMA TOTAL"], errors="coerce").fillna(0).astype(int)

# ==============================
# GUARDAR EXCEL LIMPIO
# ==============================
df.to_excel(OUTPUT_PATH, index=False)
print("Excel limpio creado:", OUTPUT_PATH)

# ==============================
# RENOMBRAR COLUMNAS → SQL
# ==============================
df_sql = df.rename(columns={
    "COD INTERMEDIARIO": "CodIntermediario",
    "INTERMEDIARIO": "Intermediario",
    "MONEDA": "Moneda",
    "SECTOR": "Sector",
    "SUCURSAL": "Sucursal",
    "ASEGURADO": "Asegurado",
    "RAMO": "Ramo",
    "NRO. FACTURA": "NroFactura",
    "POLIZA": "Poliza",
    "ENDOSO": "Endoso",
    "FECHA EMISION": "FechaEmision",
    "ANTIGUEDAD": "Antiguedad",
    "FECHA DE VIGENCIA DESDE": "FechaVigenciaDesde",
    "FECHA DE VIGENCIA HASTA": "FechaVigenciaHasta",
    "PRIMA TOTAL": "PrimaTotal"
})

# ==============================
# NaN → None (SQL NULL)
# ==============================
df_sql = df_sql.replace({np.nan: None})

# ==============================
# INSERTAR EN SQL SERVER
# ==============================
print("Conectando a SQL Server...")
conn = pyodbc.connect(connection_string)
cursor = conn.cursor()
cursor.fast_executemany = True

insert_query = """
INSERT INTO dbo.PolizasCartera (
    CodIntermediario, Intermediario, Moneda, Sector, Sucursal,
    Asegurado, Ramo, NroFactura, Poliza, Endoso,
    FechaEmision, Antiguedad, FechaVigenciaDesde,
    FechaVigenciaHasta, PrimaTotal
)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""

data = [
    (
        row.CodIntermediario,
        row.Intermediario,
        row.Moneda,
        row.Sector,
        row.Sucursal,
        row.Asegurado,
        row.Ramo,
        row.NroFactura,
        row.Poliza,
        row.Endoso,
        row.FechaEmision,
        row.Antiguedad,
        row.FechaVigenciaDesde,
        row.FechaVigenciaHasta,
        row.PrimaTotal
    )
    for row in df_sql.itertuples(index=False)
]

cursor.executemany(insert_query, data)
conn.commit()

cursor.close()
conn.close()

print("=== CARGA COMPLETADA EN SQL SERVER ===")
