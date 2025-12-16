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

df.columns = df.columns.map(str).str.strip()
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
    df["SECTOR"] = "PRIVADO"
    sector_col = "SECTOR"

# Normalizar SECTOR
df["SECTOR"] = (
    df[sector_col]
    .astype(str)
    .str.upper()
    .str.replace(r"[^A-ZÁÉÍÓÚÑ ]", "", regex=True)
    .str.strip()
)

df["SECTOR"] = df["SECTOR"].replace(["NAN", "NONE", "NAT", ""], "PRIVADO")

# ------------------------------
# LIMPIEZA GENERAL
# ------------------------------

cols_texto = ["SECTOR", "SUCURSAL", "ASEGURADO", "INTERMEDIARIO", "RAMO"]

for col in cols_texto:
    if col in df.columns:
        df[col] = df[col].astype(str)
        df[col] = df[col].apply(lambda x: re.sub(r"[^A-Za-z0-9 ÁÉÍÓÚÑáéíóú.-]", "", x)).str.strip()
    else:
        print(f"ADVERTENCIA: La columna {col} no existe en el Excel.")

# ------------------------------
# FILTROS ESPECIALES
# ------------------------------

print("Aplicando reglas de sector...")

# FECHA EMISIÓN
if "FECHA EMISION" in df.columns:
    df["FECHA EMISION"] = (
        pd.to_datetime(df["FECHA EMISION"], errors="coerce", dayfirst=True)
        .dt.date
    )


# Normalizar cadenas
df["Asegurado_tmp"] = df["ASEGURADO"].astype(str).str.upper()
df["Sucursal_tmp"] = df["SUCURSAL"].astype(str).str.upper()

# MUNICIPIOS → OFICIAL
municipios_keywords = [
    "MUNICIPIO", "ALCALDIA", "ALCALDÍA", "GOBERNACION", "GOBERNACIÓN"
]

df.loc[
    df["Asegurado_tmp"].str.contains("|".join(municipios_keywords), na=False),
    "SECTOR"
] = "OFICIAL"

# HOSPITALES, AGUAS, POLICIA → OFICIAL
oficial_keywords = [
    "HOSPITAL", "AGUAS", "ENERGIA", "ENERGÍA",
    "POLICIA", "POLICÍA", "ASORECIO"
]

df.loc[
    df["Asegurado_tmp"].str.contains("|".join(oficial_keywords), na=False),
    "SECTOR"
] = "OFICIAL"

# SUCURSAL ESTATAL
df.loc[
    df["Sucursal_tmp"].str.contains("ESTATAL", na=False),
    "SECTOR"
] = "OFICIAL"

# SUCURSAL VIRTUAL
df.loc[
    df["Sucursal_tmp"].str.contains("VIRTUAL", na=False),
    "SECTOR"
] = "PRIVADO"

# PERSONA NATURAL
empresa_keywords = [
    "S.A", "SAS", "LTDA", "E.S.E", "ESE", "EMPRESA",
    "HOSPITAL", "FUNDACION", "FUNDACIÓN", "ASOCIACION", "ASOCIACIÓN"
]

df["EsEmpresa"] = df["Asegurado_tmp"].str.contains("|".join(empresa_keywords), na=False)

df.loc[
    ~df["EsEmpresa"] & (df["SECTOR"] != "OFICIAL"),
    "SECTOR"
] = "PRIVADO"

# Eliminar columnas temporales
df = df.drop(columns=["Asegurado_tmp", "Sucursal_tmp", "EsEmpresa"], errors="ignore")

# ------------------------------
# NORMALIZAR COLUMNAS
# ------------------------------

df.columns = df.columns.str.strip().str.replace(" ", "_")

# ------------------------------
# FECHAS
# ------------------------------

# ------------------------------
# LIMPIAR FECHAS
# ------------------------------

def limpiar_fecha(col):
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.date
    else:
        df[col] = None

# Nombres reales ya transformados
limpiar_fecha("FECHA_EMISION")
limpiar_fecha("FECHA_DE_VIGENCIA_DESDE")
limpiar_fecha("FECHA_DE_VIGENCIA_HASTA")

# Crear las columnas finales que SQL espera
df["FechaEmision"] = df["FECHA_EMISION"]
df["FechaVigenciaDesde"] = df["FECHA_DE_VIGENCIA_DESDE"]
df["FechaVigenciaHasta"] = df["FECHA_DE_VIGENCIA_HASTA"]

# ------------------------------
# ANTIGUEDAD
# ------------------------------
# Asegurarnos que sea varchar como en Excel
df["ANTIGUEDAD"] = df["ANTIGUEDAD"].astype(str).str.strip()

# ------------------------------
# NUMÉRICOS / PRIMA TOTAL
# -----------------------------

# --- LIMPIAR PRIMATOTAL ---

from decimal import Decimal

def limpiar_prima(x):
    if pd.isna(x):
        return None

    x = str(x).strip()

    # Solo permitir números
    x = re.sub(r"[^0-9]", "", x)

    if x == "":
        return None

    try:
        return Decimal(x)
    except:
        return None


if "PRIMA_TOTAL" in df.columns:
    df["PrimaTotal"] = df["PRIMA_TOTAL"].apply(limpiar_prima)
else:
    df["PrimaTotal"] = None



# ENDOSO

#//////////////
# LIMPIAR ENDOSO ANTES DEL RENOMBRE
def limpiar_endoso(x):
    if pd.isna(x):
        return None
    x = str(x).strip()

    # Vacíos → NULL
    if x == "" or x.upper() in ["NAN", "NONE", ".", "-", "--"]:
        return None

    # Solo números
    x = re.sub(r"[^0-9]", "", x)

    if x == "":
        return None

    return int(x)

if "ENDOSO" in df.columns:
    df["ENDOSO"] = df["ENDOSO"].apply(limpiar_endoso)
else:
    print("La columna ENDOSO no existe en el Excel.")


    

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

def clean_date(x):
    if pd.isna(x):
        return None
    if str(x).strip() in ["", " ", "0000-00-00", "00/00/0000", "NaT"]:
        return None
    return x

df["FechaEmision"] = df["FechaEmision"].apply(clean_date)
df["FechaVigenciaDesde"] = df["FechaVigenciaDesde"].apply(clean_date)
df["FechaVigenciaHasta"] = df["FechaVigenciaHasta"].apply(clean_date)

errores_fecha = df[df["FechaEmision"].isna()]
print("Filas con fecha inválida:")
print(errores_fecha[["FechaEmision"]])


insert_query = f"""
INSERT INTO {TABLE_NAME} (
    CodIntermediario, Intermediario, Moneda, Sector, Sucursal, Asegurado, Ramo,
    NumeroFactura, Poliza, Endoso, FECHA_EMISION, Antiguedad,
    FechaVigenciaDesde, FechaVigenciaHasta, PrimaTotal
)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""

print(df["FECHA_EMISION"].head())



count = 0
print("Iniciando inserciones...")

for index, row in df.iterrows():
    try:
        cursor.execute(insert_query,
            row.get("CodIntermediario"),
            row.get("Intermediario"),
            row.get("Moneda"),
            row.get("Sector"),
            row.get("Sucursal"),
            row.get("Asegurado"),
            row.get("Ramo"),
            row.get("NumeroFactura"),
            row.get("Poliza"),
            row.get("Endoso"),
            row.get("FechaEmision"),
            row.get("Antiguedad"),
            row.get("FechaVigenciaDesde"),
            row.get("FechaVigenciaHasta"),
            row["PrimaTotal"]

        )
        count += 1
    except Exception as e:

        print(f"Error al insertar fila {index}: {e}")

        print("❌ ERROR:")
    print((
        row.get("CodIntermediario"),
        row.get("Intermediario"),
        row.get("Moneda"),
        row.get("Sector"),
        row.get("Sucursal"),
        row.get("Asegurado"),
        row.get("Ramo"),
        row.get("NumeroFactura"),
        row.get("Poliza"),
        row.get("Endoso"),
        row.get("FechaEmision"),
        row.get("Antiguedad"),
        row.get("FechaVigenciaDesde"),
        row.get("FechaVigenciaHasta"),
        row.get("PrimaTotal")
    ))
    print("===============================================")

conn.commit()
cursor.close()
conn.close()

print(f" Inserción finalizada. Registros insertados: {count}")
print("=== Automatización finalizada ===")