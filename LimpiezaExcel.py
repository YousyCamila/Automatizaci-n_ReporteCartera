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
    df["FECHA EMISION"] = pd.to_datetime(df["FECHA EMISION"], errors="coerce").dt.date

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
# ANTIGUEDAD
# ------------------------------
# Asegurarnos que sea varchar como en Excel
df["ANTIGUEDAD"] = df["ANTIGUEDAD"].astype(str).str.strip()


# ------------------------------
# NUMÉRICOS / PRIMA TOTAL  (VERSIÓN ROBUSTA)
# ------------------------------
if "PRIMA TOTAL" in df.columns or "PRIMA_TOTAL" in df.columns:
    
    # Detectar el nombre real de la columna
    prima_col = "PRIMA TOTAL" if "PRIMA TOTAL" in df.columns else "PRIMA_TOTAL"

    # Limpiar completamente valores raros, comas, puntos, símbolos y espacios
    df["PrimaTotal"] = (
        df[prima_col]
        .astype(str)
        .str.strip()
        .str.replace(r"[^0-9,.-]", "", regex=True)   # dejar solo números y signos
        .str.replace(".", "", regex=False)           # quitar separador de miles
        .str.replace(",", ".", regex=False)          # convertir coma decimal a punto
    )

    # Convertir a número real seguro
    df["PrimaTotal"] = pd.to_numeric(df["PrimaTotal"], errors="coerce").fillna(0).astype(float)

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



# Convierte números y reemplaza valores inválidos por 0
columnas_float = ["PrimaTotal", "ValorFactura", "Antiguedad", "OtroCampoQueSeaFloat"]

for col in columnas_float:
    if col in df.columns:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace(",", "", regex=False)    # Quita comas
            .str.replace("$", "", regex=False)    # Quita símbolos
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)


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
            float(row.get("PrimaTotal") or 0)

        )
        count += 1
    except Exception as e:
        print(f"Error al insertar fila {index}: {e}, valor PrimaTotal: {row.get('PrimaTotal')}")
        print(f"Error al insertar fila {index}: {e}")

conn.commit()
cursor.close()
conn.close()

print(f" Inserción finalizada. Registros insertados: {count}")
print("=== Automatización finalizada ===")
