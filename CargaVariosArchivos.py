import os
import re
import pandas as pd
import pyodbc

# ------------------------------
# CONFIGURACIONES
# ------------------------------

FOLDER_PATH = r"C:\Users\ASUS\Documents\prueba_automatizacionlogs_TANIA"
SERVER = "NANOYOKI-06\\SQLEXPRESS"
DATABASE = "ReporteCartera"
TABLE_NAME = "LogProcesos"

connection_string = (
    f"DRIVER={{SQL Server}};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    "Trusted_Connection=yes;"
)

# ------------------------------
# FUNCIONES
# ------------------------------

def extract_number(filename):
    match = re.match(r"(\d+)", filename)
    return int(match.group(1)) if match else 999999


def load_excel(path):
    """ Carga solo XLSX. Si es XLS, lo rechaza. """
    if path.lower().endswith(".xls"):
        raise Exception("Archivo XLS detectado. Convierte manualmente a XLSX.")
    return pd.read_excel(path, engine="openpyxl")


def clean_columns(df):
    df.columns = df.columns.str.strip()

    # Convertir todo lo posible a string para evitar errores en ODBC
    df = df.astype(str)

    # Quitar saltos y espacios
    df = df.apply(lambda col: col.str.replace(r"[\n\r\t]", " ", regex=True).str.strip())

    return df


def insert_data(df, cursor):
    for _, row in df.iterrows():
        cursor.execute(
            f"""
            INSERT INTO {TABLE_NAME} (ID_PROCESS, RESPONSE, DATE_LOG, HORA)
            VALUES (?, ?, ?, ?)
            """,
            row.get("ID_PROCESS", ""),
            row.get("RESPONSE", ""),
            row.get("Fecha", ""),
            row.get("Hora", "")
        )


# ------------------------------
# PROCESO PRINCIPAL
# ------------------------------

print("\n=== Iniciando carga masiva de Excel 2024 ===")

files = [f for f in os.listdir(FOLDER_PATH) if f.endswith(".xlsx")]
files_sorted = sorted(files, key=extract_number)

print("\nArchivos encontrados en orden:")
for f in files_sorted:
    print(" ->", f)

conn = pyodbc.connect(connection_string)
cursor = conn.cursor()

for file in files_sorted:
    file_path = os.path.join(FOLDER_PATH, file)
    print(f"\nCargando archivo: {file_path}")

    try:
        df = load_excel(file_path)
        df = clean_columns(df)
        insert_data(df, cursor)
        conn.commit()
        print("✔ Archivo cargado con éxito")

    except Exception as e:
        print(f"❌ Error cargando {file}: {e}")

cursor.close()
conn.close()

print("\n=== Proceso finalizado ===")
