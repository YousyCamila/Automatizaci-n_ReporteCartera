import os
import pyodbc
import pandas as pd
import re
 

FOLDER_PATH = r"C:\Users\PEDROZARTM\Desktop\CARPETA DE TRABAJO - TANIA PEDROZA\2024"
 
# ===== 2. CONEXIÓN A SQL =====
connection_string = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=PRDPVWTPBIDB;"
    "DATABASE=PolizasDB;"
    "UID=usr_app_indemniza;"  
    "PWD=Colombia123;"        
)
 
conn = pyodbc.connect(connection_string)
cursor = conn.cursor()
 
# ===== 3. LEER ARCHIVOS XLSX =====
files = [f for f in os.listdir(FOLDER_PATH) if f.endswith(".xlsx")]
print("Archivos encontrados:", files)

# ------------------------------
# FUNCIONES
# ------------------------------

def extract_number(filename):
    match = re.match(r"(\d+)", filename)
    return int(match.group(1)) if match else 999999


def load_excel(path):
    """Carga solo XLSX."""
    if path.lower().endswith(".xls"):
        raise Exception("Archivo XLS detectado. Convierte manualmente a XLSX.")
    return pd.read_excel(path, engine="openpyxl")


def clean_columns(df):
    df.columns = df.columns.str.strip()

    # Convertir a string todo lo que sea posible
    df = df.astype(str)

    # Limpiar caracteres dañinos
    df = df.apply(lambda col: col.str.replace(r"[\n\r\t]", " ", regex=True).str.strip())

    # Normalizar fechas y horas
    if "Fecha" in df.columns:
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date.astype(str)

    if "Hora" in df.columns:
        df["Hora"] = df["Hora"].astype(str).str.strip()

    return df


def fetch_existing_ids(cursor):
    """Traer IDs ya cargados a la tabla para evitar duplicados."""
    cursor.execute(f"SELECT ID_PROCESS FROM {TABLE_NAME}")
    return {str(row[0]) for row in cursor.fetchall()}


def insert_data(df, cursor, existing_ids):
    nuevos = 0
    saltados = 0

    for _, row in df.iterrows():
        id_process = str(row.get("ID_PROCESS", "")).strip()

        if id_process in existing_ids:
            saltados += 1
            continue   # evitar duplicados

        cursor.execute(
            f"""
            INSERT INTO {TABLE_NAME} (ID_PROCESS, RESPONSE, DATE_LOG, HORA)
            VALUES (?, ?, ?, ?)
            """,
            id_process,
            row.get("RESPONSE", ""),
            row.get("Fecha", ""),
            row.get("Hora", "")
        )

        existing_ids.add(id_process)
        nuevos += 1

    return nuevos, saltados


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

# Obtener IDs que ya existen en SQL
existing_ids = fetch_existing_ids(cursor)

for file in files_sorted:
    file_path = os.path.join(FOLDER_PATH, file)
    print(f"\nCargando archivo: {file_path}")

    try:
        df = load_excel(file_path)
        df = clean_columns(df)

        nuevos, saltados = insert_data(df, cursor, existing_ids)
        conn.commit()

        print(f"✔ Archivo cargado: {nuevos} nuevos, {saltados} duplicados ignorados")

    except Exception as e:
        print(f"❌ Error cargando {file}: {e}")

cursor.close()
conn.close()

print("\n=== Proceso finalizado ===")
