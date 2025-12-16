import pandas as pd
import pyodbc

# Ruta del Excel limpio
excel_path = r"C:\Users\ASUS\Documents\EXCEL CARTERA\Balance_Antiguedad_Limpio.xlsx"

# Leer Excel
df = pd.read_excel(excel_path)

# Conexión SQL Server


conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=NANOYOKI-06\\SQLEXPRESS;"
    "DATABASE=ReporteCartera;"
    #"UID=sa;"
   # "PWD=Sh@damy159"
    "Trusted_Connection=yes;"
)

cursor = conn.cursor()

sql_insert = """
INSERT INTO dbo.ReporteCartera (
    CodIntermediario, Intermediario, Moneda, Sector, Sucursal,
    Asegurado, Ramo, NumeroFactura, Poliza, Endoso,
    FechaEmision, Antiguedad, FechaVigenciaDesde,
    FechaVigenciaHasta, PrimaTotal
)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""

for _, row in df.iterrows():
    cursor.execute(sql_insert, tuple(row))

conn.commit()
cursor.close()
conn.close()

print("=== CARGUE DESDE EXCEL COMPLETADO ===")
