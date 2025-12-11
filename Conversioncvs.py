import os
import pandas as pd
 
# 📁 Carpeta donde están los XLSX
INPUT_FOLDER = r"C:\Users\PEDROZARTM\Desktop\CARPETA DE TRABAJO - TANIA PEDROZA\2025"
 
# 📁 Carpeta donde se guardarán los CSV
OUTPUT_FOLDER = r"C:\Users\PEDROZARTM\Desktop\CARPETA DE TRABAJO - TANIA PEDROZA\2025-CSV"
 
# Crear carpeta si no existe
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
 
print("=== Iniciando conversión masiva XLSX → CSV ===")
 
counter = 0
 
for file in os.listdir(INPUT_FOLDER):
    if file.lower().endswith(".xlsx"):
        xlsx_path = os.path.join(INPUT_FOLDER, file)
 
        # Nombre del nuevo CSV
        csv_filename = file.replace(".xlsx", ".csv")
        csv_path = os.path.join(OUTPUT_FOLDER, csv_filename)
 
        try:
            # Convertir
            df = pd.read_excel(xlsx_path)
            df.to_csv(csv_path, index=False, sep=";")
 
            counter += 1
            print(f"✔ Convertido ({counter}): {file}")
 
        except Exception as e:
            print(f"❌ Error en {file}: {e}")
 
print(f"\n=== Conversión finalizada. Total archivos convertidos: {counter} ===")