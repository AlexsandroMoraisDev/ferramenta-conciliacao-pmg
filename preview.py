import pandas as pd
import json
import sys

def preview_excel(file_path):
    try:
        df = pd.read_excel(file_path, nrows=5)
        print(f"--- {file_path} ---")
        print("Columns:", list(df.columns))
        print("First row:", df.iloc[0].to_dict() if len(df) > 0 else "Empty")
    except Exception as e:
        print(f"Error reading {file_path}: {e}")

preview_excel(r"c:\Users\Alexsandro Morais\Documents\workspace\projetos\Ferramenta de Conciliação\0.APROVADOS\TÍTULOS\ROMANEIO.xlsx")
preview_excel(r"c:\Users\Alexsandro Morais\Documents\workspace\projetos\Ferramenta de Conciliação\0.APROVADOS\TÍTULOS\TÍTULOS SIENGE.xlsx")
preview_excel(r"c:\Users\Alexsandro Morais\Documents\workspace\projetos\Ferramenta de Conciliação\0.APROVADOS\TÍTULOS\TÍTULOS ZEPP.xlsx")
