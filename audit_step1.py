import pandas as pd
import warnings, sys, os
warnings.filterwarnings("ignore")

# Usando raw string via os.path para lidar com caracteres especiais
base = os.path.join(
    "C:", os.sep, "Users", "Alexsandro Morais",
    "ABR GERENCIAMENTO E ENGENHARIA LTDA",
    "LYO004 SIMOES FILHO G200 - Documentos",
    "03. CUSTOS",
    "3.0 MEDI\u00c7\u00d5ES APROVADAS E NF'S",
    "1. ROMANEIOS", "0.APROVADOS", "T\u00cdTULOS"
)

p_s = os.path.join(base, "T\u00cdTULOS SIENGE.xlsx")
p_z = os.path.join(base, "T\u00cdTULOS ZEPP.xlsx")
p_r = os.path.join(base, "ROMANEIO.xlsx")
p_c = os.path.join(base, "Conciliacao_T\u00edtulos.xlsx")

print("Paths:")
for p in [p_s, p_z, p_r, p_c]:
    print(f"  {'OK' if os.path.exists(p) else 'MISSING'}: {p}")

sienge = pd.read_excel(p_s, dtype=str)
zepp   = pd.read_excel(p_z, dtype=str)
rom    = pd.read_excel(p_r, dtype=str)
conc   = pd.read_excel(p_c, dtype=str)

print("\nColunas SIENGE:", list(sienge.columns))
print("Colunas ZEPP:", list(zepp.columns))
print("Colunas ROMANEIO:", list(rom.columns))
print("Colunas CONCILIADO:", list(conc.columns))
print(f"\nLinhas: Sienge={len(sienge)} Zepp={len(zepp)} Rom={len(rom)} Conc={len(conc)}")
