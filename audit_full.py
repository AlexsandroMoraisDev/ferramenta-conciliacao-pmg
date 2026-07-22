import pandas as pd
import warnings, os
warnings.filterwarnings("ignore")

base = os.path.join("C:", os.sep, "Users", "Alexsandro Morais",
    "ABR GERENCIAMENTO E ENGENHARIA LTDA",
    "LYO004 SIMOES FILHO G200 - Documentos",
    "03. CUSTOS", "3.0 MEDI\u00c7\u00d5ES APROVADAS E NF'S",
    "1. ROMANEIOS", "0.APROVADOS", "T\u00cdTULOS")

p_s = os.path.join(base, "T\u00cdTULOS SIENGE.xlsx")
p_z = os.path.join(base, "T\u00cdTULOS ZEPP.xlsx")
p_r = os.path.join(base, "ROMANEIO.xlsx")
p_c = os.path.join(base, "Conciliacao_T\u00edtulos.xlsx")

sienge = pd.read_excel(p_s, dtype=str)
zepp   = pd.read_excel(p_z, dtype=str)
rom    = pd.read_excel(p_r, dtype=str)
conc   = pd.read_excel(p_c, dtype=str)

print("="*70)
print("AUDITORIA COMPLETA - ABA TITULOS")
print("="*70)

# Colunas chave (com caracteres corrigidos)
S_ID = "T\u00edtulo"   # Sienge
Z_ID = "C\u00f3digo Origem"  # Zepp
R_ID = "T\u00cdTULOS SIENGE"  # Romaneio
C_ID = "T\u00edtulo"   # Conciliado
S_CREDOR = "Credor"

# Normalizar IDs
s_ids_raw = sienge[S_ID].fillna('').astype(str).str.strip()
z_ids_raw = zepp[Z_ID].fillna('').astype(str).str.split('/').str[0].str.strip()
r_ids_raw = rom[R_ID].fillna('').astype(str).str.strip()
c_ids_raw = conc[C_ID].fillna('').astype(str).str.strip()

s_ids = set(s_ids_raw) - {'', 'nan'}
z_ids = set(z_ids_raw) - {'', 'nan'}
r_ids = set(r_ids_raw) - {'', 'nan'}
c_ids = set(c_ids_raw) - {'', 'nan'}

print(f"\n[INFO] Totais:")
print(f"  Sienge:     {len(sienge)} linhas -> {len(s_ids)} IDs unicos")
print(f"  Zepp:       {len(zepp)} linhas -> {len(z_ids)} IDs unicos (antes do /)")
print(f"  Romaneio:   {len(rom)} linhas -> {len(r_ids)} IDs unicos")
print(f"  Conciliado: {len(conc)} linhas -> {len(c_ids)} IDs unicos")

# BUG 1: Guard clause - linhas ignoradas silenciosamente
print("\n[BUG 1] Guard clause linha 142 do engine:")
print("  Logica: se Titulo vazio E Credor vazio -> ignora linha")
sem_titulo = sienge[s_ids_raw == '']
s_credor_col = "Credor"
ignoradas = sem_titulo[sem_titulo[s_credor_col].fillna('').astype(str).str.strip() == '']
so_credor  = sem_titulo[sem_titulo[s_credor_col].fillna('').astype(str).str.strip() != '']
print(f"  Linhas SEM titulo E SEM credor (IGNORADAS): {len(ignoradas)}")
print(f"  Linhas SEM titulo MAS COM credor (processadas via credor+valor): {len(so_credor)}")

# BUG 2: Type mismatch
print("\n[BUG 2] Type mismatch - tipos de dado na coluna Titulo (sem dtype=str):")
sienge_nocast = pd.read_excel(p_s)
zepp_nocast   = pd.read_excel(p_z)
tipos_s = sienge_nocast[S_ID].apply(lambda x: type(x).__name__).value_counts()
tipos_z = zepp_nocast[Z_ID].apply(lambda x: type(x).__name__).value_counts()
print(f"  Sienge coluna '{S_ID}': {tipos_s.to_dict()}")
print(f"  Zepp   coluna '{Z_ID}': {tipos_z.to_dict()}")

# Amostra valores raw (pode ter .0)
s_floats = sienge_nocast[sienge_nocast[S_ID].apply(lambda x: isinstance(x, float))]
if not s_floats.empty:
    print(f"  ATENCAO: {len(s_floats)} titulos no Sienge sao float! Exemplos: {list(s_floats[S_ID].head(5))}")
    print(f"  -> String(float) gera '97588.0' vs '97588' no Zepp = FALHA DE MATCH!")

# Verificar 97588 especificamente
print("\n[ALVO] Titulo 97588:")
alvo = '97588'
# Busca exata e com .0
s97  = sienge[s_ids_raw.isin([alvo, alvo+'.0', alvo+',0'])]
s97n = sienge_nocast[sienge_nocast[S_ID].astype(str).str.strip().isin([alvo, alvo+'.0', str(float(alvo))])]
z97  = zepp[z_ids_raw == alvo]
r97  = rom[r_ids_raw == alvo]
c97  = conc[c_ids_raw.isin([alvo, alvo+'.0'])]

print(f"  Sienge (dtype=str): {len(s97)} linha(s) | raw_id: {list(sienge[S_ID][s97.index])}")
print(f"  Sienge (sem cast):  {len(s97n)} linha(s) | raw_val: {list(sienge_nocast[S_ID][s97n.index])}")
print(f"  Zepp:               {len(z97)} linha(s)")
print(f"  Romaneio:           {len(r97)} linha(s)")
print(f"  Conciliado:         {len(c97)} linha(s) -> {'AUSENTE (OCULTADO)' if len(c97)==0 else 'PRESENTE'}")

if not s97.empty:
    row = s97.iloc[0]
    titulo_val = str(row[S_ID]).strip()
    credor_val = str(row[S_CREDOR]).strip() if S_CREDOR in row else ''
    has_titulo = titulo_val not in ('', 'nan')
    has_credor = credor_val not in ('', 'nan')
    guard_passa = has_titulo or has_credor
    print(f"  Guard clause: titulo='{titulo_val}' credor='{credor_val}' -> passa={guard_passa}")
    
    # Verifica se titulo bate com zepp
    z_match = z_ids_raw == titulo_val
    print(f"  Match Zepp exato '{titulo_val}': {z_match.sum()} registros")
    
    if not z97.empty:
        print(f"  Status no Zepp: '{z97.iloc[0].get('Status','?')}'")

# Quantidade de titulos ocultos
print("\n[FALHAS CRITICAS] Titulos no Sienge AUSENTES no Conciliado:")
faltando = s_ids - c_ids
# Remover IDs vazios/nan
faltando = {t for t in faltando if t and t != 'nan'}
print(f"  TOTAL OCULTOS: {len(faltando)}")
faltando_info = []
for t in faltando:
    em_z = t in z_ids
    em_r = t in r_ids
    faltando_info.append((t, em_z, em_r))
# Ordenar
try:
    faltando_info.sort(key=lambda x: float(x[0]))
except:
    faltando_info.sort(key=lambda x: x[0])

print(f"\n  {'Titulo':<12} {'Zepp':>6} {'Rom':>5}  Diagnostico")
print(f"  {'-'*50}")
for t, em_z, em_r in faltando_info:
    z_s = 'SIM' if em_z else 'NAO'
    r_s = 'SIM' if em_r else 'NAO'
    if em_z and em_r:   diag = "Conciliado OK mas OCULTADO no export"
    elif em_z:          diag = "Esta no Zepp, falta Romaneio"
    elif em_r:          diag = "Esta no Romaneio, falta Zepp"
    else:               diag = "Falta nos dois - orphao Sienge"
    print(f"  {t:<12} {z_s:>6} {r_s:>5}  {diag}")

# Duplicatas
print("\n[DUPLICATAS]")
def show_dupes(series, nome):
    d = series[series != ''][series.duplicated(keep=False)]
    if not d.empty:
        uniq = d.unique()
        print(f"  {nome}: {len(uniq)} IDs duplicados ({len(d)} linhas afetadas)")
        for u in uniq[:8]:
            print(f"    '{u}' aparece {list(d).count(u)}x")
    else:
        print(f"  {nome}: sem duplicatas")

show_dupes(s_ids_raw, "Sienge")
show_dupes(z_ids_raw, "Zepp (Codigo Origem)")
show_dupes(r_ids_raw, "Romaneio")
show_dupes(c_ids_raw, "Conciliado")

# Orphaos no Zepp
print("\n[ZEPP ORFAOS] No Zepp mas NAO no Sienge:")
orfaos_z = z_ids - s_ids - {'', 'nan'}
print(f"  Total: {len(orfaos_z)}")
try:    of_s = sorted(orfaos_z, key=float)
except: of_s = sorted(orfaos_z)
for o in of_s[:20]: print(f"    {o}")
if len(of_s)>20: print(f"    ... e mais {len(of_s)-20}")

# Orphaos no Romaneio
print("\n[ROMANEIO ORFAOS] No Romaneio mas NAO no Sienge:")
orfaos_r = r_ids - s_ids - {'', 'nan'}
print(f"  Total: {len(orfaos_r)}")
try:    or_s = sorted(orfaos_r, key=float)
except: or_s = sorted(orfaos_r)
for o in or_s[:20]: print(f"    {o}")
if len(or_s)>20: print(f"    ... e mais {len(or_s)-20}")

# BUG 3: fallback credor+valor
print("\n[BUG 3] Fallback credor+valor - risco de falso-positivo:")
print("  O engine usa os 15 primeiros chars do credor + valor como chave de fallback.")
print("  Se dois titulos diferentes tiverem mesmo credor e mesmo valor -> COLISAO!")
valor_col_s = "Valor liquido" if "Valor liquido" in sienge.columns else None
for c in sienge.columns:
    if "valor" in c.lower() and "liq" in c.lower(): valor_col_s = c
if valor_col_s:
    sienge['_cv'] = sienge[S_CREDOR].fillna('').astype(str).str.lower().str.strip().str[:15] + '_' + sienge[valor_col_s].fillna('0').astype(str)
    colisoes = sienge[sienge['_cv'].duplicated(keep=False) & (sienge['_cv'] != '_0')]
    print(f"  Colisoes de fallback no Sienge: {len(colisoes)} linhas ({colisoes['_cv'].nunique()} chaves ambiguas)")
    if not colisoes.empty:
        for cv, grp in list(colisoes.groupby('_cv'))[:5]:
            print(f"    Chave '{cv}': titulos {list(grp[S_ID])}")
else:
    print("  (coluna de valor liquido nao encontrada automaticamente)")

print("\n[FIM]")
print("="*70)
