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
p_c = os.path.join(base, "Conciliacao_T\u00edtulos.xlsx")

sienge = pd.read_excel(p_s, dtype=str)
zepp   = pd.read_excel(p_z, dtype=str)
conc   = pd.read_excel(p_c, dtype=str)

S_ID = "T\u00edtulo"
Z_ID = "C\u00f3digo Origem"
C_ID = "T\u00edtulo"
C_ACAO = "A\u00e7\u00e3o Requerida"
C_OBS  = "Observa\u00e7\u00e3o"

print("="*70)
print("AUDITORIA v2 - DIAGNOSTICO DAS 2 FALHAS")
print("="*70)

# Normalizar IDs
s_ids_raw = sienge[S_ID].fillna('').astype(str).str.strip()
s_ids_norm = s_ids_raw.apply(lambda x: str(int(float(x))) if x.replace('.','',1).replace('-','',1).isdigit() else x)
s_set = set(s_ids_norm) - {'','nan'}

z_ids_raw = zepp[Z_ID].fillna('').astype(str).str.split('/').str[0].str.strip()
z_set = set(z_ids_raw) - {'','nan'}

c_ids_raw = conc[C_ID].fillna('').astype(str).str.strip()

print(f"\n[INFO] Sienge: {len(sienge)} linhas | {len(s_set)} IDs unicos")
print(f"[INFO] Zepp:   {len(zepp)} linhas | {len(z_set)} IDs unicos")
print(f"[INFO] Conciliado: {len(conc)} linhas")
print(f"[INFO] Diferenca: {len(conc) - len(sienge)} linhas extras no conciliado")

# =========================================================
# FALHA 1: Por que 907 e nao 736?
# O scan reverso adicionou entradas extras do Zepp
# =========================================================
print("\n" + "="*70)
print("[FALHA 1] ANALISE DAS LINHAS EXTRAS DO SCAN REVERSO")
print("="*70)

# Pegar linhas do conciliado marcadas como "No Zepp" (adicionadas pelo scan reverso)
mask_scan_reverso = conc[C_OBS].fillna('').str.contains('N\u00e3o encontrado na extra\u00e7\u00e3o do Sienge', na=False)
scan_reverso_rows = conc[mask_scan_reverso]
print(f"\n  Total de linhas adicionadas pelo scan reverso: {len(scan_reverso_rows)}")

# Dessas, quantas REALMENTE existem no Sienge atual?
presentes_sienge = 0
ausentes_sienge  = 0
rows_presente = []
rows_ausente  = []

for _, row in scan_reverso_rows.iterrows():
    tid_conc = str(row[C_ID]).strip()
    # Normalizar
    try:
        tid_norm = str(int(float(tid_conc)))
    except:
        tid_norm = tid_conc
    
    if tid_norm in s_set or tid_conc in s_set:
        presentes_sienge += 1
        rows_presente.append({'titulo': tid_conc, 'credor': row.get('Credor',''), 'acao': row.get(C_ACAO,'')})
    else:
        ausentes_sienge += 1
        rows_ausente.append({'titulo': tid_conc, 'credor': row.get('Credor',''), 'acao': row.get(C_ACAO,'')})

print(f"  Dessas linhas extras:")
print(f"    - Que EXISTEM no Sienge (falso positivo do scan reverso): {presentes_sienge}")
print(f"    - Que NAO existem no Sienge (legit orfaos do Zepp):       {ausentes_sienge}")

if rows_presente:
    print(f"\n  TITULOS QUE EXISTEM NO SIENGE MAS FORAM ADICIONADOS COMO ORFAOS:")
    print(f"  {'Titulo':<12} {'Credor':<40} Acao")
    for r in rows_presente[:20]:
        print(f"  {r['titulo']:<12} {str(r['credor'])[:40]:<40} {r['acao']}")

# =========================================================
# DIAGNOSTICO: Por que o matching falhou para esses titulos?
# =========================================================
print(f"\n[DIAGNOSTICO] Por que titulos do Sienge foram marcados como orfaos do Zepp?")
print(f"  O matching compara: String(siengeRow['Titulo']).trim() vs Zepp.split('/')[0].trim()")
print(f"  Verificando os IDs do Sienge:")

# Mostrar amostra dos IDs normalizados do Sienge
sienge_raw_nocast = pd.read_excel(p_s)
tipos = sienge_raw_nocast[S_ID].apply(lambda x: type(x).__name__).value_counts()
print(f"\n  Tipos de dado no Sienge (coluna Titulo, sem dtype=str): {tipos.to_dict()}")

# Amostra dos valores raw e como seriam convertidos em JS
sample_s = sienge_raw_nocast[S_ID].dropna().head(10)
print(f"\n  Amostra valores raw Sienge:")
for v in sample_s:
    js_str = str(int(v)) if isinstance(v, float) else str(v)
    print(f"    raw={repr(v)} (type={type(v).__name__}) -> JS String() = '{js_str}'")

# Verificar IDs do Zepp - formato
print(f"\n  Amostra Zepp Codigo Origem (10 primeiros):")
for v in zepp[Z_ID].head(10):
    extracted = str(v).split('/')[0].strip()
    print(f"    Codigo='{v}' -> extraido='{extracted}'")

# =========================================================
# FALHA 2: 907-736 = 171 extras. Quais sao esses 171?
# =========================================================
print("\n" + "="*70)
print("[FALHA 2] ANALISE DOS 171 TITULOS EXTRAS (Sienge=736, Conciliado=907)")
print("="*70)

# IDs no conciliado que nao estao no Sienge
c_set_norm = set()
for v in c_ids_raw:
    try: c_set_norm.add(str(int(float(v))))
    except: c_set_norm.add(v)
c_set_norm -= {'', 'nan'}

extras = c_set_norm - s_set
print(f"\n  IDs no Conciliado mas NAO no Sienge: {len(extras)}")

# Verificar quais desses extras estao no Zepp
extras_info = []
for t in extras:
    em_zepp = t in z_set
    # Verificar o status no Zepp
    status_z = ''
    credor_z = ''
    cod_origem_z = ''
    if em_zepp:
        match = zepp[zepp[Z_ID].fillna('').astype(str).str.split('/').str[0].str.strip() == t]
        if not match.empty:
            status_z = match.iloc[0].get('Status', '')
            credor_z = match.iloc[0].get('Credor', '')
            cod_origem_z = match.iloc[0].get(Z_ID, '')
    # Verificar se e numero puro
    is_numeric = t.replace('.','',1).isdigit()
    
    # Verificar o range do Sienge
    sienge_min = min(float(x) for x in s_set if x.replace('.','',1).isdigit())
    sienge_max = max(float(x) for x in s_set if x.replace('.','',1).isdigit())
    in_range = False
    try:
        in_range = sienge_min <= float(t) <= sienge_max
    except:
        pass
    
    extras_info.append({
        'titulo': t,
        'em_zepp': em_zepp,
        'status_zepp': status_z,
        'credor': credor_z,
        'cod_origem': cod_origem_z,
        'is_numeric': is_numeric,
        'in_range': in_range
    })

# Classificar
numeric_in_range = [e for e in extras_info if e['is_numeric'] and e['in_range']]
numeric_out_range = [e for e in extras_info if e['is_numeric'] and not e['in_range']]
non_numeric = [e for e in extras_info if not e['is_numeric']]

print(f"\n  Classificacao dos {len(extras)} extras:")
print(f"    Numericos dentro do range Sienge ({int(sienge_min)}-{int(sienge_max)}): {len(numeric_in_range)} <- POSSIVEIS ORFAOS REAIS")
print(f"    Numericos FORA do range Sienge:                                         {len(numeric_out_range)} <- OUTROS PROJETOS")
print(f"    Nao numericos (outros projetos/sistemas):                               {len(non_numeric)} <- OUTROS SISTEMAS")

if numeric_in_range:
    print(f"\n  Titulos numericos no range (possiveis orfaos reais):")
    for e in numeric_in_range[:20]:
        print(f"    {e['titulo']:<10} Status='{e['status_zepp']}' Credor='{str(e['credor'])[:30]}' Origem='{e['cod_origem']}'")

if numeric_out_range[:10]:
    print(f"\n  Exemplos numericos FORA do range (outros projetos):")
    for e in numeric_out_range[:10]:
        print(f"    {e['titulo']:<10} Status='{e['status_zepp']}' Credor='{str(e['credor'])[:30]}'")

if non_numeric[:10]:
    print(f"\n  Exemplos NAO numericos (outros sistemas):")
    for e in non_numeric[:10]:
        print(f"    {e['titulo']:<15} Status='{e['status_zepp']}' Credor='{str(e['credor'])[:30]}'")

# =========================================================
# RECOMENDACAO DE FILTRO
# =========================================================
print(f"\n" + "="*70)
print("[RECOMENDACAO] Filtros necessarios no scan reverso:")
print("="*70)
print(f"  1. So incluir IDs numericos puros (descartar 'CLSF/XXX', 'PPC/XXX', etc.)")
print(f"  2. So incluir IDs dentro do range do Sienge ({int(sienge_min)}-{int(sienge_max)})")
print(f"     OU com margem de +/- 5000 numeros do range")
print(f"  3. Manter filtro de cancelado")
print(f"\n  Com esses filtros, o scan reverso adicionaria apenas:")
legit = [e for e in extras_info if e['is_numeric'] and e['in_range']]
print(f"    {len(legit)} titulo(s) legit ao inves de {len(extras)}")

print("\n[FIM]")
