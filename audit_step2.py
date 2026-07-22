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

# Ler sem dtype=str para ver tipos reais
sienge_raw = pd.read_excel(p_s)
zepp_raw   = pd.read_excel(p_z)
rom_raw    = pd.read_excel(p_r)
conc_raw   = pd.read_excel(p_c)

S_ID = "T\u00edtulo"
Z_ID = "C\u00f3digo Origem"
R_ID = "T\u00cdTULOS SIENGE"
C_ID = "T\u00edtulo"

print("="*70)
print("INVESTIGACAO PROFUNDA - ROOT CAUSE ANALISE")
print("="*70)

# --- O BUG REAL ---
print("\n[ROOT CAUSE] XLSX lê numeros como float. O engine.js usa XLSX.utils.sheet_to_json")
print("  com {raw: true}. Isso faz o mesmo: retorna numero JS (97588) sem aspas.")
print("  Porem a linha do Zepp armazena '97588' como STRING (texto).")
print()
print("  No engine.js linha 143:")
print("    const tituloSienge = String(siengeRow['Titulo'] || '').trim()")
print("    -> String(97588) = '97588'   # correto")
print("    -> String(97588.0) = ... em JS: '97588' (JS nao tem float vs int como Python)")
print()
print("  Portanto o BUG DE TIPO nao ocorre em JS diretamente!")
print("  A causa do sumiço do 97588 deve ser outra.")
print()

# Verificar o valor raw do 97588 no Sienge
print("[INVESTIGACAO] Procurando 97588 no Sienge (raw):")
col_titulo = S_ID
rows_97588 = sienge_raw[sienge_raw[col_titulo] == 97588]
rows_97588_str = sienge_raw[sienge_raw[col_titulo].astype(str).str.strip() == '97588']
rows_97588_float = sienge_raw[sienge_raw[col_titulo].astype(str).str.strip() == '97588.0']

print(f"  Busca numerica (== 97588): {len(rows_97588)} linha(s)")
print(f"  Busca string  ('97588'):   {len(rows_97588_str)} linha(s)")
print(f"  Busca float   ('97588.0'): {len(rows_97588_float)} linha(s)")

if rows_97588.empty and rows_97588_str.empty:
    print("  *** 97588 NAO EXISTE NO SIENGE ATUAL! ***")
    print("  O titulo pode ter sido removido/pago/arquivado na planilha que foi importada.")
    
    # Verificar o range de titulos no Sienge
    titulos_num = pd.to_numeric(sienge_raw[col_titulo], errors='coerce').dropna()
    print(f"  Titulos Sienge: min={titulos_num.min():.0f} max={titulos_num.max():.0f}")
    print(f"  97588 esta no range? {titulos_num.min() <= 97588 <= titulos_num.max()}")
    
    # Verificar se existe proximo
    titulos_sort = sorted(titulos_num)
    idx = None
    for i, t in enumerate(titulos_sort):
        if t >= 97585:
            idx = i
            break
    if idx is not None:
        nearby = titulos_sort[max(0,idx-3):idx+5]
        print(f"  Titulos proximos de 97588 no Sienge: {[int(x) for x in nearby]}")

# Verificar no Zepp
print("\n[INVESTIGACAO] Procurando 97588 no Zepp:")
zepp_raw['_id'] = zepp_raw[Z_ID].fillna('').astype(str).str.split('/').str[0].str.strip()
z97 = zepp_raw[zepp_raw['_id'] == '97588']
print(f"  Encontrado: {len(z97)} linha(s)")
if not z97.empty:
    for _, r in z97.iterrows():
        print(f"  Status='{r.get('Status','')}' Codigo='{r.get(Z_ID,'')}' Credor='{r.get('Credor','')}' Valor='{r.get('Valor Origem','')}'")

# Verificar no Romaneio
print("\n[INVESTIGACAO] Procurando 97588 no Romaneio:")
rom_raw['_id'] = rom_raw[R_ID].fillna('').astype(str).str.strip()
r97 = rom_raw[rom_raw['_id'] == '97588']
print(f"  Encontrado: {len(r97)} linha(s)")
if not r97.empty:
    for _, r in r97.iterrows():
        print(f"  Rom='{r.get('N\u00ba ROMANEIO','')}' Razao='{r.get('RAZ\u00c3O SOCIAL','')}' Valor='{r.get('VALOR LIQUIDO+JUROS E MULTAS','')}'")

# Verificar na planilha conciliada gerada
print("\n[INVESTIGACAO] Procurando 97588 no Conciliado:")
conc_raw['_id'] = conc_raw[C_ID].fillna('').astype(str).str.strip()
c97 = conc_raw[conc_raw['_id'].isin(['97588','97588.0'])]
print(f"  Encontrado: {len(c97)} linha(s)")

# Entender a diferença de linhas: Sienge=720 mas Conciliado=719
print("\n[DISCREPANCIA DE CONTAGEM]")
print(f"  Sienge tem 720 linhas mas Conciliado tem apenas 719!")
print(f"  Diferenca: 1 linha a mais no Sienge do que no Conciliado.")
print(f"  IDs unicos Sienge: {len(set(sienge_raw[S_ID].dropna().astype(str).str.strip()) - {'','nan'})}")
print(f"  IDs unicos Conciliado: {len(set(conc_raw[C_ID].dropna().astype(str).str.strip()) - {'','nan'})}")

# Descobrir qual linha está faltando
s_ids_all = sienge_raw[S_ID].dropna().astype(str).str.strip()
c_ids_all = conc_raw[C_ID].dropna().astype(str).str.strip()

# Como JS converte numeros: 97588.0 -> "97588"
# Normalizar: remover .0
s_ids_norm = s_ids_all.apply(lambda x: str(int(float(x))) if x.replace('.','',1).isdigit() else x)
c_ids_norm = c_ids_all.apply(lambda x: str(int(float(x))) if x.replace('.','',1).isdigit() else x)

s_set = set(s_ids_norm) - {'', 'nan'}
c_set = set(c_ids_norm) - {'', 'nan'}

faltando = s_set - c_set
sobra    = c_set - s_set

print(f"\n  Apos normalizacao (remover .0):")
print(f"  Sienge IDs unicos: {len(s_set)}")
print(f"  Conciliado IDs unicos: {len(c_set)}")
print(f"  IDs no Sienge mas NAO no Conciliado: {faltando}")
print(f"  IDs no Conciliado mas NAO no Sienge:  {sobra}")

# Qual linha do Sienge corresponde ao ID faltando?
for f_id in faltando:
    row_f = sienge_raw[s_ids_norm == f_id]
    if not row_f.empty:
        r = row_f.iloc[0]
        titulo_raw = r.get(S_ID, '?')
        credor_raw = r.get('Credor', '?')
        print(f"\n  LINHA FALTANDO NO CONCILIADO:")
        print(f"    Titulo raw = '{titulo_raw}' (type={type(titulo_raw).__name__})")
        print(f"    Credor     = '{credor_raw}'")
        # Verificar se seria afetado pela guard clause
        titulo_str = str(titulo_raw).strip() if titulo_raw is not None else ''
        credor_str = str(credor_raw).strip() if credor_raw is not None else ''
        # Em JS: String(num) de 97588.0 (float) = "97588" - sem problema
        # Verificar se na leitura XLSX.js, o titulo vem como numero
        print(f"    Guard clause: titulo='{titulo_str}' credor='{credor_str}'")
        has_titulo = titulo_str and titulo_str != 'nan' and titulo_str != 'None'
        has_credor = credor_str and credor_str != 'nan' and credor_str != 'None'
        print(f"    has_titulo={has_titulo} has_credor={has_credor} -> passa={has_titulo or has_credor}")
        
        # Verificar no Zepp
        zepp_match_z = zepp_raw[zepp_raw['_id'] == str(int(float(titulo_raw))) if str(titulo_raw).replace('.','',1).isdigit() else zepp_raw['_id'] == str(titulo_raw)]
        print(f"    Zepp match: {len(zepp_match_z)} linhas")
        if not zepp_match_z.empty:
            print(f"    Status Zepp: '{zepp_match_z.iloc[0].get('Status','')}'")
        
        # Verificar no Romaneio
        rom_match = rom_raw[rom_raw['_id'] == str(int(float(titulo_raw))) if str(titulo_raw).replace('.','',1).isdigit() else rom_raw['_id'] == str(titulo_raw)]
        print(f"    Romaneio match: {len(rom_match)} linhas")
        
        # Diagnóstico final
        in_zepp = len(zepp_match_z) > 0
        in_rom  = len(rom_match)  > 0
        status_z = zepp_match_z.iloc[0].get('Status','') if in_zepp else ''
        status_z_lower = status_z.lower()
        
        if in_rom and in_zepp and ('aprovado' in status_z_lower or 'conclu' in status_z_lower):
            acao_esperada = 'OK: Lancado no Romaneio e Enviado'
        elif in_rom and in_zepp:
            acao_esperada = 'ALERTA: Em Aprovacao no Zepp'
        elif in_zepp and not in_rom:
            acao_esperada = 'ALERTA: Falta Romaneio'
        elif in_rom and not in_zepp:
            acao_esperada = 'ALERTA: Falta Zepp'
        else:
            acao_esperada = 'ALERTA: Falta Romaneio e Falta Zepp'
        print(f"    Acao que deveria aparecer: '{acao_esperada}'")

print("\n[FIM DA INVESTIGACAO]")
