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

print("="*70)
print("ANALISE COMPLETA - EXTRAS DO SCAN REVERSO")
print("="*70)

# IDs no Sienge
s_ids_raw = sienge[S_ID].fillna('').astype(str).str.strip()
s_ids_norm = s_ids_raw.apply(lambda x: str(int(float(x))) if x.replace('.','',1).isdigit() else x)
s_set = set(s_ids_norm) - {'','nan'}

# IDs no Conciliado
c_ids_raw = conc[C_ID].fillna('').astype(str).str.strip()
c_set = set()
for v in c_ids_raw:
    try: c_set.add(str(int(float(v))))
    except: c_set.add(v)
c_set -= {'', 'nan'}

extras = c_set - s_set

# Pegar status de cada extra no Zepp
zepp['_id'] = zepp[Z_ID].fillna('').astype(str).str.split('/').str[0].str.strip()
zepp_by_id = {}
for _, r in zepp.iterrows():
    tid = r['_id']
    if tid not in zepp_by_id:
        zepp_by_id[tid] = []
    zepp_by_id[tid].append(r)

status_count = {}
reprovados = []
aprovados = []
outros_status = []

for t in extras:
    matches = zepp_by_id.get(t, [])
    if matches:
        status = str(matches[0].get('Status', '')).strip()
        # Pega status normalizado
        import unicodedata
        def rm_acc(s):
            return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').lower()
        status_norm = rm_acc(status)
        
        status_count[status] = status_count.get(status, 0) + 1
        
        if 'reprovado' in status_norm:
            reprovados.append({'titulo': t, 'status': status, 'credor': str(matches[0].get('Credor', ''))})
        elif 'aprovado' in status_norm or 'concluido' in status_norm:
            aprovados.append({'titulo': t, 'status': status, 'credor': str(matches[0].get('Credor', ''))})
        else:
            outros_status.append({'titulo': t, 'status': status, 'credor': str(matches[0].get('Credor', ''))})

print(f"\nTotal de extras: {len(extras)}")
print(f"\nDistribuicao por status no Zepp:")
for st, cnt in sorted(status_count.items(), key=lambda x: -x[1]):
    print(f"  '{st}': {cnt} titulos")

print(f"\n{'='*50}")
print(f"REPROVADOS (nao deveriam aparecer): {len(reprovados)}")
print(f"APROVADOS/CONCLUIDOS (sao orfaos reais): {len(aprovados)}")
print(f"OUTROS STATUS: {len(outros_status)}")

if aprovados:
    print(f"\nAPROVADOS que foram removidos do Sienge (orfaos reais):")
    for a in aprovados[:20]:
        print(f"  {a['titulo']:<10} Status='{a['status']}' Credor='{a['credor'][:40]}'")

if outros_status:
    print(f"\nOUTROS STATUS:")
    for a in outros_status[:10]:
        print(f"  {a['titulo']:<10} Status='{a['status']}' Credor='{a['credor'][:40]}'")

# Verifica se titulo 97588 esta no Zepp desta vez
print(f"\n{'='*50}")
print(f"Verificacao 97588 na nova planilha Zepp:")
z97 = zepp[zepp['_id'] == '97588']
print(f"  Encontrado no Zepp: {len(z97)} linha(s)")
if not z97.empty:
    for _, r in z97.iterrows():
        print(f"  Status='{r.get('Status','')}' Credor='{r.get('Credor','')}'")

# Verifica o titulo 97588 no Sienge atual
s97 = sienge[s_ids_norm == '97588']
print(f"  Encontrado no Sienge: {len(s97)} linha(s) (nao deveria estar ausente na nova extracao)")

print(f"\n{'='*50}")
print(f"CONCLUSAO:")
print(f"  O scan reverso adicionou {len(extras)} titulos extras.")
print(f"  Desses, {len(reprovados)} sao REPROVADOS no Zepp -> NAO deveriam aparecer")
print(f"  E {len(aprovados)} sao APROVADOS que foram removidos do Sienge")
print(f"  A CORRECAO e: filtrar status 'Reprovado' no scan reverso")
print(f"  Resultado esperado apos correcao: {len(sienge)} + {len(aprovados)+len(outros_status)} = {len(sienge)+len(aprovados)+len(outros_status)} linhas")
