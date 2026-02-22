import pandas as pd
import json
import os

# Nome do arquivo Excel (deve estar na raiz)
excel_file = "planilha.xlsx"

print(f"📁 Lendo arquivo: {excel_file}")
print(f"📍 Diretório atual: {os.getcwd()}")
print(f"📄 Arquivo existe? {os.path.exists(excel_file)}")

# ------------------------------------------------------------
# Aba 'comp' – composições SINAPI
# ------------------------------------------------------------
comp = pd.read_excel(excel_file, sheet_name="comp", header=None, skiprows=3)
print(f"📊 Linhas brutas na aba comp: {len(comp)}")

# Colunas por índice (0‑based):
# 2 = CÓD. SINAPI, 4 = DESCRICAO, 5 = UNID., 8 = CUSTO MATERIAL, 9 = CUSTO MÃO DE OBRA, 11 = TIPO ITEM
comp = comp.iloc[:, [2, 4, 5, 8, 9, 11]].copy()
comp.columns = ['codigo', 'descricao', 'unidade', 'custo_material', 'custo_mao_obra', 'tipo']

comp['custo_material'] = pd.to_numeric(comp['custo_material'], errors='coerce').fillna(0)
comp['custo_mao_obra'] = pd.to_numeric(comp['custo_mao_obra'], errors='coerce').fillna(0)
comp = comp.dropna(subset=['codigo'])
print(f"✅ Composições válidas: {len(comp)}")

# ------------------------------------------------------------
# Aba 'sin' – itens do orçamento
# ------------------------------------------------------------
sin = pd.read_excel(excel_file, sheet_name="sin", header=None, skiprows=13)
print(f"📊 Linhas brutas na aba sin: {len(sin)}")

# Colunas: 0 = Item, 3 = Cód. SINAPI, 4 = Descrição, 7 = Unid., 9 = Qtd.
sin = sin.iloc[:, [0, 3, 4, 7, 9]].copy()
sin.columns = ['item', 'codigo', 'descricao', 'unidade', 'quantidade']

sin = sin.dropna(subset=['codigo'])
sin['quantidade'] = pd.to_numeric(sin['quantidade'], errors='coerce').fillna(0)
print(f"✅ Itens de orçamento válidos: {len(sin)}")

# ------------------------------------------------------------
# Identificar etapas (baseado na coluna 'item')
# ------------------------------------------------------------
etapas = []
etapa_atual = None
itens_etapa = []

for idx, row in sin.iterrows():
    item_str = str(row['item']) if pd.notna(row['item']) else ''
    if item_str.endswith('.'):          # é uma etapa
        if etapa_atual is not None:
            etapas.append({
                'id': len(etapas),
                'nome': etapa_atual['nome'],
                'itens': itens_etapa
            })
        etapa_atual = {'nome': row['descricao']}
        itens_etapa = []
    else:                               # é um item de orçamento
        if etapa_atual is not None and pd.notna(row['codigo']):
            itens_etapa.append({
                'codigo': row['codigo'],
                'descricao': row['descricao'],
                'unidade': row['unidade'],
                'quantidade': row['quantidade']
            })

# Adiciona a última etapa
if etapa_atual is not None:
    etapas.append({
        'id': len(etapas),
        'nome': etapa_atual['nome'],
        'itens': itens_etapa
    })
print(f"📂 Etapas identificadas: {len(etapas)}")

# ------------------------------------------------------------
# Parâmetros BDI e encargos
# ------------------------------------------------------------
parametros = {
    "adm_central": 0.0389,
    "despesas_fin": 0.0162,
    "garantia": 0.0109,
    "lucro": 0.0705,
    "pis": 0.03,
    "cofins": 0.05,
    "iss": 0.006,
    "encargos_sociais": 0.22
}

# ------------------------------------------------------------
# Montar JSON final
# ------------------------------------------------------------
dados = {
    "composicoes": comp.to_dict(orient='records'),
    "etapas": etapas,
    "parametros": parametros
}

output_file = 'dados.json'
with open(output_file, 'w', encoding='utf-8') as f:
    json.dump(dados, f, ensure_ascii=False, indent=2)

print(f"✅ Arquivo {output_file} criado com sucesso!")
print(f"📏 Tamanho: {os.path.getsize(output_file)} bytes")
