import pandas as pd
import json

# Nome do arquivo Excel (deve estar na raiz do repositório)
excel_file = "planilha.xlsx"

# ------------------------------------------------------------
# 1. Ler a aba 'comp' – composições SINAPI
# ------------------------------------------------------------
# Pula as 3 primeiras linhas (header=None e skiprows=3)
comp = pd.read_excel(excel_file, sheet_name="comp", header=None, skiprows=3)

# Seleciona as colunas por índice (0-based):
# C (índice 2) -> código SINAPI
# E (índice 4) -> descrição
# F (índice 5) -> unidade
# I (índice 8) -> custo material
# J (índice 9) -> custo mão de obra
# L (índice 11) -> tipo item
comp = comp.iloc[:, [2, 4, 5, 8, 9, 11]].copy()
comp.columns = ['codigo', 'descricao', 'unidade', 'custo_material', 'custo_mao_obra', 'tipo']

# Converte para número e preenche nulos
comp['custo_material'] = pd.to_numeric(comp['custo_material'], errors='coerce').fillna(0)
comp['custo_mao_obra'] = pd.to_numeric(comp['custo_mao_obra'], errors='coerce').fillna(0)
comp = comp.dropna(subset=['codigo'])

# ------------------------------------------------------------
# 2. Ler a aba 'sin' – itens do orçamento
# ------------------------------------------------------------
# Pula 13 linhas (header=None e skiprows=13)
sin = pd.read_excel(excel_file, sheet_name="sin", header=None, skiprows=13)

# Seleciona colunas por índice:
# A (0) -> Item
# D (3) -> Cód. SINAPI
# E (4) -> Descrição
# H (7) -> Unid.
# J (9) -> Qtd.
sin = sin.iloc[:, [0, 3, 4, 7, 9]].copy()
sin.columns = ['item', 'codigo', 'descricao', 'unidade', 'quantidade']

sin = sin.dropna(subset=['codigo'])
sin['quantidade'] = pd.to_numeric(sin['quantidade'], errors='coerce').fillna(0)

# ------------------------------------------------------------
# 3. Identificar etapas (baseado na coluna 'item')
# ------------------------------------------------------------
etapas = []
etapa_atual = None
itens_etapa_atual = []

for idx, row in sin.iterrows():
    item_str = str(row['item']) if pd.notna(row['item']) else ''
    if item_str.endswith('.'):
        # É uma etapa
        if etapa_atual is not None:
            etapas.append({
                'id': len(etapas),
                'nome': etapa_atual['nome'],
                'itens': itens_etapa_atual
            })
        etapa_atual = {'nome': row['descricao']}
        itens_etapa_atual = []
    else:
        # É um item de orçamento
        if etapa_atual is not None and pd.notna(row['codigo']):
            itens_etapa_atual.append({
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
        'itens': itens_etapa_atual
    })

# ------------------------------------------------------------
# 4. Parâmetros BDI e encargos (valores fixos da planilha)
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
# 5. Montar JSON final
# ------------------------------------------------------------
dados = {
    "composicoes": comp.to_dict(orient='records'),
    "etapas": etapas,
    "parametros": parametros
}

# Salva como JSON
with open('dados.json', 'w', encoding='utf-8') as f:
    json.dump(dados, f, ensure_ascii=False, indent=2)

print("✅ dados.json gerado com sucesso!")
print(f"Total de composições: {len(comp)}")
print(f"Total de etapas: {len(etapas)}")
