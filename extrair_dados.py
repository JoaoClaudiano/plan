import pandas as pd
import json

# Nome do arquivo Excel (fixo, para o workflow)
excel_file = "planilha.xlsx"

# ------------------------------------------------------------
# 1. Ler a aba 'comp' (composições SINAPI)
# ------------------------------------------------------------
comp = pd.read_excel(excel_file, sheet_name="comp", header=3)
comp = comp[['CÓD. SINAPI', 'DESCRICAO', 'UNID.', 'CUSTO MATERIAL', 'CUSTO MÃO DE OBRA', 'TIPO ITEM']].copy()
comp.columns = ['codigo', 'descricao', 'unidade', 'custo_material', 'custo_mao_obra', 'tipo']
comp['custo_material'] = pd.to_numeric(comp['custo_material'], errors='coerce').fillna(0)
comp['custo_mao_obra'] = pd.to_numeric(comp['custo_mao_obra'], errors='coerce').fillna(0)
comp = comp.dropna(subset=['codigo'])

# ------------------------------------------------------------
# 2. Ler a aba 'sin' (itens do orçamento)
# ------------------------------------------------------------
sin = pd.read_excel(excel_file, sheet_name="sin", header=13)
sin = sin[['Item', 'Cód. SINAPI', 'Descrição', 'Unid.', 'Qtd.']].copy()
sin.columns = ['item', 'codigo', 'descricao', 'unidade', 'quantidade']
sin = sin.dropna(subset=['codigo'])
sin['quantidade'] = pd.to_numeric(sin['quantidade'], errors='coerce').fillna(0)

# ------------------------------------------------------------
# 3. Identificar etapas
# ------------------------------------------------------------
etapas = []
etapa_atual = None
itens_etapa_atual = []

for idx, row in sin.iterrows():
    item_str = str(row['item']) if pd.notna(row['item']) else ''
    if item_str.endswith('.'):
        if etapa_atual is not None:
            etapas.append({
                'id': len(etapas),
                'nome': etapa_atual['nome'],
                'itens': itens_etapa_atual
            })
        etapa_atual = {'nome': row['descricao']}
        itens_etapa_atual = []
    else:
        if etapa_atual is not None and pd.notna(row['codigo']):
            itens_etapa_atual.append({
                'codigo': row['codigo'],
                'descricao': row['descricao'],
                'unidade': row['unidade'],
                'quantidade': row['quantidade']
            })

if etapa_atual is not None:
    etapas.append({
        'id': len(etapas),
        'nome': etapa_atual['nome'],
        'itens': itens_etapa_atual
    })

# ------------------------------------------------------------
# 4. Parâmetros BDI e encargos
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

with open('dados.json', 'w', encoding='utf-8') as f:
    json.dump(dados, f, ensure_ascii=False, indent=2)

print("dados.json gerado com sucesso!")
