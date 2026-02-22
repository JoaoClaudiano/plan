import pandas as pd
import json

# Carregar as abas relevantes
excel_file = "Planilha-de-Orcamento-de-Obras-5.0-2-3.xlsx"

# Aba 'comp' - composições SINAPI
comp = pd.read_excel(excel_file, sheet_name="comp", header=3)  # pula as primeiras linhas
comp = comp[['CÓD. SINAPI', 'DESCRICAO', 'UNID.', 'CUSTO MATERIAL', 'CUSTO MÃO DE OBRA', 'TIPO ITEM']].copy()
comp.columns = ['codigo', 'descricao', 'unidade', 'custo_material', 'custo_mao_obra', 'tipo']
# Preencher NaN e converter para float
comp['custo_material'] = pd.to_numeric(comp['custo_material'], errors='coerce').fillna(0)
comp['custo_mao_obra'] = pd.to_numeric(comp['custo_mao_obra'], errors='coerce').fillna(0)
comp = comp.dropna(subset=['codigo'])  # remove linhas sem código

# Aba 'sin' - orçamento sintético (itens por etapa)
sin = pd.read_excel(excel_file, sheet_name="sin", header=13)  # começa na linha 14
# Selecionar colunas relevantes: Item, Cód. SINAPI, Descrição, Unid., Qtd.
sin = sin[['Item', 'Cód. SINAPI', 'Descrição', 'Unid.', 'Qtd.']].copy()
sin.columns = ['item', 'codigo', 'descricao', 'unidade', 'quantidade']
sin = sin.dropna(subset=['codigo'])  # remove linhas sem código
sin['quantidade'] = pd.to_numeric(sin['quantidade'], errors='coerce').fillna(0)

# Parâmetros BDI e encargos (da aba sin, células fixas)
# Valores extraídos manualmente (posições fixas)
params = {
    "adm_central": 0.0389,
    "despesas_fin": 0.0162,
    "garantia": 0.0109,
    "lucro": 0.0705,
    "pis": 0.03,
    "cofins": 0.05,
    "iss": 0.006,
    "encargos_sociais": 0.22
}

# Construir estrutura final
dados = {
    "composicoes": comp.to_dict(orient='records'),
    "itens_orcamento": sin.to_dict(orient='records'),
    "parametros": params
}

# Salvar JSON
with open('dados.json', 'w', encoding='utf-8') as f:
    json.dump(dados, f, ensure_ascii=False, indent=2)

print("Arquivo dados.json criado com sucesso!")
