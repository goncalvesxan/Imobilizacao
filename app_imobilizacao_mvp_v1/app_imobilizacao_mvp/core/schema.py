from __future__ import annotations

CAMPOS_TEMPLATE_SAP = [
    "Classe", "Empresa", "Denominação", "Descr_02", "Descr_03", "Série", "Inventário",
    "Qtd", "Un_Med", "C.C_SAP", "CENTRO", "Segmento", "Código_FORNC", "NOME_Fornecedor",
    "Fabricante / Pep", "Vida Útil", "Início Depreciação", "Nota Fiscal / Pedido", "Valor Capitalizado",
    "Tipo Decisão", "Tratamento", "Critério Rateio", "Status"
]

COLUNAS_CJI3_CANDIDATAS = {
    "Empresa": ["Empresa", "Empr", "Company Code"],
    "Centro": ["Centro", "CENTRO", "Plant"],
    "PEP": ["Elemento PEP", "PEP", "Elemento do projeto", "Projeto"],
    "Descricao": ["Texto", "Texto do pedido", "Denominação", "Descrição", "Descricao", "Nome"],
    "Classe Custo": ["Classe de custo", "Classe Custo", "Conta", "Conta contábil", "Conta Contabil"],
    "Quantidade": ["Quantidade", "Qtd", "QTD"],
    "Unidade": ["Unidade", "Un.Med", "UMB", "Un_Med", "UM"],
    "Valor": ["Valor", "Valor/MI", "Montante", "Valor Total"],
    "Fornecedor Cod": ["Fornecedor", "Código_FORNC", "Codigo Fornecedor", "Cod Fornecedor"],
    "Fornecedor Nome": ["Nome do fornecedor", "NOME_Fornecedor", "Fornecedor Nome", "Nome Fornecedor"],
    "Pedido": ["Pedido", "Documento de compras", "Pedido Compras"],
    "Nota Fiscal": ["Nota Fiscal", "NF", "Referência", "Referencia", "Doc. referência"]
}

TIPOS_DECISAO = [
    "Produto", "Servico", "Componente", "Principal", "Complemento Ativo Existente", "Nao Capitalizavel", "Pendente"
]

TRATAMENTOS = [
    "Novo ativo unitário", "Novo ativo múltiplo por quantidade", "Ativo composto por agrupamento",
    "Ativo principal com componentes", "Servico incorporado ao ativo", "Complemento de ativo existente",
    "Não capitalizável", "Pendência de análise"
]

CRITERIOS_RATEIO = [
    "Sem rateio", "Por quantidade", "Proporcional por valor", "Percentual manual", "Por PEP",
    "Por centro", "Por ativo principal"
]
