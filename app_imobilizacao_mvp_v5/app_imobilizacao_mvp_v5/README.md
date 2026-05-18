# Motor de Imobilização SAP - MVP V5

Versão piloto em Python + Streamlit para carregar a CJI3 e tabelas auxiliares, selecionar múltiplos itens de forma intuitiva, aplicar tratamento em lote e gerar Template AS01.

## Como rodar

```bash
cd app_imobilizacao_mvp_v5
pip install -r requirements.txt
streamlit run app.py
```

## Melhorias V5

- Quadro CJI3 com visão resumida para seleção.
- Checkbox direto na grade para marcar vários IDs.
- Filtros visuais por Empresa, Centro, Classe de Custo e Fornecedor.
- Métricas rápidas da visão filtrada: registros, quantidade, valor, PEPs e fornecedores.
- Alternância entre visão resumida e todas as colunas.
- Quadro acumulado de itens selecionados para aplicar o tratamento.
