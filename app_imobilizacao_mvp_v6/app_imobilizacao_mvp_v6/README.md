# Motor de Imobilização SAP - MVP V6

Versão piloto em Python + Streamlit para carregar a CJI3 e tabelas auxiliares, gerir chamados, selecionar múltiplos itens CJI3, aplicar tratamento em lote e gerar Template AS01.

## Como rodar

```bash
cd app_imobilizacao_mvp_v6
pip install -r requirements.txt
streamlit run app.py
```

## Melhorias V6

- Nova aba **2. Gestão de Chamados**.
- Filtros de chamados por Empresa, Centro, Status, Classe e texto geral.
- Indicadores de chamados filtrados, PEPs, status e cruzamento com CJI3.
- Coluna **Existe na CJI3** para identificar PEPs com lançamentos.
- Análise detalhada por PEP, exibindo dados do chamado e lançamentos CJI3 vinculados.
- Exportação da gestão de chamados para Excel.
- Mantidas as melhorias da V5: seleção intuitiva da CJI3, checkbox, quadro acumulado e tratamento em lote.
