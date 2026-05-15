# MVP - Motor de Imobilização e Rateio SAP

Aplicação piloto em Python + Streamlit para transformar bases CJI3 e tabelas auxiliares em uma central de decisão de ativos, com geração inicial do template SAP AS01.

## Como rodar

```bash
cd app_imobilizacao_mvp
pip install -r requirements.txt
streamlit run app.py
```

## Fluxo

1. Upload das bases Excel.
2. Validação de colunas obrigatórias.
3. Normalização dos dados.
4. Sugestão inicial de classificação: produto, serviço, componente, complemento ou pendência.
5. Definição do tratamento: ativo unitário, múltiplo, composto, serviço incorporado, complemento etc.
6. Processamento do rateio.
7. Exportação de Template SAP, Log de Validações e Relatório de Rateio.

## Observação

Esta versão é propositalmente parametrizável. As regras de negócio ficam na tela de regras e podem ser evoluídas para banco de dados posteriormente.
