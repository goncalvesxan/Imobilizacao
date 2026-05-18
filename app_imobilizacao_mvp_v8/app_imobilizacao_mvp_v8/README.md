# Motor de Imobilização SAP - MVP V8

Versão piloto em Python + Streamlit para carregar a CJI3 e tabelas auxiliares, gerir chamados, selecionar múltiplos itens CJI3, aplicar tratamento em lote, replicar cadastros por QTD e gerar Template AS01.

## Como rodar

```bash
cd app_imobilizacao_mvp_v8
pip install -r requirements.txt
streamlit run app.py
```

## Melhorias V8

- Incluído o flag **Replicar cadastro pela QTD** na tela de Cadastro / Decisão.
- Quando o flag estiver ativado, o cadastro salvo é repetido conforme o campo **QTD**.
- O **Valor Capitalizado** é dividido igualmente entre as linhas replicadas.
- Cada linha replicada recebe rastreabilidade com número da réplica, total de réplicas, origem dos IDs CJI3 e observação automática.
- A tabela **Itens cadastrados / decisões salvas** passa a exibir as linhas repetidas que serão usadas no Template AS01.
- Mantidas as melhorias da V6: gestão de chamados, seleção intuitiva da CJI3, checkbox, quadro acumulado e tratamento em lote.


## Novidades V8
- Marca automaticamente IDs CJI3 já salvos como tratados/bloqueados.
- Evita adicionar novamente ao lote itens já salvos.
- Inclui cartões de status e visual mais profissional.
- Mantém rastreabilidade de origem dos IDs CJI3.
