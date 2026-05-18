# Motor de Imobilização SAP - MVP V9

Versão Streamlit com seleção acumulada de itens CJI3, painel permanente de lote/decisões, marcação visual por cores e bloqueio de duplicidade.

## Como rodar

```bash
cd app_imobilizacao_mvp_v9
pip install -r requirements.txt
streamlit run app.py
```

## Melhorias V9

- Painel permanente do lote atual e das decisões salvas.
- Marcação visual por cor dos itens CJI3:
  - Verde: disponível.
  - Amarelo: no lote atual.
  - Vermelho: já utilizado/bloqueado.
- Mapa visual da CJI3 com linhas coloridas.
- Visual mais profissional com cards, painel e legenda operacional.
- Mantém bloqueio lógico para evitar duplicidade de ID_CJI3.
