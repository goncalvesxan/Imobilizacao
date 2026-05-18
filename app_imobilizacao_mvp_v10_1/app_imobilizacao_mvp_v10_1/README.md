# Motor de Imobilização SAP - MVP V10

Versão Streamlit com seleção acumulada de itens CJI3, painel permanente de lote/decisões, marcação visual por cores, bloqueio de duplicidade, gestão de chamados e exportação AS01 padronizada.

## Como rodar

```bash
cd app_imobilizacao_mvp_v10
pip install -r requirements.txt
streamlit run app.py
```

## Melhorias V10

- Aba **Template AS01 / Exportação** reativada e estruturada.
- Geração do pacote Excel final com abas padronizadas:
  - Template_AS01_Padronizado
  - Validacoes_AS01
  - Decisoes_Completas
  - Pendencias
  - CJI3_IDs_Utilizados
  - Parametros_Exportacao
  - Leia_me
  - CJI3_Enriquecida
  - CJI3_Original_Padronizada
- Métricas de decisões salvas, IDs CJI3 tratados e valor capitalizado.
- Validação de campos obrigatórios antes da carga SAP AS01.
- Mantém bloqueio lógico para evitar duplicidade de ID_CJI3.
