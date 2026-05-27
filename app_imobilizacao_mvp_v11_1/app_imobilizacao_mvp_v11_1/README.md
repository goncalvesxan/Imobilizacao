# App Imobilização MVP V10.8

Versão com integração entre Gestão de Chamados e Cadastro/Decisão.

## Rodar
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Principais melhorias V10.8
- A aba 2 Gestão de Chamados agora marca PEPs como “Cadastrado no Ativo” quando houver decisão salva na aba 3 para IDs_CJI3 vinculados ao PEP.
- Inclusão do campo Status Ativo na base analítica e no detalhe do chamado.
- Métrica de PEPs/Registros cadastrados no ativo.
- Exportação da gestão de chamados com lista de PEPs cadastrados no ativo.
