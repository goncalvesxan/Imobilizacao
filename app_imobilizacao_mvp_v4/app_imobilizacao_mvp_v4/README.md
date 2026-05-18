# Motor de Imobilização SAP - MVP V3

Aplicação piloto em Python + Streamlit para substituir o fluxo pesado do Excel na criação de cadastro de ativos e preparação do template AS01.

## Como rodar

```bash
cd app_imobilizacao_mvp_v3
pip install -r requirements.txt
streamlit run app.py
```

## Bases suportadas

- CJI3
- Chamados / Status
- Classes
- Vida útil / depreciação
- Empresa x Centro x Local
- Empresas
- Razão / Contas do Imobilizado em andamento

## Correções da V3

- Quadro da CJI3 agora exibe todas as colunas do arquivo carregado.
- Campo Empresa é preenchido pela coluna Empresa da CJI3 ao selecionar o ID_CJI3.
- Campo C.C SAP é preenchido pela coluna Centro da CJI3.
- Campo FORNEC é preenchido pela coluna Conta lnçto.contrap. da CJI3.
- Campo Classe virou seleção baseada na Tabela Classes.
- Campo Vida Útil virou seleção baseada na Tabela Vida Útil, com sugestão por Classe + Empresa.
- Exportação final inclui CJI3 enriquecida e CJI3 original padronizada.
