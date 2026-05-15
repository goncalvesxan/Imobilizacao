# Motor de Imobilização SAP - MVP V2

Piloto em Python + Streamlit para substituir o formulário/macro do Excel e iniciar a montagem do template AS01 com CJI3 + tabelas auxiliares.

## Como executar

```bash
cd app_imobilizacao_mvp_v2
pip install -r requirements.txt
streamlit run app.py
```

## O que mudou na V2

- Mantém upload da CJI3.
- Inclui upload das tabelas auxiliares: Chamados, Classes, Vida Útil, Empresas, Localizações e Razão.
- Tela no formato próximo ao formulário antigo: cadastro à esquerda, filtro/tabela CJI3 à direita e itens salvos abaixo.
- Enriquecimento automático dos campos do template AS01 usando PEP, empresa, centro, classe, chamado e vida útil.
- Botões para salvar decisão, excluir decisão e exportar pacote Excel.
- Estrutura preparada para motor de regras/rateio nas próximas versões.

## Fluxo sugerido

1. Carregar CJI3 e tabelas auxiliares.
2. Revisar diagnóstico das bases.
3. Ir para Cadastro / Decisão.
4. Filtrar por Fabricante/PEP ou texto.
5. Selecionar uma linha CJI3.
6. Confirmar/ajustar campos do ativo.
7. Salvar decisão.
8. Exportar Template AS01.
