# Dashboard de KPIs de TI

Painel front-end para acompanhar indicadores mensais de infraestrutura de TI. O projeto lê uma planilha Excel, organiza os dados por competência e apresenta cartões, gráficos e detalhamento tabular.

## Funcionalidades

- Importação de planilhas `.xlsx` e `.xls`.
- Seleção da competência exibida.
- Gráfico de evolução mensal por indicador.
- Tabela com totais, composição e competência dos dados.
- Exportação do resumo para CSV.

## Executar localmente

Abra `KPI - INVENTÁRIO/index.html` em um navegador moderno. O painel usa a biblioteca SheetJS via CDN para ler arquivos Excel; por isso, a importação da planilha requer conexão com a internet.

## Dados e publicação

Os dados iniciais estão definidos em `KPI - INVENTÁRIO/script.js`. Confirme a autorização e a classificação desses dados antes de manter o repositório público. Para apresentar o projeto como portfólio, substitua os valores e referências organizacionais por dados fictícios e revise o histórico público do repositório.

## Tecnologias

- HTML5
- CSS3
- JavaScript
- SheetJS
