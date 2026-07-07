# Dashboard_porsche_Dados
Um Dashboard de analise de vendas a partir de um banco de dados da porsche, ultilizando sanitizaçao de dados com agentes de ia 

## Estrutura do projeto

- [`data/porsche_sales_raw.xlsx`](data/porsche_sales_raw.xlsx) — dataset bruto original
  (12 colunas, 100 linhas, `sale_id` 6–105), fonte da verdade.
- [`data/porsche_sales_sanitized.csv`](data/porsche_sales_sanitized.csv) — dataset canônico
  sanitizado (22 colunas: cada `*Sanitized` imediatamente após sua coluna bruta), gerado
  pelo script de sanitização em conformidade com o schema.
- [`scripts/sanitize.py`](scripts/sanitize.py) — aplica todas as regras do
  `SchemaPorshe.md` (datas, modelo, ano, preço, milhagem com conversão KM→milhas,
  pagamento, cidade, estado, status de entrega) sobre o dataset bruto e gera o CSV
  canônico.
- [`scripts/build_dashboard_data.py`](scripts/build_dashboard_data.py) — gera o array
  `DATA` embutido no `index.html` a partir do CSV canônico (projeção só com as colunas
  sanitizadas, para exibição).
- [`index.html`](index.html) — dashboard interativo (dados sanitizados embutidos no HTML).
- [`SchemaPorshe.md`](SchemaPorshe.md) — schema de sanitização: colunas de entrada, colunas
  sanitizadas de saída e regras de normalização.
- [`docs/sanitization-log.md`](docs/sanitization-log.md) — histórico da sanitização feita
  originalmente no Excel, o conflito com o schema que foi identificado e como foi resolvido
  (dataset canônico de 22 colunas reconstruído a partir do bruto, com 0 divergências na
  auditoria contra os valores sanitizados anteriores).

## Como regerar os dados

```
python3 scripts/sanitize.py             # raw xlsx -> CSV canônico (22 colunas)
python3 scripts/build_dashboard_data.py # CSV canônico -> DATA embutido no index.html
```
