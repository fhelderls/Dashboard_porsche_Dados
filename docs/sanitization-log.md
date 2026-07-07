# Log de Sanitização — Agente de IA no Excel

> Registro (traduzido e limpo) da sessão em que um agente de IA, operando via add-in do
> Excel/M365, sanitizou a base de vendas Porsche na aba **"Sanitized"** (tabela `Tabela1`)
> e gerou o dataset que alimenta o [`index.html`](../index.html).

## Estado inicial da planilha

A tabela `Tabela1` tinha 21 colunas em pares bruto → sanitizado:

| Bruta | Sanitizada |
|---|---|
| `sale_id` | — (sem par) |
| `sale_date` | `SaleDateSanitized` |
| `customer_name` | — (sem par) |
| `porsche_model` | `PorscheModelSanitized` |
| `model_year` | `ModelYearSanitized` |
| `sale_price` | `SalesPriceSanitized` |
| `vehicle_mileage` | `VehicleMileageSanitized` |
| `payment_method` | `PayMethodSanitized` |
| `city` | `CitySanitized` |
| `state` | `StateSanitized` |
| `salesperson` | — (sem par) |
| `delivery_status` | `DeliveryStatusSanitized` |

100 linhas de dados, `sale_id` de 6 a 105.

## Etapas executadas pelo agente

1. **Leitura da estrutura** (`A1:U3`) para mapear cabeçalhos, estilos e exemplos de valores brutos vs. sanitizados.
2. **Verificação de fórmulas** — usando `execute_office_js`, confirmou que nenhuma célula da faixa sanitizada continha fórmula; todos os valores já eram estáticos (resultado de um processo de sanitização anterior, fora desta sessão).
3. **Conversão de tipo em `VehicleMileageSanitized`** — a coluna guardava números como texto (ex.: `"9800"`). Foi convertida para número real e formatada como `#,##0`.
4. **Remoção das 9 colunas brutas** que tinham par sanitizado (`sale_date`, `porsche_model`, `model_year`, `sale_price`, `vehicle_mileage`, `payment_method`, `city`, `state`, `delivery_status`), via `table.columns.getItem(name).delete()`. Colunas sem par (`sale_id`, `customer_name`, `salesperson`) foram preservadas.
5. Resultado: tabela reduzida de 21 para **12 colunas tratadas**:
   `sale_id, SaleDateSanitized, customer_name, PorscheModelSanitized, ModelYearSanitized, SalesPriceSanitized, VehicleMileageSanitized, PayMethodSanitized, CitySanitized, StateSanitized, salesperson, DeliveryStatusSanitized`.
6. Dataset exportado em CSV (100 linhas) — essa é a mesma base embutida no array `DATA` do [`index.html`](../index.html).

## Conflito identificado — não resolvido nesta sessão

Depois da limpeza (passos 1–6), o usuário enviou o [`SchemaPorshe.md`](../SchemaPorshe.md) como
referência que "sempre deve ser seguida". A seção **Quality checks** do schema exige:

- "No original columns were removed."
- "Sanitized columns appear immediately after their source columns."

Isso contradiz diretamente a etapa 4 acima (colunas brutas foram removidas, não mantidas
lado a lado com as sanitizadas). O agente sinalizou o conflito e começou a auditar os
*valores* sanitizados linha a linha contra as regras do schema (datas, ano, preço, milhagem
com conversão KM→milhas, sigla de estado, etc.), mas a transcrição da sessão foi cortada
antes da auditoria terminar e antes de uma decisão ser tomada sobre o layout de colunas.

### Decisão — resolvida

O schema é seguido à risca. A remoção das colunas brutas (etapa 4) foi apenas um atalho
de conveniência para facilitar a identificação das colunas ao montar o dashboard — não é
o estado final do dataset canônico.

O usuário forneceu o arquivo bruto original (`Tratamento Dados Porshe`, planilha
`db_psc_25354543_no_tracked`, 100 linhas, `sale_id` 6–105, 12 colunas brutas), versionado
em [`data/porsche_sales_raw.xlsx`](../data/porsche_sales_raw.xlsx). A partir dele,
[`scripts/sanitize.py`](../scripts/sanitize.py) reimplementa todas as regras do
[`SchemaPorshe.md`](../SchemaPorshe.md) (datas, modelo, ano, preço, milhagem com conversão
KM→milhas, forma de pagamento, cidade, estado, status de entrega) e gera o dataset
canônico de **22 colunas** — cada `*Sanitized` imediatamente após sua coluna bruta —
em [`data/porsche_sales_sanitized.csv`](../data/porsche_sales_sanitized.csv).

**Auditoria de valores:** os 100 registros recalculados a partir do bruto foram
comparados campo a campo com os valores que já estavam no `index.html` (produzidos na
sessão do Excel) — **0 divergências** em data, modelo, ano, preço, milhagem, pagamento,
cidade, estado e status. Isso confirma que os valores sanitizados originais já estavam
corretos; o único problema real era o layout de colunas (bruto removido), agora corrigido
na fonte versionada neste repositório.

`index.html` continua consumindo apenas o subconjunto sanitizado (12 campos) — isso é uma
projeção de exibição gerada por
[`scripts/build_dashboard_data.py`](../scripts/build_dashboard_data.py) a partir do CSV
canônico, não o dataset fonte, então não viola o schema.

Pipeline reprodutível:

```
data/porsche_sales_raw.xlsx
  → scripts/sanitize.py           → data/porsche_sales_sanitized.csv (22 colunas)
  → scripts/build_dashboard_data.py → index.html (const DATA, 12 campos)
```

## Observação sobre esta transcrição

A transcrição original (formato JSON exportado do add-in do Excel) foi fornecida
parcialmente — havia um limite de 50.000 caracteres na mensagem, então os registros de
auditoria a partir de aproximadamente `sale_id 49` em diante não foram capturados. Este
documento reflete apenas o que estava disponível até o corte.
