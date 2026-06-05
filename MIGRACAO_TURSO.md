# Migração de dados para o Turso

O arquivo **`seed_dados_manuais.sql`** contém os dados manuais do sistema
(estoque, custos/CMP, compras, atacado, personalizações). Os pedidos da
NuvemShop **não** estão incluídos — eles são re-sincronizados automaticamente
pelo app quando conectado à loja.

## Como importar no Turso

Pré-requisito: ter a [Turso CLI](https://docs.turso.tech/cli/introduction)
instalada e autenticada (`turso auth login`).

```bash
# Substitua <nome-do-banco> pelo nome do seu banco no Turso
turso db shell <nome-do-banco> < seed_dados_manuais.sql
```

As inserções usam `INSERT OR IGNORE`, então é seguro rodar mais de uma vez —
registros já existentes não são duplicados.

## Tabelas incluídas

| Tabela                | Conteúdo                                  |
|-----------------------|-------------------------------------------|
| `sku_stock`           | Níveis de estoque por variante            |
| `sku_costs`           | Custos e Custo Médio Ponderado (CMP)      |
| `sku_pers_pricing`    | Preços de personalização                  |
| `sku_stock_movements` | Histórico de movimentações de estoque     |
| `compras_registros`   | Compras registradas                       |
| `compras_tamanhos`    | Quantidade por tamanho de cada compra     |
| `compras_manual`      | Itens da lista de compras                 |
| `atacado_pedidos`     | Pedidos de atacado                        |
| `atacado_itens`       | Itens dos pedidos de atacado              |
| `personalizacoes`     | Personalizações de pedidos                |
| `romaneios`           | Romaneios                                 |
| `config`              | Configurações do sistema                  |

## Variáveis de ambiente necessárias (Vercel)

O app lê o banco Turso destas variáveis (configurar no painel da Vercel):

```
TURSO_DATABASE_URL = libsql://<seu-banco>.turso.io
TURSO_AUTH_TOKEN   = <seu-token>
```
