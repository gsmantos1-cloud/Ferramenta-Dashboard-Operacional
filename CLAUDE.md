# GS Mantos — Dashboard Operacional

Flask + Turso (libSQL via HTTP, `turso_db.py`) hospedado no Vercel (serverless).

---

## 🚨 PRIORIDADE Nº 1 — NÃO ESTOURAR A QUOTA GRÁTIS DO TURSO

**Esta é a regra mais importante do projeto. Antes de qualquer outra coisa.**

O banco roda no **plano grátis do Turso**: **500 milhões de leituras (rows read)/mês**.
- O limite é **por conta** e **reseta no ciclo mensal de billing**.
- Estourar **bloqueia as leituras da conta inteira** até o reset → **a ferramenta para de funcionar**.
- Isso já aconteceu uma vez (962M de leituras por um bug de sync) e exigiu migrar de conta.

### Regra de ouro
**Toda mudança deve MINIMIZAR ao máximo as leituras, sem comprometer as funções da ferramenta.**

Antes de finalizar qualquer alteração, pergunte‑se:
1. Isso adiciona **leituras recorrentes** (em polling, em cada page load, em cada request)?
2. As consultas novas têm **índice** na coluna filtrada?
3. Tem algum SELECT rodando **dentro de um loop** por item?

Se a resposta a 1 for "sim", reduza a frequência ao mínimo. Se 2 for "não", adicione o índice. Se 3 for "sim", refatore para pré-carregar em memória.

### Práticas obrigatórias

1. **Índice em tudo que filtra.** Toda coluna usada em `WHERE`, `JOIN` ou `GROUP BY` precisa de índice — criados em `init_db()` (lista `indices`). Sem índice, a query faz *full table scan* e lê a tabela inteira (milhares de leituras por chamada). Ao escrever uma consulta nova, adicione o índice correspondente.

2. **Nunca ler dentro de loop por item.** Pré‑carregue o necessário com **um** SELECT (ex: um `set`/`dict` em memória) e processe em memória. Para gravar muitos registros, use `conn.execute_batch(lista_de_(sql, params))` — uma requisição em vez de N.

3. **Sync da NuvemShop = sempre incremental e retomável.**
   - Filtro de data **em ISO 8601** (`updated_at_min=YYYY-MM-DD`), **NUNCA** timestamp unix (a API ignora e baixa o histórico inteiro — foi a causa do estouro).
   - Processa em blocos de ~6s com cursor `sync_estado` em `config`; o front faz loop até `done:true`. Não remova essa lógica.

4. **Recálculos caros no máximo 1x/dia.** `atualizar_status_db()` só recalcula uma vez por dia (guard `status_calc_data` em `config`), porque o status de atraso só muda na virada do dia. Não volte a rodá‑lo em todo request.

5. **Frontend econômico.** Auto‑refresh no mínimo necessário (hoje 5 min) e **pausado quando a aba está oculta** (Page Visibility). Não refaça uma leitura que já foi feita no mesmo carregamento (ex: `/api/notificacoes` é compartilhada entre cabeçalho e dashboard via `window.__notifPromise`).

6. **Pedidos sem duplicar.** Índice ÚNICO `uniq_pedidos_numero` em `pedidos(numero)` + todos os INSERT de pedido usam `INSERT OR IGNORE`. Mantenha assim.

### Por que cada item importa
O consumo normal da ferramenta hoje é de ~**0,1% do limite mensal** mesmo nos dias pesados. Essa folga só existe por causa das práticas acima. **Qualquer regressão** (uma query sem índice num endpoint de polling, um loop com SELECT por item, um sync que volta a baixar tudo) pode multiplicar as leituras por milhares e estourar a quota de novo.

---

## ⚡ PRIORIDADE Nº 2 — NÃO ESTOURAR O PLANO GRÁTIS DO VERCEL

O app roda como função **serverless** no Vercel (plano Hobby). Aqui o que consome cota é **cada requisição** (invocação de função) + **tempo de compute**. Práticas obrigatórias:

1. **Nada de conexão aberta (SSE/WebSocket).** Serverless não segura conexão viva: a função reconecta em loop e queima invocações. Use **polling de carga única** (busca ao abrir a aba / botão "Atualizar"), **nunca** `EventSource`/WebSocket. (Já foi removido o SSE de `/api/stream/alertas` — os alertas de estoque usam `/api/alertas/poll`, 1 leitura ao abrir a aba.)
2. **Polling recorrente no mínimo e pausado com a aba oculta** (`document.hidden`). Referências atuais: Dashboard 5 min; Personalizações (colaboração em 2 PCs) 15 s. Não reduza esses intervalos sem necessidade real, e sempre com o guard de visibilidade.
3. **Cron: 1x/dia.** O Hobby permite ~2 crons e só rodam 1x/dia. Hoje há **1** (`/api/cron/snapshot-estoque`, `0 2 * * *` = 23h BRT, em `vercel.json`). Não adicione crons frequentes.
4. **Função rápida (< 10 s).** O sync da NuvemShop é fatiado em ~6 s e retomado pelo front. Não crie endpoints que demorem/segurem a função.
5. **Sem threads e sem estado em memória entre requests** (cada invocação é isolada) e **sem gravar em disco** (FS efêmero) — persista tudo no Turso.

Regra prática: antes de subir, pergunte "isso abre conexão viva, faz polling curto, ou roda demais?". Se sim, reduza.

---

## Ferramentas administrativas (pelo navegador, sem CLI)

- `/admin/migrar` → backup/restauração em JSON (`/api/admin/export`, `/api/admin/import`).
- `POST /api/admin/dedup` → remove pedidos/itens duplicados e garante o índice único.
- `POST /api/nuvemshop/registrar-webhooks` → registra webhooks de pedido pago/enviado/cancelado.

## Deploy
Push na `main` → o Vercel atualiza sozinho em 1‑2 min. Variáveis sensíveis (`TURSO_*`, `NUVEMSHOP_*`, `SECRET_KEY`, `LOGIN_*`) ficam só no Vercel.
