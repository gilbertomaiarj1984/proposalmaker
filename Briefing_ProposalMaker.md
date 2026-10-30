# ProposalMaker — briefing para o "capitão" do projeto

> Cole este documento no início de uma conversa/projeto no Claude. Idioma de trabalho: **português**.

## 1. O que é
Gerador de **propostas comerciais 100% local no navegador**. Parte de um **timeline** (Excel ou montado na própria tela),
calcula a análise financeira e produz o documento da proposta em **PT/EN/ES**, exportável em **PDF, Word, PowerPoint
(editável) e Excel**. Nada vai para servidores; os dados ficam no computador do usuário.

- **Entregável único:** `gerador_proposta_comercial.html` (arquivo gerado, ~3 MB, abre com duplo clique, **funciona offline**; bibliotecas embutidas).
- **Repositório:** `gilbertomaiarj1984/proposalmaker` (branch `main`; trabalho em branch própria → PR → squash).
- **Internet é opcional:** só para câmbio (Frankfurter/BCE) e conferência de feriados (BrasilAPI). Se falhar, **avisa e nunca trava**.
- **Sem IA** no produto (foi pedido e depois removido de propósito).

## 2. Como o usuário usa (5 abas)
1. **Timeline** — parâmetros (sprint/semana, nº de períodos, horas/dia, início por data/mês/sem data), editor em grade
   (fase → perfil por combo → nível por combo → FTE digitado ou pintado), **entregáveis (◆) por fase**, **até 3 cenários**,
   desfazer/refazer, prévia ao vivo da proposta.
2. **Parâmetros & Perfis** — dados da proposta (cliente), calendário, premissas de preço, **Margens e custos indiretos**,
   **desconto** (%, valor fixo ou preço alvo; por fase; embutido ou linha visível), moeda/câmbio, impostos, faturamento,
   e no fim **Perfis e taxas** do cliente selecionado.
3. **Cenários** — painel comparativo (cartões, tabela com melhor valor em verde e diferenças, gráficos, leitura automática).
4. **Análise financeira** — alertas, KPIs, **painel GM/GPM**, gráficos, tabelas por fase/perfil/mês/pagamentos, sensibilidade.
5. **Proposta** — idioma, opções de timeline (barras, fases+perfis, só marcos; por sprint/semana; com/sem FTE), edição dos
   textos na página, exportações.

A **barra superior** identifica sempre a proposta aberta: cliente — projeto, ref., cenário ativo, arquivo, preço, GM, GPM, status de gravação.

## 3. Conceitos de negócio (as regras que mandam)
- **Ordem do preço:** preço de lista (taxa × horas) → contingência → **desconto** → **impostos (gross-up)**. Desconto fixo/alvo é informado **com impostos** (o que o cliente vê).
- **GM** = (receita líquida − custo direto) ÷ receita líquida.
- **GPM** = GM depois dos **custos indiretos**: overhead/G&A e comissão (% da receita líquida), custo financeiro
  (% a.m. × prazo de pagamento ÷ 30 × valor com impostos) e reserva de risco (% do custo direto). **Margens são internas: nunca aparecem no documento/Word/PPT do cliente.**
- **Desconto máximo** sugerido protege o **GM mínimo e o GPM mínimo** (usa o mais restritivo).
- **Calculadora reversa:** dado GM/GPM alvo, mostra o preço, o ajuste de rates ou a redução de custo necessários.
- **Salário médio sugerido** (em Perfis e taxas) = Cost/h × horas/mês ÷ (1 + encargos e benefícios CLT). Padrão **70,44%**
  (INSS 20 + RAT 2 + Sistema S 5,8 + FGTS 8 + 13º 8,33 + férias 11,11 + multa FGTS 3,2 + benefícios 12), editável. É referência, não folha.
- **Racional ao passar o mouse**: toda conta relevante mostra o cálculo em tooltip (`data-why`). **Regra permanente: novos números calculados devem ganhar tooltip com o racional.**
- Perfil só exige taxa se estiver **realmente alocado**. Fase sem alocação não entra na proposta.
- **Entregáveis:** por padrão distribuídos igualmente no tempo da fase (auto); ajuste manual vira "manual", sempre dentro da janela da fase.
- **Cenários:** timeline, alocação, entregáveis e parâmetros do timeline são **por cenário**; preços, impostos, desconto e textos são **comuns**. Análise, proposta e exportações usam o **cenário ativo**.
- A timeline do PowerPoint **sempre cabe em um único slide**.

## 4. Os dois arquivos Excel (modelo de dados)
**a) Proposta (um `.xlsx` por proposta) — o "banco" da proposta, edição em 2 vias.** Edita-se no navegador e persiste no Excel;
reabrir restaura tudo. Abas: `Parametros`, `Impostos`, `Descontos`, `Modelo` (cenário ativo, com **fórmulas** e cópia das taxas
usadas: `Cost/h (catálogo)`/`Rate/h (catálogo)`), `Entregaveis`, `Cenarios` + `Cen_n` (só quando há >1 cenário), `Resumo` e
`PorFase` (fórmulas, inclui GM/GPM), `PorPerfil`, `Mensal`, `Pagamentos`, `Feriados`, `Textos`, `_meta` (oculta).
Um timeline "bruto" (template) **nunca é sobrescrito** — o app pede "Salvar como…".

**b) Catálogo de perfis e taxas (`Perfis_e_Taxas.xlsx`) — separado da proposta.** **Uma aba por cliente** (hoje Roche e JnJ),
**layout padrão para todos: `Role | Level | Location | Cost/h | Rate net/h`** (JnJ: Local e Cost/h podem ficar em branco).
Cliente novo = nova aba, e o app **pergunta em qual cliente basear**. Edições são gravadas na aba do cliente **preservando a
formatação**. Ao reabrir uma proposta com o catálogo carregado, **o catálogo prevalece** e o app avisa se as taxas mudaram (com o impacto no preço).
Proposta e catálogo **não se misturam**: a biblioteca de propostas oculta o catálogo; o catálogo recusa abrir uma proposta e vice-versa.

Persistência: Chrome/Edge gravam direto no arquivo (File System Access) e fazem autosave; outros navegadores baixam o `.xlsx`.
Há também rascunho no IndexedDB e uma **biblioteca de propostas** (pasta) com comparação entre propostas.

## 5. Arquitetura técnica
Código em `src/`; `node build.mjs` concatena tudo num único HTML (bibliotecas embutidas e carregadas sob demanda; `docx` via esbuild).
**Edite `src/` e rode `npm run build` — o HTML na raiz é gerado e vai no commit.**

| Arquivo | Papel |
|---|---|
| `engine.js` | Cálculo **puro**: parse do timeline, `compute` (preço, desconto, GM/GPM, fases, pagamentos, alertas), calendário/feriados BR calculados localmente, `timelineView`, marcos, entregáveis, cenários (`scenarioMetrics`, `compareScenarios`), `solveMargin`, `salaryFromCost` |
| `doc.js` | Modelo do documento em blocos (i18n PT/EN/ES) → HTML |
| `docx-export.js` / `pptx-export.js` | Word e PowerPoint nativos/editáveis |
| `workbook.js` | Leitura/gravação do Excel da proposta (ExcelJS grava, SheetJS lê) + `PARAM_DEFS` |
| `catalog.js` | Catálogo: detecta colunas pelo cabeçalho, edita a planilha original, clona aba de cliente |
| `editor.js` | Editor de timeline (grade, combos, pintar, undo/redo, entregáveis) |
| `app.js` | Interface/estado (`S`), abas, cenários, catálogo, painel de margens, tooltips, exportações |
| `net.js`, `storage.js` | Extras online com aviso; IndexedDB e arquivos |
| `styles.css`, `index.template.html` | Visual e estrutura |

Estado: `S.rows/nPeriods/unit/deliverables` = **cópia de trabalho do cenário ativo**; `S.scen[]` guarda os 3 cenários
(`commitScen()` grava, `loadScen()` troca). Taxas efetivas = catálogo do cliente + perfis fora do catálogo (cópia gravada na proposta).
Ganchos de teste: `window.__pm`.

## 6. Como trabalhar (comandos e regras)
```sh
npm install
npm run build      # gera o HTML único
npm run template   # Timeline_Template.xlsx
npm run catalog    # Perfis_e_Taxas.xlsx (exemplo)
npm test           # 61 testes: motor, Excel (recalcula fórmulas no LibreOffice), pptx, catálogo
npm run e2e        # 20 verificações no Chromium com a rede bloqueada
```
- **Fluxo:** branch `claude/...` → commit → push → PR (rascunho) → o usuário diz **"mesclar"** → squash merge. Sempre responder em português.
- Funcionalidades novas: o usuário gosta de **sugestões + preview (mock-up) antes de implementar**, e de ver verificação visual.
- Antes de mesclar: `npm test` e `npm run e2e` verdes; atualizar o README; regenerar o HTML.
- Toda função pura nova ganha teste; mudança de UI ganha passo no e2e.
- Limites conhecidos: sem edição simultânea (um arquivo por proposta); gravação direta só em Chrome/Edge.

## 7. Histórico (resumo)
#1 reescrita com Excel como arquivo da proposta, modo offline e correções → #2 PowerPoint editável → #3 editor de timeline no
navegador, desconto, Excel em 2 vias, barra de identificação, "Perfis" → #4 catálogo de perfis por cliente + entregáveis por fase +
cenários com comparativo → #5 correção de cor dos entregáveis + layout padrão do catálogo → #6 painel GM/GPM, salário sugerido,
`Rate net/h`, racional ao passar o mouse, separação proposta × catálogo.

## 8. Pontos em aberto / ideias (não implementados)
- Tooltip de racional ainda ausente em: sensibilidade desconto × contingência, comparador da biblioteca, totais das tabelas de fase/pagamento.
- Margem-alvo **por fase** com alerta; simulação de mix de níveis (sênior/pleno/júnior); "margem aprovada" congelada vs. atual.
- Definição de GPM e dos indiretos foi **assumida** (ver §3) — confirmar com a área financeira e ajustar `indirects()` em `engine.js` se divergir.
- Os dados de Roche no catálogo são **ilustrativos**; JnJ traz linhas de amostra — substituir pelos oficiais.
- Percentual de encargos CLT (70,44%) é referência — validar com o RH.
