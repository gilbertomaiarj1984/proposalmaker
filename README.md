# Gerador de Proposta Comercial (100% local no navegador)

Transforma um **timeline** (Excel) em uma **proposta comercial**: análise financeira, documento em PT/EN/ES,
PDF, Word e um arquivo Excel que guarda tudo da proposta. Nada é enviado a servidores — os dados ficam no
seu computador.

## Como usar
1. Abra `gerador_proposta_comercial.html` no navegador (duplo clique; funciona **sem internet**).
2. **Timeline** (aba 1): defina os **parâmetros do timeline** (sprint ou semana, nº de períodos, semanas por sprint,
   horas/dia, dias úteis e início por data, mês ou sem data) e **monte a timeline na própria tela**: fases, perfil e
   **nível por combo** (dá para cadastrar novos), FTE digitado ou **pintado com arrastar**, com desfazer/refazer e
   prévia da proposta. Ou abra um Excel existente. Em cada fase dá para inserir **entregáveis** (◆) e dá para criar até
   **3 cenários** de timeline/alocação.
3. **Parâmetros & Perfis** (aba 2): dados da proposta (**cliente**), calendário, premissas de preço, **desconto**,
   moeda/câmbio, impostos, faturamento e, no fim da página, **Perfis e taxas** do cliente selecionado.
4. **Cenários**: painel comparativo dos cenários (prazo, esforço, equipe, preço, margem…).
5. **Análise**: alertas (taxas faltando, GM% baixo…), KPIs, gráficos, sensibilidade desconto × contingência.
6. **Proposta**: idioma e opções de timeline (barras, fases + perfis ou só marcos; por sprint ou semana; com/sem FTE),
   edição dos textos na página e exportação em **PDF** (imprimir), **Word (.docx)**, **PowerPoint (.pptx)** ou **Excel**.

A **barra superior** mostra sempre qual proposta está aberta: cliente — projeto, referência, arquivo, preço total,
margem (GM) e o status de gravação.

## Catálogo de perfis e taxas (segundo arquivo Excel)
Perfis e taxas **não ficam mais no arquivo da proposta**: vêm de um **catálogo** (`Perfis_e_Taxas.xlsx`) com **uma aba por
cliente** (inicialmente `Roche` e `JnJ`), todas no **layout padrão** `Role | Level | Location | Cost/h | Rate net/h`
(JnJ também; Local e Custo/h podem ficar em branco). Abas com outros títulos de coluna continuam sendo lidas pelo
cabeçalho, e clientes novos em branco já nascem nesse padrão.
* **📚 Catálogo** (topo): abre o arquivo (Chrome/Edge lembram dele e gravam direto; outros navegadores baixam o `.xlsx`).
* Em *Dados da proposta* (aba 2, ou no seletor da aba 1) **escolha o cliente** ou crie um novo com **＋ Novo**: o app
  **pergunta em qual cliente basear** (copia perfis, níveis e formatação) ou cria uma aba em branco; cada cliente novo
  gera uma **aba nova** no catálogo.
* No fim da aba 2, **Perfis e taxas** mostra só os dados do cliente selecionado e permite editar/adicionar/excluir; as
  edições são gravadas na aba do cliente, **preservando a formatação** do arquivo original.
* A proposta guarda apenas uma **cópia das taxas usadas** (colunas `Cost/h (catálogo)` e `Rate/h (catálogo)` na aba
  `Modelo`) para recalcular sem o catálogo. Ao reabrir com o catálogo carregado, o **catálogo prevalece** e o app avisa
  se as taxas mudaram, com o impacto no preço (antes → depois).
* `npm run catalog` regenera o arquivo de exemplo (Roche é ilustrativo; JnJ traz as linhas da amostra, com Rate/h em BRL sem impostos).

Na timeline, cada linha escolhe o **perfil** por uma combo e o **nível** por outra; níveis novos podem ser adicionados na
hora. Um perfil só exige taxa se estiver realmente alocado.

## Margens: GM e GPM (painel na Análise)
* **GM** = (receita líquida − custo direto) ÷ receita líquida. **GPM** = GM depois dos **custos indiretos**: overhead/G&A e
  comissão (% da receita líquida), custo financeiro (% a.m. × prazo de pagamento ÷ 30, sobre o valor com impostos) e
  reserva de risco (% do custo direto). Tudo é configurado em *Margens e custos indiretos* (aba 2) e é **interno**: não
  aparece no documento, Word ou PowerPoint do cliente.
* O painel (aba Análise) mostra GM/GPM contra os mínimos, a **cascata** do preço ao GPM, a margem **por fase** e **por perfil**,
  a **calculadora reversa** (informe o GM ou GPM alvo e veja o preço, o ajuste de rates ou a redução de custo necessários, com
  “aplicar como preço alvo”) e o **simulador** desconto × margens. O desconto máximo sugerido passa a respeitar também o
  GPM mínimo; os cenários, a barra superior, os alertas e o Excel (`Resumo`, `PorFase`, com fórmulas) trazem o GPM.
* **Salário médio sugerido** (Perfis e taxas): Cost/h × horas por mês ÷ (1 + encargos e benefícios CLT, padrão 70,44%,
  ajustável). **Racional ao passar o mouse**: preço, GM, GPM, KPIs, descontos, cenários, perfis e o salário mostram a conta.
* A biblioteca de propostas oculta arquivos que não são propostas (ex.: o catálogo) e o catálogo recusa abrir uma proposta.

## Entregáveis por fase
Em cada fase, o botão **◆** adiciona entregáveis (ficam abaixo das linhas da fase). Por padrão são **distribuídos
igualmente** no tempo alocado da fase (e acompanham a fase se a alocação mudar); ajuste **início/fim** nas caixas
(ou clique numa célula: define o fim; Shift+clique: define o início) e o entregável vira **manual** — sempre dentro da
janela da fase. **⇄ distribuir igualmente** volta tudo ao automático. Aparecem como ◆ na timeline (tela, Word e
PowerPoint), numa tabela *Entregáveis* e nos marcos (opção *Mostrar entregáveis* na aba Proposta).

## Cenários e comparativo
Cada proposta aceita **até 3 cenários** (barra acima do editor: ＋ Novo, ⧉ Duplicar, ✎ Renomear, 🗑 Excluir). Timeline,
alocação, entregáveis e parâmetros do timeline (unidade, semanas por sprint, início…) são **por cenário**; preços,
impostos, desconto e textos são comuns. A aba **3. Cenários** compara lado a lado: cartões com a forma da equipe, tabela
com o melhor valor em verde e a diferença em relação ao cenário ativo (prazo, datas, esforço, pico/FTE médio, fases,
entregáveis, preço de lista, desconto, preço final, GM%, rate médio, alertas), gráficos de FTE por período e de preço ×
margem, e uma leitura automática das diferenças. **Usar este cenário** o torna ativo: análise, proposta e exportações
sempre usam o cenário ativo (o nome aparece na barra superior). No Excel: abas `Cenarios` (lista e parâmetros),
`Cen_n` (matriz dos cenários inativos) e `Entregaveis`; `Modelo` é sempre o cenário ativo, com as fórmulas.

## Desconto
Ordem de cálculo: preço de lista (taxa × horas) → contingência → **desconto** → impostos (gross-up).
* **Tipo:** percentual sobre o total (com exceções **por fase**), **valor fixo** ou **preço final alvo** — os dois
  últimos em valores **com impostos** (o que o cliente vê); o app calcula o percentual equivalente.
* **Exibição:** embutido nas taxas ou **linha visível** (preço de lista → desconto → preço final).
* **Sugestões:** desconto máximo que preserva o GM% mínimo, escada de negociação com preço e margem de cada degrau,
  e dicas de como aplicar.

## PowerPoint editável
Um slide por seção da proposta (capa, sumário executivo, cronograma, time, investimento, pagamentos, opcionais e
premissas). Tudo é nativo do PowerPoint — texto, tabelas, formas e gráfico, sem imagens — e respeita idioma, textos
editados e as opções de timeline. **A timeline sempre ocupa um único slide**: com muitas linhas, a altura das linhas e a fonte
reduzem até caber (e a tabela de fases/datas entra no mesmo slide quando couber). Demais tabelas longas continuam em outro
slide, com o cabeçalho repetido.

## O Excel como "banco" de cada proposta
Cada proposta é **um arquivo `.xlsx`**. A edição acontece no navegador; o Excel guarda o resultado:
parâmetros, timeline, cópia das taxas usadas, impostos, descontos, entregáveis, cenários e os textos editados (abas
`Parametros`, `Modelo`, `Impostos`, `Descontos`, `Entregaveis`, `Cenarios`/`Cen_n`, `Textos`), além de abas de auditoria com **fórmulas** (`Resumo`, `PorFase`).

* **Salvar como…** / **Salvar**: no Chrome/Edge grava direto no arquivo e, depois disso, **cada alteração é
  gravada automaticamente** nele. Em outros navegadores o `.xlsx` é baixado.
* **Abrir do Excel** (ou arrastar o arquivo na aba 1) restaura tudo. Um timeline “bruto” (template) **nunca é
  sobrescrito**: o app pede “Salvar como…” para criar a proposta e só então passa a gravar nela.
* **Propostas** (Chrome/Edge): conecta uma pasta do computador (pode ser OneDrive/rede), lista as propostas,
  abre, salva e **compara** várias.
* Um rascunho também é guardado no navegador (IndexedDB) e oferecido para restauração ao reabrir.
* Limite: não há controle de edição simultânea; é um arquivo por proposta.

## Máquinas com internet restrita
Todas as bibliotecas vão **embutidas** no HTML. A internet só é usada em dois extras opcionais e ambos
falham com aviso claro, sem travar o app:
* **Câmbio** (Frankfurter/BCE, com fallback open.er-api.com) — alternativa: digitar o câmbio ou usar a última cotação salva.
* **Conferir feriados** (BrasilAPI) — os feriados nacionais já são **calculados localmente**; a consulta só confere/acrescenta.

O botão *Conexão* no topo testa o acesso e a faixa amarela avisa quando algo está indisponível.

## Desenvolvimento
```sh
npm install
npm run build      # gera gerador_proposta_comercial.html (arquivo único) a partir de src/
npm run template   # regenera Timeline_Template.xlsx (mesma função do botão do app)
npm run catalog    # regenera Perfis_e_Taxas.xlsx (catálogo de exemplo, uma aba por cliente)
npm test           # testes do motor de cálculo e do arquivo Excel (compara fórmulas com o LibreOffice, se instalado)
npm run e2e        # teste no Chromium com a rede bloqueada
```
Estrutura de `src/`: `engine.js` (cálculo puro), `doc.js` (modelo do documento), `workbook.js` (arquivo Excel),
`catalog.js` (catálogo de perfis e taxas), `editor.js` (editor de timeline), `docx-export.js` (Word), `pptx-export.js` (PowerPoint), `net.js` (extras online), `storage.js` (IndexedDB e arquivos), `app.js` (interface).
**Edite `src/` e rode `npm run build`** — o HTML na raiz é gerado.
