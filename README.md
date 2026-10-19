# Gerador de Proposta Comercial (100% local no navegador)

Transforma um **timeline** (Excel) em uma **proposta comercial**: análise financeira, documento em PT/EN/ES,
PDF, Word e um arquivo Excel que guarda tudo da proposta. Nada é enviado a servidores — os dados ficam no
seu computador.

## Como usar
1. Abra `gerador_proposta_comercial.html` no navegador (duplo clique; funciona **sem internet**).
2. **Timeline** (aba 1): defina os **parâmetros do timeline** (sprint ou semana, nº de períodos, semanas por sprint,
   horas/dia, dias úteis e início por data, mês ou sem data) e **monte a timeline na própria tela**: fases, perfil e
   **nível por combo** (dá para cadastrar novos), FTE digitado ou **pintado com arrastar**, com desfazer/refazer e
   prévia da proposta. Ou abra um Excel existente.
3. **Parâmetros & Perfis** (aba 2): calendário, premissas de preço, **desconto**, moeda/câmbio, impostos, faturamento
   e as taxas de cada perfil/nível.
4. **Análise**: alertas (taxas faltando, GM% baixo…), KPIs, gráficos, sensibilidade desconto × contingência.
5. **Proposta**: idioma e opções de timeline (barras, fases + perfis ou só marcos; por sprint ou semana; com/sem FTE),
   edição dos textos na página e exportação em **PDF** (imprimir), **Word (.docx)**, **PowerPoint (.pptx)** ou **Excel**.

A **barra superior** mostra sempre qual proposta está aberta: cliente — projeto, referência, arquivo, preço total,
margem (GM) e o status de gravação.

## Perfis e níveis
Os perfis são cadastrados com **nível**, local e taxas (Custo/h e Rate/h). Na timeline, cada linha escolhe o **perfil**
por uma combo e o **nível** por outra; níveis novos podem ser adicionados na hora e perfis novos entram no cadastro
(aba `Perfis` do Excel). Um perfil só exige taxa se estiver realmente alocado.

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
parâmetros, timeline, perfis/níveis e taxas, impostos, descontos e os textos editados (abas `Parametros`, `Modelo`,
`Perfis`, `Impostos`, `Descontos`, `Textos`), além de abas de auditoria com **fórmulas** (`Resumo`, `PorFase`).

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
npm test           # testes do motor de cálculo e do arquivo Excel (compara fórmulas com o LibreOffice, se instalado)
npm run e2e        # teste no Chromium com a rede bloqueada
```
Estrutura de `src/`: `engine.js` (cálculo puro), `doc.js` (modelo do documento), `workbook.js` (arquivo Excel),
`docx-export.js` (Word), `pptx-export.js` (PowerPoint), `net.js` (extras online), `storage.js` (IndexedDB e arquivos), `app.js` (interface).
**Edite `src/` e rode `npm run build`** — o HTML na raiz é gerado.
