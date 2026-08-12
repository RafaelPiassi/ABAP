# ZFI051 — Contratos de Compra × Vendas (ZSD034)

Programa `ZFIR051`. Módulo FI / SD. Desenvolvedor: RPIASSI — Rafael Piassi.

Cruza os contratos de grãos do cockpit (`ZCOCKPITECC` / `ZFI_WSYS_CAB`) com o
relatório de carga da **ZSD034**, usando o CPF / CNPJ como chave entre o
**fornecedor** (compra) e o **cliente** (venda) — ou, no modo separado, entrega
os dois lados lado a lado sem cruzar nada.

O programa **não refaz** a lógica da ZSD034: ele submete o programa dela
(`ZSDR006_NEW`, constante `GC_PROG`) com os filtros do bloco de vendas e captura
a saída em memória pela `CL_SALV_BS_RUNTIME_INFO`. Os números são exatamente os
que a ZSD034 mostra, e coluna nova que ela ganhar aparece aqui sozinha — o bloco
de vendas é montado em tempo de execução.

---

## Modos de extração

O bloco **Modo de extração** decide a forma da saída.

### Cruzado (padrão — `P_MCRUZ`)

Uma saída só. Bloco do contrato e bloco da ZSD034 na **mesma linha**, ligados
pelo CPF / CNPJ:

| lado | de onde vem o documento |
|---|---|
| compra | `ZFI_WSYS_CAB-CODFORNECEDOR` → `LFA1-STCD1` / `STCD2` |
| venda | (ZSD034) `KUNNR2` / `KUNNR` → `KNA1-STCD1` / `STCD2` |

A comparação usa só os **dígitos** e separa CNPJ de CPF (`NORM_DOC`), então
máscara e zeros à esquerda não atrapalham.

É um cruzamento, não um resumo: fornecedor com 3 contratos e 10 linhas de venda
gera 30 linhas. A coluna **Cruzamento** marca cada linha como `Compra+Venda`,
`So compra` ou `So venda`.

### Separado (`P_MSEP`) — sem cruzar compra com venda

**Dois ALVs distintos**, empilhados num splitter sobre docking:

- **de cima** — contratos de compra (as colunas do cockpit + CNPJ, CPF e IE);
- **de baixo** — a saída da ZSD034, exatamente como ela devolve.

Cada lado obedece só aos **seus** filtros: o de compra ao bloco *Contratos*, o
de venda ao bloco *Vendas*. Nenhuma linha é multiplicada e nenhuma é descartada
por falta de par — os dois totais fecham com os relatórios de origem. Cada grid
tem sua barra de funções, seu layout e seu export próprios.

Neste modo:

- o bloco **Regra do cruzamento** fica cinza (não há cruzamento a regular);
- a coluna **Cruzamento** some da saída;
- a ZSD034 é submetida **sem** a restrição de clientes, ou seja, traz o período
  inteiro do bloco de vendas.

Para conciliar os dois lados manualmente, a chave é o **CNPJ / CPF** — ele está
nas duas saídas (no ALV de compra como coluna própria; no de venda, via
`Emissor da Ordem` / `Recebedor mercadoria`).

---

## Arquivo da extração

No modo separado, o bloco **Modo de extração** também escolhe a forma do arquivo:

| opção | resultado |
|---|---|
| `P_X1ARQ` (padrão) | **um** `.xlsx` com **duas abas**: `Compras (Contratos)` e `Vendas (ZSD034)` |
| `P_X2ARQ` | **dois** `.xlsx`: o nome informado ganha os sufixos `_COMPRAS` e `_VENDAS` antes da extensão |

No modo cruzado nada muda: continua um arquivo único, gerado pelo
`CL_SALV_BS_LEX`.

O arquivo de duas abas é montado no próprio programa (OOXML + `CL_ABAP_ZIP`),
**sem GUI e sem ABAP2XLSX** — condição para o Job funcionar, já que quem grava é
o servidor de aplicação. O `CL_SALV_BS_LEX` só gera arquivo de uma aba e OLE /
DOI exigiria GUI, daí a montagem na unha. O que sai do gerador:

- linha 1 com os títulos em negrito, congelada e com autofiltro;
- data como data de verdade (número de série do Excel, formato curto), então
  ordena e filtra como data;
- número como número, com ponto decimal e sinal à esquerda;
- texto passando pelo `WRITE ... TO`, para respeitar o exit de conversão do
  domínio — material, fornecedor e cliente saem como o usuário vê, não com os
  zeros à esquerda internos;
- zero é escrito como zero, não como célula vazia (branco se leria como "não
  informado").

Os títulos das colunas são lidos de uma ALV headless, então são **os mesmos**
que aparecem na tela — inclusive para colunas novas que a ZSD034 vier a ganhar.

### Destino: servidor ou PC

`P_DSRV` (padrão) grava no **servidor de aplicação**, por `OPEN DATASET`.
`P_DPC` grava na **estação de trabalho**, por `GUI_DOWNLOAD`.

> **É aqui que quase todo mundo tropeça.** O Job roda no servidor de aplicação,
> que não enxerga o disco do seu PC. Se você apontar `C:\Users\...\Desktop` com
> destino "servidor", o job termina com status **Concl.** e **não gera arquivo
> nenhum** — a falha fica no *spool* do job, não no log. Para o Job, use um
> caminho dos que a **AL11** lista. O agendamento agora barra caminho de PC na
> hora de agendar, e o F4 do diretório só navega no PC quando o destino é
> "estação de trabalho".

Em background o destino é sempre o servidor: gravar no PC exige SAPGUI.

### Quando o arquivo é gerado

- **Sempre** em background — vale tanto o Job agendado pelo botão quanto uma
  execução solta por F9 / SM36, porque o programa detecta `SY-BATCH`.
- **Em primeiro plano**, só se `P_GERAR` estiver marcado. Aí o programa gera o
  arquivo **e** exibe os ALVs, e informa num popup o que gravou (ou por que não
  gravou).

### Onde aparecem as mensagens

Em background as mensagens vão para a **lista do job**, que fica no **spool** —
não no log do job. Em primeiro plano elas saem juntas num popup. Se o job
terminar sem arquivo, o motivo está no spool: SM37 → selecione o job → botão
**Spool**.

---

## Job de extração automática

Bloco **Job**: diretório (com F4), nome do arquivo, data/hora da primeira
execução e a **periodicidade definida pelo usuário** — a cada N minutos, horas,
dias, semanas ou meses; `N = 0` agenda uma execução única.

| botão | o que faz |
|---|---|
| Agendar Job de extração | cria o agendamento recorrente com a seleção da tela |
| Cancelar agendamento | remove os agendamentos futuros (não mexe no que já rodou) |
| Situação do agendamento | mostra quantos agendamentos há e a próxima execução |

O modo de extração e a forma do arquivo viajam na seleção: o Job gera exatamente
o que a tela está pedindo.

No nome do arquivo valem os curingas `&DATA&` (AAAAMMDD) e `&HORA&` (HHMMSS).
Sem curinga, cada execução **sobrescreve** o mesmo arquivo.

> O F4 do diretório navega no PC do usuário, mas quem grava é o **servidor de
> aplicação** (`OPEN DATASET`). O caminho precisa existir e ser gravável por ele
> — confira na **AL11**.

---

## Tela de seleção

| Bloco | Campos |
|---|---|
| **Contratos** | `P_BUKRS` (obrigatório), contrato, data do contrato, data de pagamento, safra, fornecedor, centro, CNPJ, CPF, IE, status |
| **Vendas** | data de criação da OV (obrigatório), documento, org. vendas, canal, setor, escritório, equipe, tipo de OV, ano-safra, centro, material, grupo de mercadorias |
| **Modo de extração** | cruzado × separado; um arquivo de duas abas × dois arquivos |
| **Regra do cruzamento** | casa por Emissor da ordem (AG) / Recebedor (WE) / qualquer um; trazer contrato sem venda; trazer venda sem contrato |
| **Arquivo / Job** | destino (servidor × PC), gerar nesta execução, diretório, nome do arquivo, nome do job, data/hora inicial, periodicidade |

---

## Tabelas e objetos envolvidos

`ZFI_WSYS_CAB` (contratos do cockpit) · `LFA1` · `KNA1` · `T001W` · `TBTCO` ·
`VBAK` / `VBAP` / `MARA` (só como referência de tipo dos SELECT-OPTIONS).

Programa submetido: **`ZSDR006_NEW`** (o programa por trás da ZSD034).

Classes: `CL_SALV_TABLE`, `CL_SALV_BS_RUNTIME_INFO`, `CL_SALV_BS_LEX`,
`CL_SALV_EX_UTIL`, `CL_SALV_CONTROLLER_METADATA`, `CL_GUI_DOCKING_CONTAINER`,
`CL_GUI_SPLITTER_CONTAINER`, `CL_ABAP_ZIP`, `CL_ABAP_CONV_OUT_CE`, RTTI
(`CL_ABAP_STRUCTDESCR` / `CL_ABAP_TABLEDESCR`).

---

## Instalação

1. **SE38** — criar o programa `ZFIR051` e colar o fonte de `ZFIR051.abap`.
2. **SE38 → Ir para → Elementos de texto → Símbolos de texto**, cadastrar:

   | Símbolo | Texto |
   |---|---|
   | `B01` | Contratos |
   | `B02` | Vendas (ZSD034) |
   | `B03` | Regra do cruzamento |
   | `B04` | Arquivo e Job de extração automática |
   | `B05` | Modo de extração |

   E os textos de seleção dos parâmetros novos:

   | Nome | Texto |
   |---|---|
   | `P_MCRUZ` | Cruzar compra x venda |
   | `P_MSEP` | Dois ALVs, sem cruzar |
   | `P_X1ARQ` | Arquivo único, duas abas |
   | `P_X2ARQ` | Dois arquivos separados |
   | `P_DSRV` | Destino: servidor (AL11) |
   | `P_DPC` | Destino: estação de trabalho |
   | `P_GERAR` | Gerar arquivo nesta execução |

3. **SE93** — criar a transação `ZFI051` apontando para o programa `ZFIR051`.

### Premissas do ambiente (confirmar antes de usar em produção)

- A transação **ZSD034** aponta para o programa `ZSDR006_NEW`. Se apontar para
  outro (veja na SE93), troque **só** a constante `GC_PROG`.
- A ZSD034 tem os SELECT-OPTIONS `S_AUDAT`, `S_VBELN`, `S_VKORG`, `S_VTWEG`,
  `S_SPART`, `S_VKBUR`, `S_VKGRP`, `S_AUART`, `S_ANOSF`, `S_WERKS`, `S_MATNR`,
  `S_MATKL` e `S_KUNNR` — são esses os nomes usados no `SUBMIT`.
- A saída da ZSD034 traz as colunas `KUNNR` (recebedor) e `KUNNR2` (emissor da
  ordem), que são o elo com o contrato no modo cruzado.
- A tabela `ZFI_WSYS_CAB` tem os campos `CODIGOEMPRESA`, `CENTRO`,
  `NR_CONTRATO`, `CODFORNECEDOR`, `DT_CONTRATO`, `DATAPAGAMENTO`, `ANOSAFRA`,
  `QUANTIDADE`, `VALOR_CONTRATO` e `STATUS`.
- O diretório do Job existe no servidor de aplicação e é gravável por ele
  (AL11).

Colunas da ZSD034 que não estiverem na lista de títulos (`CARREGAR_LABELS`) saem
com o nome técnico do campo — e continuam aparecendo no relatório.
