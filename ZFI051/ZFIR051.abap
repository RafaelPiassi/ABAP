*&---------------------------------------------------------------------*
*& Report  ZFIR051
*&---------------------------------------------------------------------*
*& Transacao....: ZFI051 - Contratos de Compra x Vendas (ZSD034)
*& Descricao....: Cruza os contratos do cockpit (ZCOCKPITECC) com o
*&                relatorio de carga da ZSD034, usando o CPF / CNPJ
*&                como chave entre o FORNECEDOR (compra) e o CLIENTE
*&                (venda).
*& Modulo.......: FI / SD
*& Desenvolvedor: RPIASSI - Rafael Piassi
*&---------------------------------------------------------------------*
*& Objetivo:
*&   Uma unica linha mostrando, lado a lado, o que a empresa COMPRA de
*&   um produtor (contrato de graos, cockpit ZCOCKPITECC) e o que VENDE
*&   para esse mesmo produtor (ZSD034). O elo entre os dois lados e o
*&   documento do parceiro:
*&
*&     compra ..: ZFI_WSYS_CAB-CODFORNECEDOR -> LFA1-STCD1 / STCD2
*&     venda ...: (ZSD034) KUNNR2 / KUNNR    -> KNA1-STCD1 / STCD2
*&
*&   A comparacao usa so os DIGITOS do documento e separa CNPJ de CPF
*&   (ver NORM_DOC), entao mascara e zeros a esquerda nao atrapalham.
*&
*& MODO DE EXTRACAO (bloco "Modo de extracao"):
*&   1) CRUZADO (padrao) - o comportamento historico: UMA saida, com o
*&      bloco do contrato e o bloco da ZSD034 na MESMA linha, ligados
*&      pelo CPF / CNPJ.
*&   2) SEPARADO - DOIS ALVs distintos, SEM cruzar compra com venda:
*&        ALV de cima ..: contratos de compra (bloco do cockpit)
*&        ALV de baixo .: saida da ZSD034, exatamente como ela devolve
*&      Cada lado obedece so aos SEUS filtros: o ALV de compra usa o
*&      bloco "Contratos" e o de venda usa o bloco "Vendas". Nenhuma
*&      linha e multiplicada, nenhuma e descartada por falta de par -
*&      os dois totais fecham com os relatorios de origem.
*&      No modo separado o bloco "Regra do cruzamento" fica inativo (nao
*&      ha cruzamento a regular) e a ZSD034 e submetida SEM a restricao
*&      de clientes, ou seja, traz o periodo inteiro do bloco de vendas.
*&
*&   Na tela, o modo separado desenha os dois ALVs empilhados num
*&   splitter sobre docking - cada um com sua propria barra de funcoes,
*&   seu layout e seu export.
*&
*& ARQUIVO DA EXTRACAO (modo separado):
*&   - UM xlsx com DUAS ABAS (padrao): aba "Compras (Contratos)" e aba
*&     "Vendas (ZSD034)" no mesmo arquivo; ou
*&   - DOIS xlsx separados: o nome informado ganha os sufixos _COMPRAS
*&     e _VENDAS antes da extensao.
*&   O arquivo de duas abas e montado aqui mesmo (OOXML + CL_ABAP_ZIP),
*&   sem GUI e sem ABAP2XLSX, porque quem grava e o servidor de
*&   aplicacao durante o Job. No modo cruzado nada muda: continua um
*&   arquivo unico gerado pelo CL_SALV_BS_LEX.
*&
*& Colunas (modo cruzado):
*&   1) Bloco CONTRATO - as mesmas colunas do cockpit ZCOCKPITECC:
*&      Centro, Contrato, Fornecedor, Nome do fornecedor, Safra,
*&      Dt Pgto, Volume Contrato Compra, Valor Contrato Compra,
*&      Situacao e Regiao (UF do centro) - MAIS o CNPJ, o CPF e a
*&      Inscricao Estadual do fornecedor.
*&      Volume e valor vem RENOMEADOS ("Volume Contrato Compra" e
*&      "Valor Contrato Compra"; no cockpit sao Vol Cont / Vlr Cont)
*&      para nao se confundirem com os numeros da venda.
*&   2) Coluna CRUZAMENTO: Compra+Venda / So compra / So venda.
*&      (no modo separado esta coluna nao aparece - nao ha cruzamento)
*&   3) Bloco ZSD034: TODAS as colunas da saida da ZSD034, com prefixo
*&      tecnico SD_ e o titulo comecando por "ZSD034 -".
*&
*& De onde vem a ZSD034:
*&   O programa NAO refaz a logica da ZSD034. Ele SUBMETE o programa
*&   dela (ZSDR006_NEW, fixo na constante GC_PROG) com os filtros do
*&   bloco de vendas e captura a saida em memoria pela CL_SALV_BS_RUNTIME_INFO
*&   (display = off). Assim os numeros sao exatamente os que a ZSD034
*&   mostra e manutencao futura dela reflete aqui sozinha - inclusive
*&   colunas novas, ja que o bloco ZSD034 e montado em TEMPO DE
*&   EXECUCAO (estrutura dinamica).
*&
*& Regra de juncao (chave = CPF / CNPJ) - so vale no modo CRUZADO:
*&   P_MAG  casa pelo Emissor da ordem  (parceiro AG - coluna KUNNR2)
*&   P_MWE  casa pelo Recebedor         (parceiro WE - coluna KUNNR)
*&   P_MAMB casa por qualquer um dos dois
*&   E um CRUZAMENTO, nao um resumo: fornecedor com 3 contratos e 10
*&   linhas de venda gera 30 linhas. Complementos:
*&     P_SEMVEN  traz tambem o contrato SEM venda (bloco ZSD034 vazio)
*&     P_SEMCTR  traz tambem a venda SEM contrato (bloco contrato vazio)
*&
*& Job de extracao automatica:
*&   Bloco "Job": diretorio (F4), nome do arquivo, data/hora da 1a
*&   execucao e a PERIODICIDADE DEFINIDA PELO USUARIO (a cada N
*&   minutos / horas / dias / semanas / meses; N = 0 agenda uma unica
*&   execucao). "Agendar Job" cria o agendamento recorrente com a
*&   selecao da tela, "Cancelar" remove os agendamentos futuros e
*&   "Situacao" mostra a proxima execucao.
*&   Na execucao em background (P_BG) o relatorio nao exibe ALV: gera
*&   o .xlsx e grava pelo OPEN DATASET no filesystem do SERVIDOR DE
*&   APLICACAO - o diretorio precisa existir e ser gravavel por ele
*&   (confira na AL11).
*&   No nome do arquivo valem os curingas &DATA& (AAAAMMDD) e &HORA&
*&   (HHMMSS). Sem curinga, cada execucao SOBRESCREVE o mesmo arquivo.
*&---------------------------------------------------------------------*
REPORT zfir051.

TABLES: zfi_wsys_cab, lfa1, vbak, vbap, mara, sscrfields.

*---------------------------------------------------------------------*
* CONSTANTES
*---------------------------------------------------------------------*
CONSTANTS:
* programa por tras da transacao ZSD034 (relatorio de carga). Fica fixo
* aqui de proposito: e o programa que este relatorio submete para pegar
* o bloco de colunas de venda. Se um dia a ZSD034 passar a apontar para
* outro programa (SE93), troque SO esta constante.
  gc_prog   TYPE progname VALUE 'ZSDR006_NEW',
* prefixo tecnico das colunas herdadas da ZSD034 (SD_VBELN, SD_KWMENG...)
  gc_pref   TYPE c LENGTH 3  VALUE 'SD_',
* prefixo do titulo dessas colunas, p/ ficarem identificadas na ALV
  gc_tit_sd TYPE c LENGTH 9  VALUE 'ZSD034 - ',
  BEGIN OF gc_st_ctr,                 " status do contrato (ZFI_WSYS_CAB)
    emandamento TYPE c LENGTH 1 VALUE 'A',
    finalizado  TYPE c LENGTH 1 VALUE 'F',
    cancelado   TYPE c LENGTH 1 VALUE 'C',
    washout     TYPE c LENGTH 1 VALUE 'T',
  END OF gc_st_ctr,
  BEGIN OF gc_cruz,                   " valores da coluna CRUZAMENTO
    ambos  TYPE c LENGTH 12 VALUE 'Compra+Venda',
    compra TYPE c LENGTH 12 VALUE 'So compra',
    venda  TYPE c LENGTH 12 VALUE 'So venda',
  END OF gc_cruz,
  BEGIN OF gc_aba,                    " nomes das abas / dos ALVs separados
    compras TYPE string VALUE 'Compras (Contratos)',
    vendas  TYPE string VALUE 'Vendas (ZSD034)',
  END OF gc_aba,
  BEGIN OF gc_suf,                    " sufixos do modo "dois arquivos"
    compras TYPE string VALUE '_COMPRAS',
    vendas  TYPE string VALUE '_VENDAS',
  END OF gc_suf,
  BEGIN OF gc_cel,                    " natureza da celula no xlsx
    texto  TYPE c LENGTH 1 VALUE 'T',
    numero TYPE c LENGTH 1 VALUE 'N',
    data   TYPE c LENGTH 1 VALUE 'D',
    hora   TYPE c LENGTH 1 VALUE 'H',
  END OF gc_cel,
  BEGIN OF gc_xf,                     " indices do cellXfs do styles.xml
    normal TYPE i VALUE 0,
    titulo TYPE i VALUE 1,
    data   TYPE i VALUE 2,
  END OF gc_xf.

*---------------------------------------------------------------------*
* TYPES
*---------------------------------------------------------------------*
* Bloco CONTRATO da saida. A ordem dos campos aqui e a ordem das
* primeiras colunas da ALV; o bloco da ZSD034 entra logo depois.
TYPES:
  BEGIN OF ty_ctr,
    empresa        TYPE zfi_wsys_cab-codigoempresa,
    centro         TYPE zfi_wsys_cab-centro,
    regiao         TYPE regio,                       " UF do centro (T001W)
    nr_contrato    TYPE zfi_wsys_cab-nr_contrato,
    codfornecedor  TYPE lifnr,
    nomefornecedor TYPE name1_gp,
    cnpj           TYPE c LENGTH 18,                 " 00.000.000/0000-00
    cpf            TYPE c LENGTH 14,                 " 000.000.000-00
    insc_estadual  TYPE stcd3,
    anosafra       TYPE c LENGTH 9,                  " "24 / 25"
    dt_pagamento   TYPE zfi_wsys_cab-datapagamento,
    vol_ctr_compra TYPE zfi_wsys_cab-quantidade,     " ex-"Vol Cont" do cockpit
    vlr_ctr_compra TYPE dmbtr,                       " ex-"Vlr Cont" do cockpit
    situacao       TYPE c LENGTH 15,
    cruzamento     TYPE c LENGTH 12,                 " Compra+Venda / So compra / So venda
  END OF ty_ctr,
  tt_ctr TYPE STANDARD TABLE OF ty_ctr WITH DEFAULT KEY.

* chave de cruzamento (documento do parceiro) -> linha da ZSD034
TYPES:
  BEGIN OF ty_key,
    doc TYPE c LENGTH 20,             " "J<digitos>" (CNPJ) / "F<digitos>" (CPF)
    idx TYPE i,                       " indice na saida da ZSD034
  END OF ty_key,
  tt_key TYPE SORTED TABLE OF ty_key WITH NON-UNIQUE KEY doc.

* de-para de componente ZSD034 -> componente da saida
TYPES:
  BEGIN OF ty_map,
    name_src TYPE abap_compname,
    name_dst TYPE abap_compname,
  END OF ty_map,
  tt_map TYPE STANDARD TABLE OF ty_map WITH DEFAULT KEY.

* titulo das colunas da ZSD034 (mesmos textos do fieldcat da ZSDR006_NEW)
TYPES:
  BEGIN OF ty_lbl,
    campo TYPE abap_compname,
    texto TYPE c LENGTH 40,
  END OF ty_lbl,
  tt_lbl TYPE SORTED TABLE OF ty_lbl WITH UNIQUE KEY campo.

* coluna de uma aba do xlsx: de onde ler, como titular e como escrever
TYPES:
  BEGIN OF ty_colx,
    campo  TYPE abap_compname,
    titulo TYPE string,
    tipo   TYPE c LENGTH 1,           " ver GC_CEL
  END OF ty_colx,
  tt_colx TYPE STANDARD TABLE OF ty_colx WITH DEFAULT KEY.

* uma aba do arquivo: nome, colunas e os dados
TYPES:
  BEGIN OF ty_aba,
    nome TYPE string,
    cols TYPE tt_colx,
    dref TYPE REF TO data,
  END OF ty_aba,
  tt_aba TYPE STANDARD TABLE OF ty_aba WITH DEFAULT KEY.

TYPES: tr_kunnr TYPE RANGE OF kunnr,
       tr_lifnr TYPE RANGE OF lifnr.

*---------------------------------------------------------------------*
* GLOBAIS
*---------------------------------------------------------------------*
DATA: gt_ctr    TYPE tt_ctr,
      gr_sd     TYPE REF TO data,     " saida da ZSD034 (tabela dinamica)
      gr_sd_pln TYPE REF TO data,     " idem, so com os campos elementares
      gr_out    TYPE REF TO data,     " saida final do modo cruzado
      gt_map    TYPE tt_map,          " ZSD034 -> bloco SD_ (modo cruzado)
      gt_map_sd TYPE tt_map,          " ZSD034 -> copia plana (modo separado)
      gt_lbl    TYPE tt_lbl,
      gv_sd_ok  TYPE abap_bool,       " a ZSD034 devolveu dados?
      gt_log    TYPE stringtab,       " mensagens da execucao (ver FORM LOG)
      gv_log    TYPE string,          " texto montado antes do PERFORM LOG
                                      " (PERFORM nao aceita expressao como parametro)
      gr_salv   TYPE REF TO cl_salv_table,
      gr_salv_c TYPE REF TO cl_salv_table,   " ALV de compras (modo separado)
      gr_salv_v TYPE REF TO cl_salv_table,   " ALV de vendas  (modo separado)
      go_dock   TYPE REF TO cl_gui_docking_container,
      go_split  TYPE REF TO cl_gui_splitter_container.

FIELD-SYMBOLS: <gt_sd>  TYPE STANDARD TABLE,
               <gt_sdp> TYPE STANDARD TABLE,
               <gt_out> TYPE STANDARD TABLE.

*---------------------------------------------------------------------*
* SELECTION SCREEN
*---------------------------------------------------------------------*
* Bloco 1 - contratos (mesmos filtros do cockpit ZCOCKPITECC)
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE text-b01.
  PARAMETERS:     p_bukrs TYPE bukrs OBLIGATORY DEFAULT 'BSSE' MEMORY ID buk.
  SELECT-OPTIONS: s_contr FOR zfi_wsys_cab-nr_contrato,
                  s_dtctr FOR zfi_wsys_cab-dt_contrato,
                  s_dtpag FOR zfi_wsys_cab-datapagamento,
                  s_safra FOR zfi_wsys_cab-anosafra,
                  s_forn  FOR zfi_wsys_cab-codfornecedor,
                  s_cctr  FOR zfi_wsys_cab-centro,
                  s_cnpj  FOR lfa1-stcd1,
                  s_cpf   FOR lfa1-stcd2,
                  s_ie    FOR lfa1-stcd3,
                  s_stat  FOR zfi_wsys_cab-status.
SELECTION-SCREEN END OF BLOCK b1.

* Bloco 2 - vendas (repassados a ZSD034 no SUBMIT)
SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE text-b02.
  SELECT-OPTIONS: s_audat FOR vbak-erdat OBLIGATORY,  " data de criacao da OV
                  s_vbeln FOR vbak-vbeln,
                  s_vkorg FOR vbak-vkorg,
                  s_vtweg FOR vbak-vtweg,
                  s_spart FOR vbak-spart,
                  s_vkbur FOR vbak-vkbur,
                  s_vkgrp FOR vbak-vkgrp,
                  s_auart FOR vbak-auart,
                  s_anosf FOR vbak-zzanosafra,
                  s_wsd   FOR vbap-werks,
                  s_matnr FOR mara-matnr,
                  s_matkl FOR mara-matkl.
SELECTION-SCREEN END OF BLOCK b2.

* Bloco 3 - modo de extracao (uma saida cruzada x dois ALVs separados)
SELECTION-SCREEN BEGIN OF BLOCK bmod WITH FRAME TITLE text-b05.
  PARAMETERS: p_mcruz RADIOBUTTON GROUP gmod DEFAULT 'X'   " cruza compra x venda
                                             USER-COMMAND uc_mod,
              p_msep  RADIOBUTTON GROUP gmod.              " dois ALVs, sem cruzar
  SELECTION-SCREEN SKIP.
* so valem no modo separado: no cruzado a saida e uma so
  PARAMETERS: p_x1arq RADIOBUTTON GROUP gxls DEFAULT 'X'   " 1 arquivo, 2 abas
                                             MODIF ID xls,
              p_x2arq RADIOBUTTON GROUP gxls MODIF ID xls. " 2 arquivos
SELECTION-SCREEN END OF BLOCK bmod.

* Bloco 4 - regra do cruzamento (inativo no modo separado)
SELECTION-SCREEN BEGIN OF BLOCK b3 WITH FRAME TITLE text-b03.
  PARAMETERS: p_mag  RADIOBUTTON GROUP gmat DEFAULT 'X'    " emissor da ordem (AG)
                                            MODIF ID crz,
              p_mwe  RADIOBUTTON GROUP gmat MODIF ID crz,  " recebedor (WE)
              p_mamb RADIOBUTTON GROUP gmat MODIF ID crz.  " qualquer um dos dois
  SELECTION-SCREEN SKIP.
  PARAMETERS: p_semven AS CHECKBOX DEFAULT 'X' MODIF ID crz,  " contrato sem venda
              p_semctr AS CHECKBOX MODIF ID crz.              " venda sem contrato
SELECTION-SCREEN END OF BLOCK b3.

* Bloco 5 - Arquivo e Job de extracao automatica
SELECTION-SCREEN BEGIN OF BLOCK b4 WITH FRAME TITLE text-b04.
* Destino do arquivo. O F4 do diretorio navega no PC, entao e facil
* apontar sem querer uma pasta da estacao de trabalho num destino de
* servidor - por isso o destino e explicito, e nao adivinhado.
  PARAMETERS: p_dsrv  RADIOBUTTON GROUP gdst DEFAULT 'X',  " servidor de aplicacao (AL11)
              p_dpc   RADIOBUTTON GROUP gdst.              " estacao de trabalho (so em dialogo)
* Gera o arquivo TAMBEM na execucao em primeiro plano, alem de exibir a
* ALV. Sem isto o arquivo so sai pelo Job / background.
  PARAMETERS: p_gerar AS CHECKBOX.
  SELECTION-SCREEN SKIP.
  PARAMETERS: p_xdir  TYPE string LOWER CASE,             " diretorio de destino
              p_xname TYPE string LOWER CASE,             " nome do arquivo (&DATA& / &HORA&)
              p_jname TYPE tbtcjob-jobname DEFAULT 'ZFI051_EXTRACAO',
              p_jdata TYPE sy-datum,                      " 1a execucao - data
              p_jhora TYPE sy-uzeit.                      " 1a execucao - hora
  SELECTION-SCREEN BEGIN OF LINE.
    SELECTION-SCREEN COMMENT 1(31) tx_rep.
    PARAMETERS p_pint TYPE i DEFAULT 1.                   " 0 = executa uma vez so
  SELECTION-SCREEN END OF LINE.
  PARAMETERS: p_umin RADIOBUTTON GROUP gper,              " minuto(s)
              p_uhor RADIOBUTTON GROUP gper,              " hora(s)
              p_udia RADIOBUTTON GROUP gper DEFAULT 'X',  " dia(s)
              p_usem RADIOBUTTON GROUP gper,              " semana(s)
              p_umes RADIOBUTTON GROUP gper.              " mes(es)
  SELECTION-SCREEN SKIP.
  SELECTION-SCREEN PUSHBUTTON /1(45) pb_ag USER-COMMAND uc_ag.
  SELECTION-SCREEN PUSHBUTTON /1(45) pb_ca USER-COMMAND uc_ca.
  SELECTION-SCREEN PUSHBUTTON /1(45) pb_st USER-COMMAND uc_st.
SELECTION-SCREEN END OF BLOCK b4.

* Internos (preenchidos pelo SUBMIT do Job): caminho completo + modo batch
PARAMETERS: p_xfile TYPE string NO-DISPLAY,
            p_bg    TYPE c      NO-DISPLAY.

*---------------------------------------------------------------------*
INITIALIZATION.
*---------------------------------------------------------------------*
  PERFORM carregar_labels.

  pb_ag  = 'Agendar Job de extracao'.
  pb_ca  = 'Cancelar agendamento do Job'.
  pb_st  = 'Situacao do agendamento'.
  tx_rep = 'Repetir a cada'.

  IF p_jdata IS INITIAL.
    p_jdata = sy-datum.
  ENDIF.
  IF p_jhora IS INITIAL.
    p_jhora = '000100'.                " 00:01 (mesmo padrao do Job da ZFI042)
  ENDIF.
* Com os curingas, cada execucao gera um arquivo proprio. Tire o &DATA&
* / &HORA& se quiser que o Job sobrescreva sempre o mesmo arquivo.
  IF p_xname IS INITIAL.
    p_xname = 'ZFI051_&DATA&_&HORA&.xlsx'.
  ENDIF.

*---------------------------------------------------------------------*
AT SELECTION-SCREEN OUTPUT.
*---------------------------------------------------------------------*
* A regra do cruzamento so faz sentido quando ha cruzamento; a escolha
* entre um arquivo de duas abas e dois arquivos, so quando ha duas
* saidas. Cada um fica cinza no modo em que nao se aplica.
  LOOP AT SCREEN.
    CASE screen-group1.
      WHEN 'CRZ'.
        IF p_msep = abap_true.
          screen-input = '0'.
        ELSE.
          screen-input = '1'.
        ENDIF.
        MODIFY SCREEN.
      WHEN 'XLS'.
        IF p_msep = abap_true.
          screen-input = '1'.
        ELSE.
          screen-input = '0'.
        ENDIF.
        MODIFY SCREEN.
    ENDCASE.
  ENDLOOP.

*---------------------------------------------------------------------*
AT SELECTION-SCREEN.
*---------------------------------------------------------------------*
* Botoes do bloco do Job: administram o agendamento e NAO executam o
* relatorio agora.
  CASE sscrfields-ucomm.
    WHEN 'UC_AG'.
      PERFORM agendar_job.
      RETURN.
    WHEN 'UC_CA'.
      PERFORM cancelar_job.
      RETURN.
    WHEN 'UC_ST'.
      PERFORM situacao_job.
      RETURN.
  ENDCASE.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_xdir.
  PERFORM f4_diretorio.

*---------------------------------------------------------------------*
START-OF-SELECTION.
*---------------------------------------------------------------------*
* Em background nao ha GUI para ALV nenhuma: a saida e sempre arquivo.
* Assim vale tanto o Job agendado pelo botao (que ja manda P_BG = 'X')
* quanto uma execucao solta por F9 / SM36.
  IF sy-batch = abap_true.
    p_bg = abap_true.
  ENDIF.

* Gravar no PC exige GUI: em background so existe o servidor.
  IF p_bg = abap_true AND p_dpc = abap_true.
    PERFORM log USING 'Destino "estacao de trabalho" nao existe em background - gravando no servidor.'.
    p_dsrv = abap_true.
    p_dpc  = abap_false.
  ENDIF.

* O nome do arquivo e resolvido a cada execucao (curingas &DATA& / &HORA&).
  IF p_bg = abap_true OR p_gerar = abap_true.
    PERFORM montar_arquivo CHANGING p_xfile.
  ENDIF.

  PERFORM carregar_contratos.
  PERFORM ler_zsd034.

  IF p_msep = abap_true.
    PERFORM montar_saida_separada.      " dois blocos independentes
*   Os dois lados sao independentes, entao um pode vir vazio sem que o
*   outro perceba. Registrar a contagem dos DOIS aqui e o que evita o
*   "so trouxe a ZSD034" sem explicacao.
    PERFORM contar_lados.
  ELSE.
    PERFORM montar_saida.               " uma saida cruzada
  ENDIF.

  DATA gv_vazio TYPE abap_bool.
  PERFORM saida_vazia CHANGING gv_vazio.

  IF gv_vazio = abap_true.
*   sem linha nenhuma o programa nao tem o que exibir nem o que gravar.
*   A mensagem diz quantas linhas vieram de CADA lado: sem isso, "nao
*   aconteceu nada" nao distingue filtro sem resultado de erro na
*   captura da ZSD034.
    DATA: gv_nsd  TYPE i,
          gv_diag TYPE string.
    IF <gt_sd> IS ASSIGNED.
      gv_nsd = lines( <gt_sd> ).       " a saida crua da ZSD034
    ENDIF.
    gv_diag = |Nenhuma linha para os filtros informados. | &&
              |Contratos de compra: { lines( gt_ctr ) }. Linhas da ZSD034: { gv_nsd }.|.
    IF gv_sd_ok = abap_false.
      gv_diag = gv_diag && ' A ZSD034 nao devolveu dados para os filtros do bloco Vendas.'.
    ENDIF.
    PERFORM log USING gv_diag.
    PERFORM mostrar_log.
    RETURN.
  ENDIF.

* Arquivo: sempre em background; em primeiro plano, so se P_GERAR pedir
  IF p_bg = abap_true OR p_gerar = abap_true.
    IF p_msep = abap_true.
      PERFORM exportar_separado.
    ELSE.
      PERFORM exportar_xlsx.
    ENDIF.
  ENDIF.

* ALV: so em primeiro plano
  IF p_bg = abap_false.
    PERFORM mostrar_log.
    IF p_msep = abap_true.
      PERFORM exibir_alv_separado.
    ELSE.
      PERFORM exibir_alv.
    ENDIF.
  ENDIF.

*&---------------------------------------------------------------------*
*& LOG / MOSTRAR_LOG
*&
*& Em background o WRITE vai para a lista do job - que so aparece no
*& SPOOL, nao no log do job. Em primeiro plano o WRITE quebraria a ALV,
*& e uma MESSAGE TYPE 'S' se perde no rodape. Entao as mensagens sao
*& acumuladas e, em dialogo, saem juntas num popup - que ninguem deixa
*& de ver.
*&---------------------------------------------------------------------*
FORM log USING iv_txt TYPE clike.
  APPEND iv_txt TO gt_log.
  IF p_bg = abap_true.
    WRITE: / iv_txt.
  ENDIF.
ENDFORM.

FORM mostrar_log.

  DATA lv_txt TYPE string.

  CHECK p_bg = abap_false.
  CHECK gt_log IS NOT INITIAL.

  LOOP AT gt_log INTO DATA(ls_l).
    IF lv_txt IS INITIAL.
      lv_txt = ls_l.
    ELSE.
      lv_txt = |{ lv_txt } { ls_l }|.
    ENDIF.
  ENDLOOP.
  CLEAR gt_log.

  MESSAGE lv_txt TYPE 'I'.

ENDFORM.

*&---------------------------------------------------------------------*
*& CONTAR_LADOS  (quantas linhas vieram de cada lado, no modo separado)
*&
*& No modo separado a compra e a venda nao se falam: uma pode vir vazia
*& e a outra cheia, e o arquivo sai com uma aba so com cabecalho. Sem
*& esta contagem o usuario nao tem como distinguir "filtro do bloco
*& Contratos nao achou nada" de "o programa deixou de gerar".
*&---------------------------------------------------------------------*
FORM contar_lados.

  DATA lv_nv TYPE i.

  IF <gt_sdp> IS ASSIGNED.
    lv_nv = lines( <gt_sdp> ).
  ENDIF.

  gv_log = |Contratos de compra: { lines( gt_ctr ) } linha(s). | &&
           |Vendas ZSD034: { lv_nv } linha(s).|.
  PERFORM log USING gv_log.

  IF gt_ctr IS INITIAL AND lv_nv > 0.
    gv_log = 'O bloco Contratos nao achou contrato nenhum. Se o filtro usado foi a Data do '
          && 'pagamento, lembre que ela fica vazia em contrato ainda nao pago - nesse caso '
          && 'filtre pela Data do contrato ou pelo Ano-safra.'.
    PERFORM log USING gv_log.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& SAIDA_VAZIA  (nao ha o que exibir / gravar?)
*&---------------------------------------------------------------------*
FORM saida_vazia CHANGING cv_vazio TYPE abap_bool.

  cv_vazio = abap_true.

  IF p_msep = abap_true.
*   no modo separado basta UM dos lados ter linha: cada ALV vive sozinho
    IF gt_ctr IS NOT INITIAL.
      cv_vazio = abap_false.
    ELSEIF <gt_sdp> IS ASSIGNED AND <gt_sdp> IS NOT INITIAL.
      cv_vazio = abap_false.
    ENDIF.
  ELSE.
    IF <gt_out> IS ASSIGNED AND <gt_out> IS NOT INITIAL.
      cv_vazio = abap_false.
    ENDIF.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& CARREGAR_CONTRATOS
*&
*& Bloco de compra: o mesmo conjunto de colunas do cockpit ZCOCKPITECC
*& (print da tela), acrescido de CNPJ / CPF / Inscricao Estadual. So o
*& necessario - sem romaneio, checklist ou vinculo, que aqui nao entram.
*&---------------------------------------------------------------------*
FORM carregar_contratos.

  DATA: lt_lif TYPE STANDARD TABLE OF lifnr,
        lr_lif TYPE tr_lifnr.

  CLEAR gt_ctr.

  SELECT * FROM zfi_wsys_cab
           INTO TABLE @DATA(lt_cab)
           WHERE codigoempresa = @p_bukrs
             AND nr_contrato   IN @s_contr
             AND dt_contrato   IN @s_dtctr
             AND datapagamento IN @s_dtpag
             AND anosafra      IN @s_safra
             AND codfornecedor IN @s_forn
             AND centro        IN @s_cctr
             AND status        IN @s_stat
           ORDER BY codfornecedor, nr_contrato.
  IF lt_cab IS INITIAL.
    RETURN.
  ENDIF.

* filtro por CNPJ / CPF / IE do fornecedor (mesma logica do cockpit)
  IF s_cnpj[] IS NOT INITIAL OR s_cpf[] IS NOT INITIAL OR s_ie[] IS NOT INITIAL.
    IF s_cnpj[] IS NOT INITIAL.
      SELECT lifnr FROM lfa1 APPENDING TABLE @lt_lif
             FOR ALL ENTRIES IN @lt_cab
             WHERE lifnr = @lt_cab-codfornecedor
               AND stcd1 IN @s_cnpj.
    ENDIF.
    IF s_cpf[] IS NOT INITIAL.
      SELECT lifnr FROM lfa1 APPENDING TABLE @lt_lif
             FOR ALL ENTRIES IN @lt_cab
             WHERE lifnr = @lt_cab-codfornecedor
               AND stcd2 IN @s_cpf.
    ENDIF.
    IF s_ie[] IS NOT INITIAL.
      SELECT lifnr FROM lfa1 APPENDING TABLE @lt_lif
             FOR ALL ENTRIES IN @lt_cab
             WHERE lifnr = @lt_cab-codfornecedor
               AND stcd3 IN @s_ie.
    ENDIF.
    SORT lt_lif.
    DELETE ADJACENT DUPLICATES FROM lt_lif.
    LOOP AT lt_lif INTO DATA(lv_lif).
      APPEND VALUE #( sign = 'I' option = 'EQ' low = lv_lif ) TO lr_lif.
    ENDLOOP.
    DELETE lt_cab WHERE codfornecedor NOT IN lr_lif.
    IF lt_cab IS INITIAL.
      RETURN.
    ENDIF.
  ENDIF.

* dados do fornecedor (nome, CNPJ, CPF e Inscricao Estadual)
  SELECT lifnr, name1, stcd1, stcd2, stcd3 FROM lfa1 INTO TABLE @DATA(lt_lfa1)
         FOR ALL ENTRIES IN @lt_cab
         WHERE lifnr = @lt_cab-codfornecedor.
  SORT lt_lfa1 BY lifnr.

* UF (regiao) do centro do contrato. So os contratos COM centro entram no
* FOR ALL ENTRIES - tabela driver vazia leria a T001W inteira.
  DATA(lt_cwrk) = lt_cab.
  DELETE lt_cwrk WHERE centro IS INITIAL.
  SORT lt_cwrk BY centro.
  DELETE ADJACENT DUPLICATES FROM lt_cwrk COMPARING centro.

  DATA lt_werks TYPE STANDARD TABLE OF t001w.
  IF lt_cwrk IS NOT INITIAL.
    SELECT werks, name1, ort01, regio
           FROM t001w
           INTO CORRESPONDING FIELDS OF TABLE @lt_werks
           FOR ALL ENTRIES IN @lt_cwrk
           WHERE werks = @lt_cwrk-centro.
    SORT lt_werks BY werks.
  ENDIF.

  LOOP AT lt_cab INTO DATA(ls_cab).

    DATA(ls_ctr) = VALUE ty_ctr(
      empresa        = ls_cab-codigoempresa
      centro         = ls_cab-centro
      nr_contrato    = ls_cab-nr_contrato
      codfornecedor  = ls_cab-codfornecedor
      dt_pagamento   = ls_cab-datapagamento
      vol_ctr_compra = ls_cab-quantidade
      vlr_ctr_compra = ls_cab-valor_contrato ).

*   safra "2425" -> "24 / 25"; formato diferente sai como veio
    IF strlen( ls_cab-anosafra ) = 4 AND ls_cab-anosafra CO '0123456789'.
      ls_ctr-anosafra = |{ ls_cab-anosafra(2) } / { ls_cab-anosafra+2(2) }|.
    ELSE.
      ls_ctr-anosafra = ls_cab-anosafra.
    ENDIF.

    CASE ls_cab-status.
      WHEN gc_st_ctr-cancelado.  ls_ctr-situacao = 'Cancelado'.
      WHEN gc_st_ctr-washout.    ls_ctr-situacao = 'WashOut'.
      WHEN gc_st_ctr-finalizado. ls_ctr-situacao = 'Finalizado'.
      WHEN OTHERS.               ls_ctr-situacao = 'Em andamento'.
    ENDCASE.

    IF ls_cab-centro IS NOT INITIAL.
      READ TABLE lt_werks INTO DATA(ls_werks)
           WITH KEY werks = ls_cab-centro BINARY SEARCH.
      IF sy-subrc = 0.
        ls_ctr-regiao = ls_werks-regio.
      ENDIF.
    ENDIF.

    READ TABLE lt_lfa1 INTO DATA(ls_lfa1)
         WITH KEY lifnr = ls_cab-codfornecedor BINARY SEARCH.
    IF sy-subrc = 0.
      ls_ctr-nomefornecedor = ls_lfa1-name1.
      ls_ctr-insc_estadual  = ls_lfa1-stcd3.
      PERFORM formata_doc USING ls_lfa1-stcd1 ls_lfa1-stcd2
                          CHANGING ls_ctr-cnpj ls_ctr-cpf.
    ENDIF.

    APPEND ls_ctr TO gt_ctr.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& FORMATA_DOC  (CNPJ / CPF com mascara, igual ao cockpit)
*&---------------------------------------------------------------------*
FORM formata_doc USING iv_stcd1 TYPE lfa1-stcd1
                       iv_stcd2 TYPE lfa1-stcd2
                 CHANGING cv_cnpj TYPE any
                          cv_cpf  TYPE any.
  CLEAR: cv_cnpj, cv_cpf.
  IF iv_stcd1 IS NOT INITIAL.
    CALL FUNCTION 'CONVERSION_EXIT_CGCBR_OUTPUT'
      EXPORTING input  = iv_stcd1
      IMPORTING output = cv_cnpj.
  ENDIF.
  IF iv_stcd2 IS NOT INITIAL.
    CALL FUNCTION 'CONVERSION_EXIT_CPFBR_OUTPUT'
      EXPORTING input  = iv_stcd2
      IMPORTING output = cv_cpf.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& NORM_DOC  (chave de cruzamento a partir do CNPJ / CPF)
*&
*& Fica so com os DIGITOS e derruba os zeros a esquerda: o mesmo
*& documento casa gravado com mascara, sem mascara ou com zeros de
*& preenchimento. O tipo entra na frente ("J" = CNPJ / STCD1, "F" = CPF
*& / STCD2) para um CNPJ nunca casar com um CPF por coincidencia de
*& digitos.
*&---------------------------------------------------------------------*
FORM norm_doc USING iv_doc  TYPE any
                    iv_tipo TYPE c
              CHANGING cv_key TYPE any.
  DATA lv_s TYPE string.

  CLEAR cv_key.
  lv_s = iv_doc.
  REPLACE ALL OCCURRENCES OF REGEX '[^0-9]' IN lv_s WITH ''.
  SHIFT lv_s LEFT DELETING LEADING '0'.
  CHECK lv_s IS NOT INITIAL.
  cv_key = |{ iv_tipo }{ lv_s }|.
ENDFORM.

*&---------------------------------------------------------------------*
*& LER_ZSD034
*&
*& Submete o programa da ZSD034 (constante GC_PROG = ZSDR006_NEW) com os
*& filtros do bloco de vendas e captura a saida em memoria, sem exibir
*& a ALV dela. A CL_SALV_BS_RUNTIME_INFO intercepta tanto o
*& CL_SALV_TABLE quanto o REUSE_ALV_GRID_DISPLAY (que e o caso da
*& ZSDR006_NEW).
*&
*& No modo CRUZADO, quando o cruzamento e pelo EMISSOR DA ORDEM e nao se
*& pedem as vendas sem contrato, a selecao da ZSD034 ja vai restrita aos
*& clientes cujo CPF / CNPJ tem contrato (S_KUNNR = filtro de
*& VBAK-KUNNR) - e o que evita ler a base de vendas inteira do periodo.
*& No modo SEPARADO essa restricao NAO se aplica: o ALV de vendas tem de
*& mostrar a ZSD034 inteira do periodo, independente de haver contrato.
*&---------------------------------------------------------------------*
FORM ler_zsd034.

  DATA: lr_kunnr TYPE tr_kunnr,
        lv_msg   TYPE string.

  CLEAR gv_sd_ok.

  IF p_mcruz = abap_true AND p_mag = abap_true AND p_semctr = abap_false.
    PERFORM clientes_dos_contratos CHANGING lr_kunnr.
    IF lr_kunnr IS INITIAL.
*     nenhum cliente com o documento dos fornecedores selecionados: nao
*     ha o que cruzar, mas os contratos ainda podem sair sozinhos
      RETURN.
    ENDIF.
  ENDIF.

  cl_salv_bs_runtime_info=>set( EXPORTING display  = abap_false
                                          metadata = abap_false
                                          data     = abap_true ).
  TRY.
      SUBMIT (gc_prog)
        WITH s_audat IN s_audat
        WITH s_vbeln IN s_vbeln
        WITH s_vkorg IN s_vkorg
        WITH s_vtweg IN s_vtweg
        WITH s_spart IN s_spart
        WITH s_vkbur IN s_vkbur
        WITH s_vkgrp IN s_vkgrp
        WITH s_auart IN s_auart
        WITH s_anosf IN s_anosf
        WITH s_werks IN s_wsd
        WITH s_matnr IN s_matnr
        WITH s_matkl IN s_matkl
        WITH s_kunnr IN lr_kunnr
        AND RETURN.

      cl_salv_bs_runtime_info=>get_data_ref( IMPORTING r_data = gr_sd ).

    CATCH cx_salv_bs_sc_runtime_info.
      CLEAR gr_sd.
  ENDTRY.
  cl_salv_bs_runtime_info=>clear_all( ).

  IF gr_sd IS INITIAL.
    lv_msg = |Nao foi possivel capturar a saida da ZSD034 ({ gc_prog }). | &&
             |Sai so o bloco de contratos.|.
    IF p_bg = abap_true.
      PERFORM log USING lv_msg.
    ELSE.
      MESSAGE lv_msg TYPE 'S' DISPLAY LIKE 'W'.
    ENDIF.
    RETURN.
  ENDIF.

  ASSIGN gr_sd->* TO <gt_sd>.
  IF <gt_sd> IS ASSIGNED AND <gt_sd> IS NOT INITIAL.
    gv_sd_ok = abap_true.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& CLIENTES_DOS_CONTRATOS
*&
*& Clientes (KNA1) cujo CNPJ / CPF e o mesmo dos fornecedores dos
*& contratos selecionados. Vira o filtro S_KUNNR do SUBMIT da ZSD034.
*&---------------------------------------------------------------------*
FORM clientes_dos_contratos CHANGING ct_kunnr TYPE tr_kunnr.

  DATA: lr_st1 TYPE RANGE OF kna1-stcd1,
        lr_st2 TYPE RANGE OF kna1-stcd2,
        lt_kun TYPE STANDARD TABLE OF kunnr.

  CLEAR ct_kunnr.
  CHECK gt_ctr IS NOT INITIAL.

  SELECT lifnr, stcd1, stcd2 FROM lfa1 INTO TABLE @DATA(lt_lfa1)
         FOR ALL ENTRIES IN @gt_ctr
         WHERE lifnr = @gt_ctr-codfornecedor.
  CHECK sy-subrc = 0.

  LOOP AT lt_lfa1 INTO DATA(ls_lfa1).
    IF ls_lfa1-stcd1 IS NOT INITIAL.
      APPEND VALUE #( sign = 'I' option = 'EQ' low = ls_lfa1-stcd1 ) TO lr_st1.
    ENDIF.
    IF ls_lfa1-stcd2 IS NOT INITIAL.
      APPEND VALUE #( sign = 'I' option = 'EQ' low = ls_lfa1-stcd2 ) TO lr_st2.
    ENDIF.
  ENDLOOP.
  SORT lr_st1 BY low. DELETE ADJACENT DUPLICATES FROM lr_st1 COMPARING low.
  SORT lr_st2 BY low. DELETE ADJACENT DUPLICATES FROM lr_st2 COMPARING low.

  IF lr_st1 IS NOT INITIAL.
    SELECT kunnr FROM kna1 APPENDING TABLE @lt_kun WHERE stcd1 IN @lr_st1.
  ENDIF.
  IF lr_st2 IS NOT INITIAL.
    SELECT kunnr FROM kna1 APPENDING TABLE @lt_kun WHERE stcd2 IN @lr_st2.
  ENDIF.
  SORT lt_kun.
  DELETE ADJACENT DUPLICATES FROM lt_kun.

  LOOP AT lt_kun INTO DATA(lv_kun).
    APPEND VALUE #( sign = 'I' option = 'EQ' low = lv_kun ) TO ct_kunnr.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& MONTAR_SAIDA  (modo CRUZADO)
*&
*& Monta EM TEMPO DE EXECUCAO a estrutura da saida = bloco do contrato
*& (TY_CTR) + todas as colunas elementares da saida da ZSD034, com o
*& prefixo SD_. Depois cruza os dois lados pelo CPF / CNPJ.
*&---------------------------------------------------------------------*
FORM montar_saida.

  DATA: lt_comp TYPE cl_abap_structdescr=>component_table,
        ls_comp TYPE abap_componentdescr,
        lt_src  TYPE cl_abap_structdescr=>component_table,
        ls_ctr  TYPE ty_ctr,
        lv_nome TYPE abap_compname,
        lv_idx  TYPE i.

* 1) componentes do bloco do contrato -----------------------------
  DATA(lo_ctr) = CAST cl_abap_structdescr(
                   cl_abap_typedescr=>describe_by_data( ls_ctr ) ).
  lt_comp = lo_ctr->get_components( ).

* 2) componentes herdados da ZSD034 -------------------------------
  CLEAR gt_map.
  IF gv_sd_ok = abap_true.
    DATA(lo_sdtab)  = CAST cl_abap_tabledescr(
                        cl_abap_typedescr=>describe_by_data( <gt_sd> ) ).
    DATA(lo_sdline) = CAST cl_abap_structdescr( lo_sdtab->get_table_line_type( ) ).
    PERFORM comp_flat USING lo_sdline '' CHANGING lt_src.

    LOOP AT lt_src INTO DATA(ls_src).
*     colunas nao elementares (tabela de cores, estruturas aninhadas)
*     ficam de fora: a ALV nao as exibe e ainda derrubariam o SALV
      CHECK ls_src-type->kind = cl_abap_typedescr=>kind_elem.

      lv_nome = |{ gc_pref }{ ls_src-name }|.
      READ TABLE lt_comp TRANSPORTING NO FIELDS WITH KEY name = lv_nome.
      CHECK sy-subrc <> 0.             " nome ja usado: ignora a coluna

      CLEAR ls_comp.
      ls_comp-name = lv_nome.
      ls_comp-type = ls_src-type.
      APPEND ls_comp TO lt_comp.
      APPEND VALUE ty_map( name_src = ls_src-name
                           name_dst = lv_nome ) TO gt_map.
    ENDLOOP.
  ENDIF.

* 3) tabela de saida dinamica -------------------------------------
  DATA: lo_struct TYPE REF TO cl_abap_structdescr,
        lo_table  TYPE REF TO cl_abap_tabledescr.
  TRY.
      lo_struct = cl_abap_structdescr=>create( lt_comp ).
      lo_table  = cl_abap_tabledescr=>create( p_line_type = lo_struct ).
    CATCH cx_sy_struct_creation cx_sy_table_creation INTO DATA(lx_cr).
      MESSAGE lx_cr->get_text( ) TYPE 'E'.
      RETURN.
  ENDTRY.
  CREATE DATA gr_out TYPE HANDLE lo_table.
  ASSIGN gr_out->* TO <gt_out>.

* 4) indice das vendas por documento (CPF / CNPJ) -----------------
  DATA: lt_sdkey TYPE tt_key,
        lt_usado TYPE HASHED TABLE OF i WITH UNIQUE KEY table_line.
  IF gv_sd_ok = abap_true.
    PERFORM indexar_vendas CHANGING lt_sdkey.
  ENDIF.

* 5) cruzamento ---------------------------------------------------
  FIELD-SYMBOLS: <ls_out> TYPE any,
                 <ls_sd>  TYPE any,
                 <lv_cz>  TYPE any.
  DATA: lr_out  TYPE REF TO data,
        lv_doc1 TYPE c LENGTH 20,
        lv_doc2 TYPE c LENGTH 20,
        lv_doc  TYPE c LENGTH 20,
        lv_from TYPE i,
        lv_ok   TYPE abap_bool.

  CREATE DATA lr_out TYPE HANDLE lo_struct.
  ASSIGN lr_out->* TO <ls_out>.

  LOOP AT gt_ctr INTO ls_ctr.

    CLEAR: lv_doc1, lv_doc2, lv_ok.
    PERFORM norm_doc USING ls_ctr-cnpj 'J' CHANGING lv_doc1.
    PERFORM norm_doc USING ls_ctr-cpf  'F' CHANGING lv_doc2.

    DO 2 TIMES.
      IF sy-index = 1.
        lv_doc = lv_doc1.
      ELSE.
        lv_doc = lv_doc2.
      ENDIF.
      CHECK lv_doc IS NOT INITIAL.

      READ TABLE lt_sdkey TRANSPORTING NO FIELDS WITH KEY doc = lv_doc.
      CHECK sy-subrc = 0.
      lv_from = sy-tabix.

      LOOP AT lt_sdkey INTO DATA(ls_k) FROM lv_from.
        IF ls_k-doc <> lv_doc.
          EXIT.
        ENDIF.
        READ TABLE <gt_sd> ASSIGNING <ls_sd> INDEX ls_k-idx.
        CHECK sy-subrc = 0.

        CLEAR <ls_out>.
        ls_ctr-cruzamento = gc_cruz-ambos.
        MOVE-CORRESPONDING ls_ctr TO <ls_out>.
        PERFORM copiar_map USING <ls_sd> gt_map CHANGING <ls_out>.
        APPEND <ls_out> TO <gt_out>.

        INSERT ls_k-idx INTO TABLE lt_usado.
        lv_ok = abap_true.
      ENDLOOP.
    ENDDO.

*   contrato sem venda casada
    IF lv_ok = abap_false AND p_semven = abap_true.
      CLEAR <ls_out>.
      ls_ctr-cruzamento = gc_cruz-compra.
      MOVE-CORRESPONDING ls_ctr TO <ls_out>.
      APPEND <ls_out> TO <gt_out>.
    ENDIF.

  ENDLOOP.

* 6) vendas sem contrato ------------------------------------------
  IF p_semctr = abap_true AND gv_sd_ok = abap_true.
    LOOP AT <gt_sd> ASSIGNING <ls_sd>.
      lv_idx = sy-tabix.
      READ TABLE lt_usado TRANSPORTING NO FIELDS WITH TABLE KEY table_line = lv_idx.
      CHECK sy-subrc <> 0.

      CLEAR <ls_out>.
      PERFORM copiar_map USING <ls_sd> gt_map CHANGING <ls_out>.
      ASSIGN COMPONENT 'CRUZAMENTO' OF STRUCTURE <ls_out> TO <lv_cz>.
      IF sy-subrc = 0.
        <lv_cz> = gc_cruz-venda.
      ENDIF.
      APPEND <ls_out> TO <gt_out>.
    ENDLOOP.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& MONTAR_SAIDA_SEPARADA  (modo SEPARADO)
*&
*& Aqui nao ha o que montar do lado da compra: GT_CTR ja e o ALV de
*& contratos, com as mesmas colunas do modo cruzado (menos a coluna
*& CRUZAMENTO, escondida na exibicao).
*&
*& Do lado da venda o que se faz e uma copia PLANA da saida da ZSD034:
*& mesmos nomes de campo, mesmos valores, so que sem os componentes nao
*& elementares (tabela de cores, estruturas aninhadas) que a ALV nao
*& exibe e que derrubariam o CL_SALV_TABLE. Nada e agregado, filtrado ou
*& reordenado - a aba de vendas fecha linha a linha com a ZSD034.
*&---------------------------------------------------------------------*
FORM montar_saida_separada.

  DATA: lt_comp TYPE cl_abap_structdescr=>component_table,
        ls_comp TYPE abap_componentdescr,
        lt_src  TYPE cl_abap_structdescr=>component_table.

  CLEAR gt_map_sd.
  CHECK gv_sd_ok = abap_true.

  DATA(lo_sdtab)  = CAST cl_abap_tabledescr(
                      cl_abap_typedescr=>describe_by_data( <gt_sd> ) ).
  DATA(lo_sdline) = CAST cl_abap_structdescr( lo_sdtab->get_table_line_type( ) ).
  PERFORM comp_flat USING lo_sdline '' CHANGING lt_src.

  LOOP AT lt_src INTO DATA(ls_src).
    CHECK ls_src-type->kind = cl_abap_typedescr=>kind_elem.
    READ TABLE lt_comp TRANSPORTING NO FIELDS WITH KEY name = ls_src-name.
    CHECK sy-subrc <> 0.               " nome repetido: fica so o primeiro

    CLEAR ls_comp.
    ls_comp-name = ls_src-name.
    ls_comp-type = ls_src-type.
    APPEND ls_comp TO lt_comp.
    APPEND VALUE ty_map( name_src = ls_src-name
                         name_dst = ls_src-name ) TO gt_map_sd.
  ENDLOOP.
  IF lt_comp IS INITIAL.
    RETURN.
  ENDIF.

  DATA: lo_struct TYPE REF TO cl_abap_structdescr,
        lo_table  TYPE REF TO cl_abap_tabledescr.
  TRY.
      lo_struct = cl_abap_structdescr=>create( lt_comp ).
      lo_table  = cl_abap_tabledescr=>create( p_line_type = lo_struct ).
    CATCH cx_sy_struct_creation cx_sy_table_creation INTO DATA(lx_cr).
      IF p_bg = abap_true.
        gv_log = |Falha ao montar o bloco de vendas: { lx_cr->get_text( ) }|.
        PERFORM log USING gv_log.
      ELSE.
        MESSAGE lx_cr->get_text( ) TYPE 'S' DISPLAY LIKE 'W'.
      ENDIF.
      RETURN.
  ENDTRY.

  CREATE DATA gr_sd_pln TYPE HANDLE lo_table.
  ASSIGN gr_sd_pln->* TO <gt_sdp>.

  DATA lr_lin TYPE REF TO data.
  FIELD-SYMBOLS: <ls_sd>  TYPE any,
                 <ls_pln> TYPE any.
  CREATE DATA lr_lin TYPE HANDLE lo_struct.
  ASSIGN lr_lin->* TO <ls_pln>.

  LOOP AT <gt_sd> ASSIGNING <ls_sd>.
    CLEAR <ls_pln>.
    PERFORM copiar_map USING <ls_sd> gt_map_sd CHANGING <ls_pln>.
    APPEND <ls_pln> TO <gt_sdp>.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& COMP_FLAT  (lista plana dos componentes de uma estrutura)
*&
*& GET_COMPONENTS devolve um INCLUDE TYPE como UM componente estruturado
*& (AS_INCLUDE = 'X'); aqui ele e aberto recursivamente, respeitando o
*& sufixo do RENAMING WITH SUFFIX, para o de-para bater com os nomes que
*& o ASSIGN COMPONENT enxerga.
*&---------------------------------------------------------------------*
FORM comp_flat USING io_st  TYPE REF TO cl_abap_structdescr
                     iv_suf TYPE clike
               CHANGING ct_comp TYPE cl_abap_structdescr=>component_table.

  DATA: lt_sub  TYPE cl_abap_structdescr=>component_table,
        lo_sub  TYPE REF TO cl_abap_structdescr,
        lv_suf2 TYPE string.

  DATA(lt_c) = io_st->get_components( ).

  LOOP AT lt_c INTO DATA(ls_c).
    IF ls_c-as_include = abap_true AND ls_c-type->kind = cl_abap_typedescr=>kind_struct.
      CLEAR lt_sub.
      lo_sub  = CAST cl_abap_structdescr( ls_c-type ).
      lv_suf2 = |{ ls_c-suffix }{ iv_suf }|.
      PERFORM comp_flat USING lo_sub lv_suf2 CHANGING lt_sub.
      APPEND LINES OF lt_sub TO ct_comp.
    ELSE.
      ls_c-name = |{ ls_c-name }{ iv_suf }|.
      CLEAR: ls_c-as_include, ls_c-suffix.
      APPEND ls_c TO ct_comp.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& INDEXAR_VENDAS
*&
*& Le o CNPJ / CPF do cliente de cada linha da ZSD034 e monta o indice
*& documento -> linha. O parceiro usado depende do modo escolhido:
*& emissor da ordem (KUNNR2 / AG), recebedor (KUNNR / WE) ou os dois.
*&---------------------------------------------------------------------*
FORM indexar_vendas CHANGING ct_key TYPE tt_key.

  TYPES: BEGIN OF ty_kun,
           kunnr TYPE kunnr,
         END OF ty_kun,
         BEGIN OF ty_kdoc,
           kunnr TYPE kunnr,
           doc1  TYPE c LENGTH 20,     " CNPJ normalizado
           doc2  TYPE c LENGTH 20,     " CPF normalizado
         END OF ty_kdoc.

  DATA: lt_kun  TYPE STANDARD TABLE OF ty_kun,
        lt_kdoc TYPE SORTED TABLE OF ty_kdoc WITH UNIQUE KEY kunnr,
        lv_idx  TYPE i,
        lv_ag   TYPE kunnr,
        lv_we   TYPE kunnr.

  FIELD-SYMBOLS: <ls_sd> TYPE any,
                 <lv_k>  TYPE any.

  CLEAR ct_key.
  DATA(lv_usa_ag) = COND abap_bool( WHEN p_mag = abap_true OR p_mamb = abap_true
                                    THEN abap_true ELSE abap_false ).
  DATA(lv_usa_we) = COND abap_bool( WHEN p_mwe = abap_true OR p_mamb = abap_true
                                    THEN abap_true ELSE abap_false ).

* 1) clientes citados na saida da ZSD034
  LOOP AT <gt_sd> ASSIGNING <ls_sd>.
    IF lv_usa_ag = abap_true.
      ASSIGN COMPONENT 'KUNNR2' OF STRUCTURE <ls_sd> TO <lv_k>.
      IF sy-subrc = 0 AND <lv_k> IS NOT INITIAL.
        APPEND VALUE ty_kun( kunnr = <lv_k> ) TO lt_kun.
      ENDIF.
    ENDIF.
    IF lv_usa_we = abap_true.
      ASSIGN COMPONENT 'KUNNR' OF STRUCTURE <ls_sd> TO <lv_k>.
      IF sy-subrc = 0 AND <lv_k> IS NOT INITIAL.
        APPEND VALUE ty_kun( kunnr = <lv_k> ) TO lt_kun.
      ENDIF.
    ENDIF.
  ENDLOOP.
  SORT lt_kun BY kunnr.
  DELETE ADJACENT DUPLICATES FROM lt_kun COMPARING kunnr.
  CHECK lt_kun IS NOT INITIAL.

  SELECT kunnr, stcd1, stcd2 FROM kna1
         INTO TABLE @DATA(lt_kna1)
         FOR ALL ENTRIES IN @lt_kun
         WHERE kunnr = @lt_kun-kunnr.
  CHECK sy-subrc = 0.

* 2) documento normalizado por cliente (calculado uma vez so)
  LOOP AT lt_kna1 INTO DATA(ls_kna1).
    DATA(ls_kdoc) = VALUE ty_kdoc( kunnr = ls_kna1-kunnr ).
    PERFORM norm_doc USING ls_kna1-stcd1 'J' CHANGING ls_kdoc-doc1.
    PERFORM norm_doc USING ls_kna1-stcd2 'F' CHANGING ls_kdoc-doc2.
    INSERT ls_kdoc INTO TABLE lt_kdoc.
  ENDLOOP.

* 3) indice documento -> linha da ZSD034
  LOOP AT <gt_sd> ASSIGNING <ls_sd>.
    lv_idx = sy-tabix.
    CLEAR: lv_ag, lv_we.

    IF lv_usa_ag = abap_true.
      ASSIGN COMPONENT 'KUNNR2' OF STRUCTURE <ls_sd> TO <lv_k>.
      IF sy-subrc = 0.
        lv_ag = <lv_k>.
      ENDIF.
    ENDIF.
    IF lv_usa_we = abap_true.
      ASSIGN COMPONENT 'KUNNR' OF STRUCTURE <ls_sd> TO <lv_k>.
      IF sy-subrc = 0.
        lv_we = <lv_k>.
      ENDIF.
    ENDIF.
*   no modo "qualquer um dos dois" emissor e recebedor costumam ser o
*   MESMO parceiro; sem isto a linha entraria duas vezes no cruzamento
    IF lv_we = lv_ag.
      CLEAR lv_we.
    ENDIF.

    DO 2 TIMES.
      DATA(lv_kun) = COND kunnr( WHEN sy-index = 1 THEN lv_ag ELSE lv_we ).
      CHECK lv_kun IS NOT INITIAL.
      READ TABLE lt_kdoc INTO DATA(ls_kd) WITH TABLE KEY kunnr = lv_kun.
      CHECK sy-subrc = 0.
      IF ls_kd-doc1 IS NOT INITIAL.
        INSERT VALUE ty_key( doc = ls_kd-doc1 idx = lv_idx ) INTO TABLE ct_key.
      ENDIF.
      IF ls_kd-doc2 IS NOT INITIAL.
        INSERT VALUE ty_key( doc = ls_kd-doc2 idx = lv_idx ) INTO TABLE ct_key.
      ENDIF.
    ENDDO.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& COPIAR_MAP  (copia campo a campo conforme o de-para informado)
*&---------------------------------------------------------------------*
FORM copiar_map USING is_sd  TYPE any
                      it_map TYPE tt_map
                CHANGING cs_out TYPE any.

  FIELD-SYMBOLS: <lv_s> TYPE any,
                 <lv_d> TYPE any.

  LOOP AT it_map INTO DATA(ls_map).
    ASSIGN COMPONENT ls_map-name_src OF STRUCTURE is_sd  TO <lv_s>.
    CHECK sy-subrc = 0.
    ASSIGN COMPONENT ls_map-name_dst OF STRUCTURE cs_out TO <lv_d>.
    CHECK sy-subrc = 0.
    <lv_d> = <lv_s>.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& EXIBIR_ALV  (modo CRUZADO - uma saida)
*&---------------------------------------------------------------------*
FORM exibir_alv.

  DATA: lo_cols TYPE REF TO cl_salv_columns_table,
        lo_agg  TYPE REF TO cl_salv_aggregations.

  TRY.
      cl_salv_table=>factory(
        IMPORTING r_salv_table = gr_salv
        CHANGING  t_table      = <gt_out> ).

      gr_salv->get_functions( )->set_all( abap_true ).
      gr_salv->get_selections( )->set_selection_mode(
        if_salv_c_selection_mode=>row_column ).

      lo_cols = gr_salv->get_columns( ).
      lo_cols->set_optimize( abap_true ).
      PERFORM titular_colunas USING lo_cols.

      lo_agg = gr_salv->get_aggregations( ).
      PERFORM totalizar USING lo_agg.

      DATA(lo_lay) = gr_salv->get_layout( ).
      lo_lay->set_key( VALUE #( report = sy-repid ) ).
      lo_lay->set_save_restriction( if_salv_c_layout=>restrict_none ).

      gr_salv->get_display_settings( )->set_list_header(
        |ZFI051 - Contratos de compra { p_bukrs } x vendas (ZSD034)| ).

      gr_salv->display( ).

    CATCH cx_salv_msg INTO DATA(lx).
      MESSAGE lx->get_text( ) TYPE 'E'.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& EXIBIR_ALV_SEPARADO  (modo SEPARADO - dois ALVs)
*&
*& Desenha os dois ALVs empilhados num splitter sobre um docking preso
*& a tela de selecao (dynpro 1000): em cima os contratos de compra, em
*& baixo a saida da ZSD034. Cada grid tem sua barra de funcoes, seu
*& layout e seu export proprios - e nenhum dos dois conhece o outro.
*& Chave de conferencia manual entre eles: o CNPJ / CPF do parceiro.
*&
*& Sem cruzamento nao ha coluna CRUZAMENTO: ela e marcada como tecnica.
*&---------------------------------------------------------------------*
FORM exibir_alv_separado.

  DATA: lo_c1 TYPE REF TO cl_gui_container,
        lo_c2 TYPE REF TO cl_gui_container.

  CREATE OBJECT go_dock
    EXPORTING
      repid                       = sy-repid
      dynnr                       = '1000'
      side                        = cl_gui_docking_container=>dock_at_left
      ratio                       = 95
    EXCEPTIONS
      cntl_error                  = 1
      cntl_system_error           = 2
      create_error                = 3
      lifetime_error              = 4
      lifetime_dynpro_dynpro_link = 5
      OTHERS                      = 6.
  IF sy-subrc <> 0.
*   sem docking nao da para empilhar os dois grids. Em vez de nao exibir
*   nada, cai para o modo sequencial: primeiro os contratos em tela
*   cheia; ao sair dele, as vendas.
    CLEAR go_dock.
    PERFORM log USING 'Nao foi possivel abrir a tela dividida - os dois ALVs sao exibidos em sequencia.'.
    PERFORM mostrar_log.
    PERFORM exibir_alv_sequencial.
    RETURN.
  ENDIF.

  CREATE OBJECT go_split
    EXPORTING
      parent  = go_dock
      rows    = 2
      columns = 1.

  lo_c1 = go_split->get_container( row = 1 column = 1 ).
  lo_c2 = go_split->get_container( row = 2 column = 1 ).

* ALV de cima - contratos de compra
  TRY.
      cl_salv_table=>factory(
        EXPORTING r_container  = lo_c1
        IMPORTING r_salv_table = gr_salv_c
        CHANGING  t_table      = gt_ctr ).

      gr_salv_c->get_functions( )->set_all( abap_true ).
      gr_salv_c->get_selections( )->set_selection_mode(
        if_salv_c_selection_mode=>row_column ).

      DATA(lo_cols_c) = gr_salv_c->get_columns( ).
      lo_cols_c->set_optimize( abap_true ).
      PERFORM titular_colunas USING lo_cols_c.
      PERFORM ocultar_cruzamento USING lo_cols_c.
      DATA(lo_agg_c) = gr_salv_c->get_aggregations( ).
      PERFORM totalizar USING lo_agg_c.

      DATA(lo_lay_c) = gr_salv_c->get_layout( ).
      lo_lay_c->set_key( VALUE #( report = sy-repid handle = 'CTR' ) ).
      lo_lay_c->set_save_restriction( if_salv_c_layout=>restrict_none ).

      gr_salv_c->get_display_settings( )->set_list_header(
        |ZFI051 - Contratos de compra { p_bukrs } ({ lines( gt_ctr ) })| ).

      gr_salv_c->display( ).

    CATCH cx_salv_msg INTO DATA(lx1).
      MESSAGE lx1->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
  ENDTRY.

* ALV de baixo - saida da ZSD034
  IF <gt_sdp> IS NOT ASSIGNED.
    MESSAGE 'Sem dados da ZSD034 para os filtros de venda informados.'
            TYPE 'S' DISPLAY LIKE 'W'.
    RETURN.
  ENDIF.

  TRY.
      cl_salv_table=>factory(
        EXPORTING r_container  = lo_c2
        IMPORTING r_salv_table = gr_salv_v
        CHANGING  t_table      = <gt_sdp> ).

      gr_salv_v->get_functions( )->set_all( abap_true ).
      gr_salv_v->get_selections( )->set_selection_mode(
        if_salv_c_selection_mode=>row_column ).

      DATA(lo_cols_v) = gr_salv_v->get_columns( ).
      lo_cols_v->set_optimize( abap_true ).
      PERFORM titular_vendas USING lo_cols_v.
      DATA(lo_agg_v) = gr_salv_v->get_aggregations( ).
      PERFORM totalizar_vendas USING lo_agg_v.

      DATA(lo_lay_v) = gr_salv_v->get_layout( ).
      lo_lay_v->set_key( VALUE #( report = sy-repid handle = 'SD34' ) ).
      lo_lay_v->set_save_restriction( if_salv_c_layout=>restrict_none ).

      gr_salv_v->get_display_settings( )->set_list_header(
        |ZFI051 - Vendas ZSD034 ({ lines( <gt_sdp> ) })| ).

      gr_salv_v->display( ).

    CATCH cx_salv_msg INTO DATA(lx2).
      MESSAGE lx2->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& EXIBIR_ALV_SEQUENCIAL
*&
*& Plano B do modo separado, quando o docking nao pode ser criado: os
*& dois ALVs em tela cheia, um depois do outro. Sai do primeiro (F3) e o
*& segundo aparece.
*&---------------------------------------------------------------------*
FORM exibir_alv_sequencial.

  DATA lo_alv TYPE REF TO cl_salv_table.

  IF gt_ctr IS NOT INITIAL.
    TRY.
        cl_salv_table=>factory(
          IMPORTING r_salv_table = lo_alv
          CHANGING  t_table      = gt_ctr ).
        lo_alv->get_functions( )->set_all( abap_true ).
        DATA(lo_cols_c) = lo_alv->get_columns( ).
        lo_cols_c->set_optimize( abap_true ).
        PERFORM titular_colunas USING lo_cols_c.
        PERFORM ocultar_cruzamento USING lo_cols_c.
        DATA(lo_agg_c) = lo_alv->get_aggregations( ).
        PERFORM totalizar USING lo_agg_c.
        lo_alv->get_display_settings( )->set_list_header(
          |ZFI051 - Contratos de compra { p_bukrs } (1 de 2)| ).
        lo_alv->display( ).
      CATCH cx_salv_msg INTO DATA(lx1).
        MESSAGE lx1->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
    ENDTRY.
  ENDIF.

  CHECK <gt_sdp> IS ASSIGNED.

  TRY.
      cl_salv_table=>factory(
        IMPORTING r_salv_table = lo_alv
        CHANGING  t_table      = <gt_sdp> ).
      lo_alv->get_functions( )->set_all( abap_true ).
      DATA(lo_cols_v) = lo_alv->get_columns( ).
      lo_cols_v->set_optimize( abap_true ).
      PERFORM titular_vendas USING lo_cols_v.
      DATA(lo_agg_v) = lo_alv->get_aggregations( ).
      PERFORM totalizar_vendas USING lo_agg_v.
      lo_alv->get_display_settings( )->set_list_header(
        |ZFI051 - Vendas ZSD034 (2 de 2)| ).
      lo_alv->display( ).
    CATCH cx_salv_msg INTO DATA(lx2).
      MESSAGE lx2->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& OCULTAR_CRUZAMENTO
*&
*& No modo separado nao existe cruzamento, entao a coluna some da saida
*& (tecnica = nem aparece no seletor de layout).
*&---------------------------------------------------------------------*
FORM ocultar_cruzamento USING io_cols TYPE REF TO cl_salv_columns_table.
  TRY.
      io_cols->get_column( 'CRUZAMENTO' )->set_technical( abap_true ).
    CATCH cx_salv_not_found.
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& TITULAR_COLUNAS
*&
*& Bloco do contrato: os mesmos titulos do cockpit, com o volume e o
*& valor RENOMEADOS para "Volume / Valor Contrato Compra".
*& Bloco da ZSD034: o titulo do fieldcat da ZSDR006_NEW, prefixado por
*& "ZSD034 -" - Centro, Ano Safra e Material existem dos dois lados e
*& sem isso ficariam ambiguos. (No modo separado nao ha colunas SD_ nesta
*& tabela e o segundo laco simplesmente nao acha nada.)
*&---------------------------------------------------------------------*
FORM titular_colunas USING io_cols TYPE REF TO cl_salv_columns_table.

  PERFORM col_txt USING io_cols 'EMPRESA'        'Empresa'   'Empresa'        'Empresa do contrato'.
  PERFORM col_txt USING io_cols 'CENTRO'         'Centro'    'Centro'         'Centro do contrato'.
  PERFORM col_txt USING io_cols 'REGIAO'         'Regiao'    'Regiao'         'Regiao (UF) do centro'.
  PERFORM col_txt USING io_cols 'NR_CONTRATO'    'Contrato'  'Contrato'       'Nr. contrato'.
  PERFORM col_txt USING io_cols 'CODFORNECEDOR'  'Fornec.'   'Fornecedor'     'Cod. fornecedor'.
  PERFORM col_txt USING io_cols 'NOMEFORNECEDOR' 'Nome For'  'Nome fornec.'   'Nome do fornecedor'.
  PERFORM col_txt USING io_cols 'CNPJ'           'CNPJ'      'CNPJ'           'CNPJ do fornecedor'.
  PERFORM col_txt USING io_cols 'CPF'            'CPF'       'CPF'            'CPF do fornecedor'.
  PERFORM col_txt USING io_cols 'INSC_ESTADUAL'  'Insc.Est'  'Insc.Estadual'  'Inscricao Estadual'.
  PERFORM col_txt USING io_cols 'ANOSAFRA'       'Safra'     'Safra'          'Ano-safra do contrato'.
  PERFORM col_txt USING io_cols 'DT_PAGAMENTO'   'Dt Pgto'   'Dt Pagamento'   'Data do pagamento'.
  PERFORM col_txt USING io_cols 'VOL_CTR_COMPRA' 'Vol Compr' 'Vol Ctr Compra' 'Volume Contrato Compra'.
  PERFORM col_txt USING io_cols 'VLR_CTR_COMPRA' 'Vlr Compr' 'Vlr Ctr Compra' 'Valor Contrato Compra'.
  PERFORM col_txt USING io_cols 'SITUACAO'       'Situacao'  'Situacao'       'Situacao do contrato'.
  PERFORM col_txt USING io_cols 'CRUZAMENTO'     'Cruzam.'   'Cruzamento'     'Cruzamento compra x venda'.

* colunas herdadas da ZSD034
  DATA: lv_nome  TYPE lvc_fname,
        lv_campo TYPE abap_compname,
        lv_lbl   TYPE c LENGTH 40.

  DATA(lt_cols) = io_cols->get( ).
  LOOP AT lt_cols INTO DATA(ls_col).
    lv_nome = ls_col-columnname.
    CHECK lv_nome(3) = gc_pref.
    CHECK ls_col-r_column IS BOUND.

    lv_campo = lv_nome+3.
    READ TABLE gt_lbl INTO DATA(ls_lbl) WITH TABLE KEY campo = lv_campo.
    IF sy-subrc = 0.
      lv_lbl = ls_lbl-texto.
    ELSE.
      lv_lbl = lv_campo.               " coluna nova da ZSD034: nome tecnico
    ENDIF.

    ls_col-r_column->set_short_text( CONV scrtext_s( lv_lbl ) ).
    ls_col-r_column->set_medium_text( CONV scrtext_m( |Z34 { lv_lbl }| ) ).
    ls_col-r_column->set_long_text( CONV scrtext_l( |{ gc_tit_sd }{ lv_lbl }| ) ).
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& TITULAR_VENDAS
*&
*& ALV / aba de vendas do modo separado: como a tabela e a copia plana
*& da ZSD034, os campos mantem o nome ORIGINAL (sem o prefixo SD_) e o
*& titulo vem direto da lista de labels, sem o prefixo "ZSD034 -" - o
*& ALV inteiro ja e a ZSD034, nao ha com o que confundir.
*&---------------------------------------------------------------------*
FORM titular_vendas USING io_cols TYPE REF TO cl_salv_columns_table.

  DATA: lv_campo TYPE abap_compname,
        lv_lbl   TYPE c LENGTH 40.

  DATA(lt_cols) = io_cols->get( ).
  LOOP AT lt_cols INTO DATA(ls_col).
    CHECK ls_col-r_column IS BOUND.

    lv_campo = ls_col-columnname.
    READ TABLE gt_lbl INTO DATA(ls_lbl) WITH TABLE KEY campo = lv_campo.
    IF sy-subrc = 0.
      lv_lbl = ls_lbl-texto.
    ELSE.
      lv_lbl = lv_campo.               " coluna nova da ZSD034: nome tecnico
    ENDIF.

    ls_col-r_column->set_short_text( CONV scrtext_s( lv_lbl ) ).
    ls_col-r_column->set_medium_text( CONV scrtext_m( lv_lbl ) ).
    ls_col-r_column->set_long_text( CONV scrtext_l( lv_lbl ) ).
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& COL_TXT  (titulos curto / medio / longo de uma coluna)
*&---------------------------------------------------------------------*
FORM col_txt USING io_cols TYPE REF TO cl_salv_columns_table
                   iv_col  TYPE lvc_fname
                   iv_s    TYPE scrtext_s
                   iv_m    TYPE scrtext_m
                   iv_l    TYPE scrtext_l.
  TRY.
      DATA(lo_c) = io_cols->get_column( iv_col ).
      lo_c->set_short_text( iv_s ).
      lo_c->set_medium_text( iv_m ).
      lo_c->set_long_text( iv_l ).
    CATCH cx_salv_not_found.
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& TOTALIZAR  (somatorio das colunas de valor / volume)
*&---------------------------------------------------------------------*
FORM totalizar USING io_agg TYPE REF TO cl_salv_aggregations.

  DATA lt_tot TYPE STANDARD TABLE OF lvc_fname.

  APPEND 'VOL_CTR_COMPRA' TO lt_tot.
  APPEND 'VLR_CTR_COMPRA' TO lt_tot.
  APPEND 'SD_KWMENG'      TO lt_tot.       " Quantidade Pendente (ZSD034)
  APPEND 'SD_KWERT'       TO lt_tot.       " Valor Total (ZSD034)
  APPEND 'SD_QTD_OV'      TO lt_tot.       " Quantidade OV (ZSD034)

  PERFORM add_agg USING io_agg lt_tot.

ENDFORM.

*&---------------------------------------------------------------------*
*& TOTALIZAR_VENDAS  (mesmas colunas de valor, agora sem o prefixo SD_)
*&---------------------------------------------------------------------*
FORM totalizar_vendas USING io_agg TYPE REF TO cl_salv_aggregations.

  DATA lt_tot TYPE STANDARD TABLE OF lvc_fname.

  APPEND 'KWMENG' TO lt_tot.               " Quantidade Pendente
  APPEND 'KWERT'  TO lt_tot.               " Valor Total
  APPEND 'QTD_OV' TO lt_tot.               " Quantidade OV
  APPEND 'NTGEW'  TO lt_tot.               " Peso liquido do item

  PERFORM add_agg USING io_agg lt_tot.

ENDFORM.

*&---------------------------------------------------------------------*
*& ADD_AGG  (tenta somar cada coluna da lista; ignora as que nao existem)
*&---------------------------------------------------------------------*
FORM add_agg USING io_agg TYPE REF TO cl_salv_aggregations
                   it_col TYPE STANDARD TABLE.

  FIELD-SYMBOLS <lv_f> TYPE lvc_fname.

  LOOP AT it_col ASSIGNING <lv_f>.
    TRY.
        io_agg->add_aggregation( columnname  = <lv_f>
                                 aggregation = if_salv_c_aggregation=>total ).
      CATCH cx_salv_data_error cx_salv_not_found cx_salv_existing.
    ENDTRY.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& EXPORTAR_XLSX  (modo CRUZADO, Job: gera o .xlsx e grava no servidor)
*&
*& Sem GUI: monta a ALV headless so para tirar dela o fieldcatalog LVC
*& exigido pelo FACTORY_RESULT_DATA_TABLE e grava o binario pelo OPEN
*& DATASET (filesystem do servidor de aplicacao).
*&---------------------------------------------------------------------*
FORM exportar_xlsx.

  DATA: lo_alv    TYPE REF TO cl_salv_table,
        lo_cols   TYPE REF TO cl_salv_columns_table,
        lt_fcat   TYPE lvc_t_fcat,
        lo_result TYPE REF TO cl_salv_ex_result_data_table,
        lv_xstr   TYPE xstring,
        lx        TYPE REF TO cx_root.

  IF p_xfile IS INITIAL.
    PERFORM log USING 'Diretorio ou nome do arquivo nao informados (bloco Arquivo / Job).'.
    RETURN.
  ENDIF.

  TRY.
      cl_salv_table=>factory(
        IMPORTING r_salv_table = lo_alv
        CHANGING  t_table      = <gt_out> ).

      lo_cols = lo_alv->get_columns( ).
      PERFORM titular_colunas USING lo_cols.

      lt_fcat = cl_salv_controller_metadata=>get_lvc_fieldcatalog(
                  r_columns      = lo_cols
                  r_aggregations = lo_alv->get_aggregations( ) ).

      lo_result = cl_salv_ex_util=>factory_result_data_table(
                    r_data         = gr_out
                    t_fieldcatalog = lt_fcat ).

      cl_salv_bs_lex=>export_from_result_data_table(
        EXPORTING
          is_format            = if_salv_bs_lex_format=>mc_format_xlsx
          ir_result_data_table = lo_result
        IMPORTING
          er_result_file       = lv_xstr ).
    CATCH cx_root INTO lx.
        gv_log = |Falha ao gerar o xlsx: { lx->get_text( ) }|.
        PERFORM log USING gv_log.
      RETURN.
  ENDTRY.

  DATA(lv_lin) = lines( <gt_out> ).
  PERFORM gravar_arquivo USING p_xfile lv_xstr lv_lin.

ENDFORM.

*&---------------------------------------------------------------------*
*& EXPORTAR_SEPARADO  (modo SEPARADO, Job)
*&
*& Duas saidas, dois destinos possiveis:
*&   P_X1ARQ - UM arquivo com DUAS ABAS (padrao)
*&   P_X2ARQ - DOIS arquivos, sufixados _COMPRAS e _VENDAS
*& Em qualquer um dos casos o xlsx e montado por GERAR_XLSX (OOXML +
*& CL_ABAP_ZIP), que roda sem GUI - condicao para o Job funcionar no
*& servidor de aplicacao.
*&---------------------------------------------------------------------*
FORM exportar_separado.

  DATA: lt_aba  TYPE tt_aba,
        ls_aba  TYPE ty_aba,
        lt_uma  TYPE tt_aba,
        lv_xstr TYPE xstring,
        lv_arq  TYPE string,
        lv_lin  TYPE i,
        lv_tot  TYPE i.

  IF p_xfile IS INITIAL.
    PERFORM log USING 'Diretorio ou nome do arquivo nao informados (bloco Arquivo / Job).'.
    RETURN.
  ENDIF.

  PERFORM montar_abas CHANGING lt_aba.
  IF lt_aba IS INITIAL.
    PERFORM log USING 'Nada a gravar.'.
    RETURN.
  ENDIF.

  IF p_x2arq = abap_true.
*   um arquivo por bloco
    LOOP AT lt_aba INTO ls_aba.
      CLEAR lt_uma.
      APPEND ls_aba TO lt_uma.

      IF ls_aba-nome = gc_aba-compras.
        PERFORM sufixar_arquivo USING p_xfile gc_suf-compras CHANGING lv_arq.
      ELSE.
        PERFORM sufixar_arquivo USING p_xfile gc_suf-vendas CHANGING lv_arq.
      ENDIF.

      PERFORM gerar_xlsx USING lt_uma CHANGING lv_xstr.
      IF lv_xstr IS INITIAL.
          gv_log = |Falha ao montar o xlsx de { ls_aba-nome }.|.
          PERFORM log USING gv_log.
        CONTINUE.
      ENDIF.

      PERFORM linhas_da_aba USING ls_aba CHANGING lv_lin.
      PERFORM gravar_arquivo USING lv_arq lv_xstr lv_lin.
    ENDLOOP.
  ELSE.
*   um arquivo so, uma aba por bloco
    PERFORM gerar_xlsx USING lt_aba CHANGING lv_xstr.
    IF lv_xstr IS INITIAL.
      PERFORM log USING 'Falha ao montar o xlsx de duas abas.'.
      RETURN.
    ENDIF.

    LOOP AT lt_aba INTO ls_aba.
      PERFORM linhas_da_aba USING ls_aba CHANGING lv_lin.
      lv_tot = lv_tot + lv_lin.
      gv_log = |Aba { ls_aba-nome }: { lv_lin } linha(s).|.
      PERFORM log USING gv_log.
    ENDLOOP.
    PERFORM gravar_arquivo USING p_xfile lv_xstr lv_tot.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& LINHAS_DA_ABA  (quantas linhas de dados a aba tem)
*&---------------------------------------------------------------------*
FORM linhas_da_aba USING is_aba TYPE ty_aba
                   CHANGING cv_n TYPE i.
  FIELD-SYMBOLS <lt> TYPE STANDARD TABLE.
  cv_n = 0.
  CHECK is_aba-dref IS BOUND.
  ASSIGN is_aba-dref->* TO <lt>.
  CHECK <lt> IS ASSIGNED.
  cv_n = lines( <lt> ).
ENDFORM.

*&---------------------------------------------------------------------*
*& MONTAR_ABAS
*&
*& Descreve as abas do arquivo: nome, colunas (com titulo e natureza da
*& celula) e o ponteiro para os dados. Os titulos saem de uma ALV
*& headless, para serem OS MESMOS que o usuario ve na tela - inclusive
*& as colunas novas que a ZSD034 vier a ganhar.
*&---------------------------------------------------------------------*
FORM montar_abas CHANGING ct_aba TYPE tt_aba.

  DATA: lo_alv TYPE REF TO cl_salv_table,
        ls_aba TYPE ty_aba.

  CLEAR ct_aba.

* aba 1 - contratos de compra.
* A aba e gerada SEMPRE, mesmo sem linha nenhuma - so com o cabecalho.
* Omitir a aba vazia era pior: o usuario recebia um arquivo so com a
* ZSD034 e nao tinha como saber se o lado da compra veio zerado ou se o
* programa tinha deixado de gerar.
  IF gt_ctr IS INITIAL.
    PERFORM log USING
      'Contratos de compra: 0 linha(s) para os filtros do bloco Contratos - aba/arquivo so com cabecalho.'.
  ENDIF.
  TRY.
      cl_salv_table=>factory(
        IMPORTING r_salv_table = lo_alv
        CHANGING  t_table      = gt_ctr ).

      DATA(lo_cols_c) = lo_alv->get_columns( ).
      PERFORM titular_colunas USING lo_cols_c.
      PERFORM ocultar_cruzamento USING lo_cols_c.

      CLEAR ls_aba.
      ls_aba-nome = gc_aba-compras.
      GET REFERENCE OF gt_ctr INTO ls_aba-dref.
      PERFORM montar_colx USING gt_ctr lo_cols_c CHANGING ls_aba-cols.
      APPEND ls_aba TO ct_aba.

    CATCH cx_salv_msg INTO DATA(lx1).
      gv_log = |Falha ao preparar a aba de compras: { lx1->get_text( ) }|.
      PERFORM log USING gv_log.
  ENDTRY.

* aba 2 - vendas (ZSD034). Sem a captura nao existe nem a estrutura das
* colunas, entao aqui a aba realmente nao tem como ser montada.
  IF <gt_sdp> IS ASSIGNED.
    TRY.
        cl_salv_table=>factory(
          IMPORTING r_salv_table = lo_alv
          CHANGING  t_table      = <gt_sdp> ).

        DATA(lo_cols_v) = lo_alv->get_columns( ).
        PERFORM titular_vendas USING lo_cols_v.

        CLEAR ls_aba.
        ls_aba-nome = gc_aba-vendas.
        ls_aba-dref = gr_sd_pln.
        PERFORM montar_colx USING <gt_sdp> lo_cols_v CHANGING ls_aba-cols.
        APPEND ls_aba TO ct_aba.

      CATCH cx_salv_msg INTO DATA(lx2).
        gv_log = |Falha ao preparar a aba de vendas: { lx2->get_text( ) }|.
        PERFORM log USING gv_log.
    ENDTRY.
  ELSE.
    PERFORM log USING
      'A ZSD034 nao devolveu estrutura nenhuma - aba/arquivo de vendas nao gerado. Confira o bloco Vendas.'.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& MONTAR_COLX
*&
*& Colunas de uma aba: nome tecnico, titulo (o mesmo da ALV) e a
*& natureza da celula no xlsx. Colunas tecnicas ou invisiveis ficam de
*& fora - e o caso da CRUZAMENTO no modo separado.
*&---------------------------------------------------------------------*
FORM montar_colx USING it_tab  TYPE STANDARD TABLE
                       io_cols TYPE REF TO cl_salv_columns_table
                 CHANGING ct_colx TYPE tt_colx.

  DATA: ls_colx TYPE ty_colx,
        lo_col  TYPE REF TO cl_salv_column,
        lv_tit  TYPE string,
        lv_pula TYPE abap_bool.

  CLEAR ct_colx.

  DATA(lo_tab) = CAST cl_abap_tabledescr(
                   cl_abap_typedescr=>describe_by_data( it_tab ) ).
  DATA(lo_lin) = CAST cl_abap_structdescr( lo_tab->get_table_line_type( ) ).
  DATA(lt_c)   = lo_lin->get_components( ).

  LOOP AT lt_c INTO DATA(ls_c).
    CHECK ls_c-type->kind = cl_abap_typedescr=>kind_elem.

    CLEAR: ls_colx, lv_tit, lv_pula.
    ls_colx-campo = ls_c-name.

    TRY.
        lo_col = io_cols->get_column( CONV lvc_fname( ls_c-name ) ).
        IF lo_col->is_technical( ) = abap_true.
          lv_pula = abap_true.         " escondida na ALV: fora do arquivo tambem
        ELSE.
          lv_tit = lo_col->get_long_text( ).
          IF lv_tit IS INITIAL.
            lv_tit = lo_col->get_medium_text( ).
          ENDIF.
        ENDIF.
      CATCH cx_salv_not_found.
    ENDTRY.
    CHECK lv_pula = abap_false.

    IF lv_tit IS INITIAL.
      lv_tit = ls_c-name.
    ENDIF.
    ls_colx-titulo = lv_tit.

    CASE ls_c-type->type_kind.
      WHEN cl_abap_typedescr=>typekind_date.
        ls_colx-tipo = gc_cel-data.
      WHEN cl_abap_typedescr=>typekind_time.
        ls_colx-tipo = gc_cel-hora.
      WHEN cl_abap_typedescr=>typekind_packed
        OR cl_abap_typedescr=>typekind_int
        OR cl_abap_typedescr=>typekind_int1
        OR cl_abap_typedescr=>typekind_int2
        OR cl_abap_typedescr=>typekind_float
        OR cl_abap_typedescr=>typekind_decfloat16
        OR cl_abap_typedescr=>typekind_decfloat34.
        ls_colx-tipo = gc_cel-numero.
      WHEN OTHERS.
*       NUMC, CHAR, STRING: texto. NUMC como numero perderia os zeros a
*       esquerda que identificam documento, material e parceiro.
        ls_colx-tipo = gc_cel-texto.
    ENDCASE.

    APPEND ls_colx TO ct_colx.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& GERAR_XLSX
*&
*& Monta o pacote OOXML (xlsx) com uma aba por entrada de IT_ABA e
*& devolve o binario. Feito na unha, com CL_ABAP_ZIP, porque:
*&   - o CL_SALV_BS_LEX so gera arquivo de UMA aba;
*&   - OLE / DOI exigiria GUI, e quem grava aqui e o servidor de
*&     aplicacao, em background.
*& Os textos vao como inlineStr (sem sharedStrings.xml), o que deixa o
*& arquivo um pouco maior e o codigo bem menor.
*&---------------------------------------------------------------------*
FORM gerar_xlsx USING it_aba TYPE tt_aba
                CHANGING cv_xstr TYPE xstring.

  DATA: lo_zip  TYPE REF TO cl_abap_zip,
        lv_s    TYPE string,
        lv_x    TYPE xstring,
        lv_i    TYPE i,
        lt_ct   TYPE stringtab,
        lt_wb   TYPE stringtab,
        lt_rels TYPE stringtab.

  CLEAR cv_xstr.
  CHECK it_aba IS NOT INITIAL.

  CREATE OBJECT lo_zip.

* --- [Content_Types].xml -----------------------------------------
  APPEND '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' TO lt_ct.
  APPEND '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' TO lt_ct.
  APPEND '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' TO lt_ct.
  APPEND '<Default Extension="xml" ContentType="application/xml"/>' TO lt_ct.
  APPEND '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' TO lt_ct.
  APPEND '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' TO lt_ct.

* --- xl/workbook.xml + xl/_rels/workbook.xml.rels ----------------
  APPEND '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' TO lt_wb.
  APPEND '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' &&
         ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' TO lt_wb.

  APPEND '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' TO lt_rels.
  APPEND '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' TO lt_rels.

  LOOP AT it_aba INTO DATA(ls_aba).
    lv_i = sy-tabix.

    DATA(lv_nome) = ls_aba-nome.
    PERFORM nome_aba CHANGING lv_nome.
    PERFORM esc_xml CHANGING lv_nome.

    APPEND |<sheet name="{ lv_nome }" sheetId="{ lv_i }" r:id="rId{ lv_i }"/>| TO lt_wb.
    APPEND |<Relationship Id="rId{ lv_i }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"| &&
           | Target="worksheets/sheet{ lv_i }.xml"/>| TO lt_rels.
    APPEND |<Override PartName="/xl/worksheets/sheet{ lv_i }.xml"| &&
           | ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>| TO lt_ct.

*   a planilha em si
    PERFORM sheet_xml USING ls_aba CHANGING lv_s.
    PERFORM str2x USING lv_s CHANGING lv_x.
    lo_zip->add( name = |xl/worksheets/sheet{ lv_i }.xml| content = lv_x ).
    CLEAR: lv_s, lv_x.
  ENDLOOP.

  APPEND '</sheets></workbook>' TO lt_wb.
* o styles.xml entra depois das planilhas para nao brigar pelos rIds
  lv_i = lines( it_aba ) + 1.
  APPEND |<Relationship Id="rId{ lv_i }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"| &&
         | Target="styles.xml"/>| TO lt_rels.
  APPEND '</Relationships>' TO lt_rels.
  APPEND '</Types>' TO lt_ct.

* --- _rels/.rels -------------------------------------------------
  CONCATENATE
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
    ' Target="xl/workbook.xml"/></Relationships>'
    INTO lv_s.
  PERFORM str2x USING lv_s CHANGING lv_x.
  lo_zip->add( name = '_rels/.rels' content = lv_x ).

* --- as partes montadas acima ------------------------------------
  CONCATENATE LINES OF lt_ct INTO lv_s.
  PERFORM str2x USING lv_s CHANGING lv_x.
  lo_zip->add( name = '[Content_Types].xml' content = lv_x ).

  CONCATENATE LINES OF lt_wb INTO lv_s.
  PERFORM str2x USING lv_s CHANGING lv_x.
  lo_zip->add( name = 'xl/workbook.xml' content = lv_x ).

  CONCATENATE LINES OF lt_rels INTO lv_s.
  PERFORM str2x USING lv_s CHANGING lv_x.
  lo_zip->add( name = 'xl/_rels/workbook.xml.rels' content = lv_x ).

* --- xl/styles.xml (normal, cabecalho em negrito e data) ---------
  PERFORM styles_xml CHANGING lv_s.
  PERFORM str2x USING lv_s CHANGING lv_x.
  lo_zip->add( name = 'xl/styles.xml' content = lv_x ).

  cv_xstr = lo_zip->save( ).

ENDFORM.

*&---------------------------------------------------------------------*
*& SHEET_XML  (a planilha de uma aba)
*&
*& Linha 1 = titulos (negrito, congelada e com autofiltro); da linha 2
*& em diante, os dados. Celula vazia nao e escrita - o xlsx nao exige e
*& o arquivo fica bem menor.
*&---------------------------------------------------------------------*
FORM sheet_xml USING is_aba TYPE ty_aba
               CHANGING cv_xml TYPE string.

  DATA: lt_x    TYPE stringtab,
        lt_let  TYPE stringtab,
        lv_let  TYPE string,
        lv_lin  TYPE i,
        lv_ref  TYPE string,
        lv_cel  TYPE string,
        lv_ult  TYPE string,
        lv_col  TYPE i,
        lv_qtd  TYPE i.

  FIELD-SYMBOLS: <lt>  TYPE STANDARD TABLE,
                 <ls>  TYPE any,
                 <lv>  TYPE any.

  CLEAR cv_xml.
  lv_qtd = lines( is_aba-cols ).

* letras das colunas, calculadas uma vez so (o SY-INDEX vai para uma
* variavel antes do PERFORM: COL_LETRA tem laco proprio)
  DO lv_qtd TIMES.
    lv_col = sy-index.
    PERFORM col_letra USING lv_col CHANGING lv_let.
    APPEND lv_let TO lt_let.
  ENDDO.

  APPEND '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' TO lt_x.
  APPEND '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' TO lt_x.
  APPEND '<sheetViews><sheetView workbookViewId="0">' &&
         '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' &&
         '</sheetView></sheetViews>' TO lt_x.
  APPEND '<sheetFormatPr defaultRowHeight="15" defaultColWidth="18"/>' TO lt_x.
  APPEND '<sheetData>' TO lt_x.

  IF lv_qtd = 0.
*   aba sem coluna nenhuma: deixa um aviso legivel em vez de um arquivo
*   que o Excel abre vazio e o usuario acha que quebrou
    APPEND '<row r="1"><c r="A1" t="inlineStr" s="1"><is><t>Sem colunas para exibir</t></is></c></row>' TO lt_x.
    APPEND '</sheetData></worksheet>' TO lt_x.
    CONCATENATE LINES OF lt_x INTO cv_xml.
    RETURN.
  ENDIF.

* --- linha 1: titulos --------------------------------------------
  APPEND '<row r="1">' TO lt_x.
  LOOP AT is_aba-cols INTO DATA(ls_col).
    READ TABLE lt_let INTO lv_let INDEX sy-tabix.
    DATA(lv_tit) = ls_col-titulo.
    PERFORM esc_xml CHANGING lv_tit.
    APPEND |<c r="{ lv_let }1" t="inlineStr" s="{ gc_xf-titulo }"><is><t>{ lv_tit }</t></is></c>| TO lt_x.
  ENDLOOP.
  APPEND '</row>' TO lt_x.

* --- dados -------------------------------------------------------
  IF is_aba-dref IS BOUND.
    ASSIGN is_aba-dref->* TO <lt>.
  ENDIF.

  IF <lt> IS ASSIGNED.
    LOOP AT <lt> ASSIGNING <ls>.
      lv_lin = sy-tabix + 1.
      APPEND |<row r="{ lv_lin }">| TO lt_x.

      LOOP AT is_aba-cols INTO ls_col.
        READ TABLE lt_let INTO lv_let INDEX sy-tabix.
        ASSIGN COMPONENT ls_col-campo OF STRUCTURE <ls> TO <lv>.
        CHECK sy-subrc = 0.

        lv_ref = |{ lv_let }{ lv_lin }|.
        PERFORM celula_xml USING lv_ref ls_col-tipo <lv> CHANGING lv_cel.
        IF lv_cel IS NOT INITIAL.
          APPEND lv_cel TO lt_x.
        ENDIF.
      ENDLOOP.

      APPEND '</row>' TO lt_x.
    ENDLOOP.
  ELSE.
    lv_lin = 1.
  ENDIF.

  APPEND '</sheetData>' TO lt_x.

* autofiltro no cabecalho (so faz sentido com pelo menos uma linha)
  READ TABLE lt_let INTO lv_ult INDEX lv_qtd.
  IF lv_lin < 1.
    lv_lin = 1.
  ENDIF.
  APPEND |<autoFilter ref="A1:{ lv_ult }{ lv_lin }"/>| TO lt_x.
  APPEND '</worksheet>' TO lt_x.

  CONCATENATE LINES OF lt_x INTO cv_xml.

ENDFORM.

*&---------------------------------------------------------------------*
*& CELULA_XML  (uma celula, conforme a natureza da coluna)
*&
*& Texto ...: sai pelo WRITE ... TO, para respeitar o exit de conversao
*&            do dominio (material, fornecedor e cliente saem como o
*&            usuario ve, nao com os zeros a esquerda internos).
*& Numero ..: WRITE ... NO-GROUPING e depois o separador decimal vira
*&            ponto e o sinal vai para a frente - e o que o OOXML aceita.
*& Data ....: numero de serie do Excel (dias desde 30.12.1899) com o
*&            formato de data curto, entao ordena e filtra como data.
*&---------------------------------------------------------------------*
FORM celula_xml USING iv_ref  TYPE string
                      iv_tipo TYPE c
                      iv_val  TYPE any
                CHANGING cv_cel TYPE string.

  DATA: lv_c   TYPE c LENGTH 1000,
        lv_s   TYPE string,
        lv_d   TYPE d,
        lv_ser TYPE i.

  CONSTANTS lc_base TYPE d VALUE '18991230'.   " serial 0 do Excel

  CLEAR cv_cel.

  CASE iv_tipo.

    WHEN gc_cel-numero.
*     zero e escrito como zero, nao como celula vazia: na ALV o usuario
*     ve 0,00 e no arquivo tem de ver a mesma coisa - branco se leria
*     como "nao informado"
      WRITE iv_val TO lv_c NO-GROUPING LEFT-JUSTIFIED.
      lv_s = lv_c.
      CONDENSE lv_s NO-GAPS.
*     separador decimal do usuario -> ponto; sinal a esquerda
      REPLACE ALL OCCURRENCES OF ',' IN lv_s WITH '.'.
      IF lv_s CS '-'.
        REPLACE ALL OCCURRENCES OF '-' IN lv_s WITH ''.
        lv_s = |-{ lv_s }|.
      ENDIF.
      IF lv_s IS INITIAL OR lv_s CA '*'.
        RETURN.
      ENDIF.
      cv_cel = |<c r="{ iv_ref }"><v>{ lv_s }</v></c>|.

    WHEN gc_cel-data.
      lv_d = iv_val.
      IF lv_d IS INITIAL.
        RETURN.
      ENDIF.
      lv_ser = lv_d - lc_base.
      IF lv_ser <= 0.
        RETURN.                        " data anterior a 1900: sai vazia
      ENDIF.
      cv_cel = |<c r="{ iv_ref }" s="{ gc_xf-data }"><v>{ lv_ser }</v></c>|.

    WHEN OTHERS.                       " texto e hora
      IF iv_val IS INITIAL.
        RETURN.
      ENDIF.
      WRITE iv_val TO lv_c LEFT-JUSTIFIED.
      lv_s = lv_c.                     " C -> STRING ja tira o branco a direita
      SHIFT lv_s LEFT DELETING LEADING space.
      IF lv_s IS INITIAL.
        RETURN.
      ENDIF.
      PERFORM esc_xml CHANGING lv_s.
      cv_cel = |<c r="{ iv_ref }" t="inlineStr"><is><t>{ lv_s }</t></is></c>|.

  ENDCASE.

ENDFORM.

*&---------------------------------------------------------------------*
*& STYLES_XML
*&
*& Tres formatos, na ordem em que GC_XF os enumera: 0 normal, 1 titulo
*& (negrito sobre fundo azul claro) e 2 data curta (numFmtId 14, que o
*& Excel exibe conforme o idioma da maquina).
*&---------------------------------------------------------------------*
FORM styles_xml CHANGING cv_xml TYPE string.

  DATA lt_x TYPE stringtab.

  APPEND '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' TO lt_x.
  APPEND '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' TO lt_x.
  APPEND '<fonts count="2">' &&
         '<font><sz val="11"/><name val="Calibri"/></font>' &&
         '<font><b/><sz val="11"/><name val="Calibri"/></font>' &&
         '</fonts>' TO lt_x.
  APPEND '<fills count="3">' &&
         '<fill><patternFill patternType="none"/></fill>' &&
         '<fill><patternFill patternType="gray125"/></fill>' &&
         '<fill><patternFill patternType="solid">' &&
         '<fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>' &&
         '</fills>' TO lt_x.
  APPEND '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' TO lt_x.
  APPEND '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' TO lt_x.
  APPEND '<cellXfs count="3">' &&
         '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' &&
         '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' &&
         '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' &&
         '</cellXfs>' TO lt_x.
  APPEND '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' TO lt_x.
  APPEND '</styleSheet>' TO lt_x.

  CONCATENATE LINES OF lt_x INTO cv_xml.

ENDFORM.

*&---------------------------------------------------------------------*
*& COL_LETRA  (1 -> A, 27 -> AA ...)
*&---------------------------------------------------------------------*
FORM col_letra USING iv_n TYPE i
               CHANGING cv_l TYPE string.

  CONSTANTS lc_ab TYPE c LENGTH 26 VALUE 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.
  DATA: lv_n TYPE i,
        lv_r TYPE i.

  CLEAR cv_l.
  lv_n = iv_n.
  WHILE lv_n > 0.
    lv_r = ( lv_n - 1 ) MOD 26.
    cv_l = |{ lc_ab+lv_r(1) }{ cv_l }|.
    lv_n = ( lv_n - 1 ) DIV 26.
  ENDWHILE.

ENDFORM.

*&---------------------------------------------------------------------*
*& ESC_XML  (escapa o que nao pode ir cru no XML)
*&---------------------------------------------------------------------*
FORM esc_xml CHANGING cv TYPE string.
  REPLACE ALL OCCURRENCES OF '&' IN cv WITH '&amp;'.    " sempre o primeiro
  REPLACE ALL OCCURRENCES OF '<' IN cv WITH '&lt;'.
  REPLACE ALL OCCURRENCES OF '>' IN cv WITH '&gt;'.
  REPLACE ALL OCCURRENCES OF '"' IN cv WITH '&quot;'.
  REPLACE ALL OCCURRENCES OF REGEX '[[:cntrl:]]' IN cv WITH ''.
ENDFORM.

*&---------------------------------------------------------------------*
*& NOME_ABA  (o Excel recusa > 31 caracteres e os sinais : \ / ? * [ ])
*&---------------------------------------------------------------------*
FORM nome_aba CHANGING cv_nome TYPE string.
  REPLACE ALL OCCURRENCES OF REGEX '[:\\/?*\[\]]' IN cv_nome WITH '-'.
  IF cv_nome IS INITIAL.
    cv_nome = 'Planilha'.
  ENDIF.
  IF strlen( cv_nome ) > 31.
    cv_nome = cv_nome(31).
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& STR2X  (string -> xstring em UTF-8, que e o que o OOXML declara)
*&---------------------------------------------------------------------*
FORM str2x USING iv_s TYPE string
           CHANGING cv_x TYPE xstring.
  DATA lo_conv TYPE REF TO cl_abap_conv_out_ce.
  CLEAR cv_x.
  lo_conv = cl_abap_conv_out_ce=>create( encoding = 'UTF-8' ).
  lo_conv->convert( EXPORTING data = iv_s IMPORTING buffer = cv_x ).
ENDFORM.

*&---------------------------------------------------------------------*
*& SUFIXAR_ARQUIVO  (".../ZFI051.xlsx" + "_COMPRAS" -> ".../ZFI051_COMPRAS.xlsx")
*&---------------------------------------------------------------------*
FORM sufixar_arquivo USING iv_file TYPE string
                           iv_suf  TYPE string
                     CHANGING cv_file TYPE string.

  DATA: lv_i   TYPE i,
        lv_pos TYPE i,
        lv_ch  TYPE c LENGTH 1.

  cv_file = iv_file.
  CHECK cv_file IS NOT INITIAL.

* varre de tras para frente atras da extensao. Se bater numa barra antes
* de achar o ponto, o nome nao tem extensao (o ponto que houver e parte
* do caminho) e o sufixo vai para o fim.
  lv_pos = -1.
  lv_i   = strlen( cv_file ) - 1.
  WHILE lv_i >= 0.
    lv_ch = cv_file+lv_i(1).
    IF lv_ch = '/' OR lv_ch = '\'.
      EXIT.
    ENDIF.
    IF lv_ch = '.'.
      lv_pos = lv_i.
      EXIT.
    ENDIF.
    lv_i = lv_i - 1.
  ENDWHILE.

  IF lv_pos < 0.
    cv_file = |{ cv_file }{ iv_suf }|.
  ELSE.
    cv_file = |{ cv_file(lv_pos) }{ iv_suf }{ cv_file+lv_pos }|.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& GRAVAR_ARQUIVO  (grava o binario no filesystem do servidor)
*&---------------------------------------------------------------------*
FORM gravar_arquivo USING iv_file TYPE string
                          iv_x    TYPE xstring
                          iv_lin  TYPE i.

  DATA: lv_msg  TYPE string,
        lv_bytes TYPE i.

  IF iv_x IS INITIAL.
    gv_log = |Arquivo vazio - nada gravado: { iv_file }|.
    PERFORM log USING gv_log.
    RETURN.
  ENDIF.
  lv_bytes = xstrlen( iv_x ).

  IF p_dpc = abap_true.
    PERFORM gravar_frontend USING iv_file iv_x iv_lin lv_bytes.
    RETURN.
  ENDIF.

* Servidor de aplicacao. O caminho e resolvido pelo SERVIDOR, nao pelo
* PC: uma pasta do Windows do usuario (C:\... ou \\servidor\...) nao
* existe para ele, e e o erro mais comum aqui.
  OPEN DATASET iv_file FOR OUTPUT IN BINARY MODE MESSAGE lv_msg.
  IF sy-subrc <> 0.
    gv_log = |Nao foi possivel gravar no servidor de aplicacao: { iv_file }|.
    PERFORM log USING gv_log.
    IF lv_msg IS NOT INITIAL.
      gv_log = |Motivo do sistema operacional: { lv_msg }|.
      PERFORM log USING gv_log.
    ENDIF.
    PERFORM diagnostico_dir USING iv_file.
    RETURN.
  ENDIF.
  TRANSFER iv_x TO iv_file.
  CLOSE DATASET iv_file.

  DATA(lv_ok) = |Arquivo gerado no servidor: { iv_file } | &&
                |({ iv_lin } linha(s), { lv_bytes } bytes).|.
  PERFORM log USING lv_ok.

ENDFORM.

*&---------------------------------------------------------------------*
*& GRAVAR_FRONTEND  (destino = estacao de trabalho; exige GUI)
*&---------------------------------------------------------------------*
FORM gravar_frontend USING iv_file  TYPE string
                           iv_x     TYPE xstring
                           iv_lin   TYPE i
                           iv_bytes TYPE i.

  DATA: lt_bin TYPE solix_tab,
        lv_len TYPE i.

  CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
    EXPORTING
      buffer        = iv_x
    IMPORTING
      output_length = lv_len
    TABLES
      binary_tab    = lt_bin.

  cl_gui_frontend_services=>gui_download(
    EXPORTING
      bin_filesize = lv_len
      filename     = iv_file
      filetype     = 'BIN'
    CHANGING
      data_tab     = lt_bin
    EXCEPTIONS
      OTHERS       = 1 ).
  IF sy-subrc <> 0.
    gv_log = |Nao foi possivel gravar no PC: { iv_file } (sy-subrc { sy-subrc }).|.
    PERFORM log USING gv_log.
    RETURN.
  ENDIF.

  DATA(lv_ok) = |Arquivo gerado no PC: { iv_file } | &&
                |({ iv_lin } linha(s), { iv_bytes } bytes).|.
  PERFORM log USING lv_ok.

ENDFORM.

*&---------------------------------------------------------------------*
*& DIAGNOSTICO_DIR
*&
*& Quando a gravacao no servidor falha, o motivo quase sempre e o mesmo:
*& o diretorio informado e uma pasta do PC. O F4 navega no frontend, o
*& que convida ao engano. Aqui a mensagem diz isso com todas as letras
*& em vez de deixar o usuario com um "job terminou sem gerar arquivo".
*&---------------------------------------------------------------------*
FORM diagnostico_dir USING iv_file TYPE string.

  DATA: lv_ini TYPE string,
        lv_txt TYPE string.

  IF strlen( iv_file ) >= 2.
    lv_ini = iv_file(2).
  ENDIF.

  IF lv_ini CS ':' OR lv_ini = '\\'.
    lv_txt = 'O caminho informado e uma pasta da ESTACAO DE TRABALHO. O Job roda no '
          && 'SERVIDOR DE APLICACAO, que nao enxerga o disco do seu PC. Informe um '
          && 'diretorio do servidor (dos que a AL11 lista) ou, em primeiro plano, '
          && 'marque o destino "Estacao de trabalho".'.
  ELSE.
    lv_txt = 'Confira na AL11 se o diretorio existe e se o usuario do servidor de '
          && 'aplicacao tem permissao de gravacao nele.'.
  ENDIF.
  PERFORM log USING lv_txt.

ENDFORM.

*&---------------------------------------------------------------------*
*& MONTAR_ARQUIVO  (diretorio + nome, resolvendo &DATA& e &HORA&)
*&---------------------------------------------------------------------*
FORM montar_arquivo CHANGING cv_file TYPE string.

  DATA: lv_nome TYPE string,
        lv_last TYPE c LENGTH 1.

  CLEAR cv_file.
  CHECK p_xdir IS NOT INITIAL.

  lv_nome = p_xname.
  IF lv_nome IS INITIAL.
    lv_nome = 'ZFI051_&DATA&_&HORA&.xlsx'.
  ENDIF.
  REPLACE ALL OCCURRENCES OF '&DATA&' IN lv_nome
          WITH |{ sy-datum(4) }{ sy-datum+4(2) }{ sy-datum+6(2) }|.
  REPLACE ALL OCCURRENCES OF '&HORA&' IN lv_nome
          WITH |{ sy-uzeit(2) }{ sy-uzeit+2(2) }{ sy-uzeit+4(2) }|.

  cv_file = p_xdir.
  lv_last = substring( val = cv_file off = strlen( cv_file ) - 1 len = 1 ).
  IF lv_last <> '/' AND lv_last <> '\'.
    IF cv_file CA '\'.
      cv_file = cv_file && '\'.
    ELSE.
      cv_file = cv_file && '/'.
    ENDIF.
  ENDIF.
  cv_file = cv_file && lv_nome.

ENDFORM.

*&---------------------------------------------------------------------*
*& F4_DIRETORIO
*&
*& Escolhe a pasta no PC do usuario (frontend). Para o servidor de
*& arquivos, navegue ate o compartilhamento de rede (\\SERVIDOR\PASTA\).
*& Lembre que quem grava e o SERVIDOR DE APLICACAO: o caminho precisa
*& existir e ser gravavel por ele (confira na AL11).
*&---------------------------------------------------------------------*
FORM f4_diretorio.

  DATA lv_folder TYPE string.

* O browse so navega no PC. Com destino "servidor" ele daria um caminho
* que o servidor de aplicacao nao enxerga - que era exatamente a
* armadilha que fazia o Job terminar sem gerar arquivo.
  IF p_dsrv = abap_true.
    MESSAGE 'Destino = servidor: informe um diretorio dos que a AL11 lista. O F4 navega no PC.'
            TYPE 'S' DISPLAY LIKE 'W'.
    RETURN.
  ENDIF.

  cl_gui_frontend_services=>directory_browse(
    EXPORTING
      window_title         = 'Pasta onde o Job grava o .xlsx'
      initial_folder       = p_xdir
    CHANGING
      selected_folder      = lv_folder
    EXCEPTIONS
      cntl_error           = 1
      error_no_gui         = 2
      not_supported_by_gui = 3
      OTHERS               = 4 ).
  cl_gui_cfw=>flush( ).
  IF sy-subrc = 0 AND lv_folder IS NOT INITIAL.
    p_xdir = lv_folder.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& AGENDAR_JOB
*&
*& Cria o job de extracao com a selecao ATUAL da tela e a periodicidade
*& informada pelo usuario (P_PINT + a unidade dos radio buttons). O job
*& roda o proprio relatorio com P_BG = 'X': sem ALV, gerando o .xlsx no
*& diretorio informado. P_PINT = 0 agenda uma UNICA execucao.
*& O modo de extracao (cruzado / separado) e a forma do arquivo (uma
*& aba x dois arquivos) tambem viajam na selecao - o Job gera exatamente
*& o que a tela esta pedindo.
*&---------------------------------------------------------------------*
FORM agendar_job.

  DATA: lv_jobcount TYPE tbtcjob-jobcount,
        lv_min      TYPE tbtcjob-prdmins,
        lv_hor      TYPE tbtcjob-prdhours,
        lv_dia      TYPE tbtcjob-prddays,
        lv_sem      TYPE tbtcjob-prdweeks,
        lv_mes      TYPE tbtcjob-prdmonths,
        lv_unid     TYPE string,
        lv_dt       TYPE sy-datum,
        lv_tm       TYPE sy-uzeit.

  IF p_xdir IS INITIAL.
    MESSAGE 'Informe o diretorio do servidor (bloco Arquivo / Job).' TYPE 'E'.
    RETURN.
  ENDIF.

* O Job roda no servidor de aplicacao. Uma pasta do PC ("C:\..." ou
* "\\servidor\...") nao existe para ele - e o job terminaria "Concl."
* sem gerar arquivo nenhum. Barra aqui, na hora de agendar.
  DATA lv_ini TYPE string.
  IF strlen( p_xdir ) >= 2.
    lv_ini = p_xdir(2).
  ENDIF.
  IF lv_ini CS ':' OR lv_ini = '\\'.
    MESSAGE 'Diretorio de PC nao serve para o Job: ele grava no servidor. Informe um caminho da AL11.'
            TYPE 'E'.
    RETURN.
  ENDIF.
  IF p_jname IS INITIAL.
    MESSAGE 'Informe o nome do Job.' TYPE 'E'.
    RETURN.
  ENDIF.
  IF p_pint < 0.
    MESSAGE 'Intervalo de repeticao invalido.' TYPE 'E'.
    RETURN.
  ENDIF.

* periodicidade escolhida pelo usuario
  IF p_umin = abap_true.
    lv_min = p_pint. lv_unid = 'minuto(s)'.
  ELSEIF p_uhor = abap_true.
    lv_hor = p_pint. lv_unid = 'hora(s)'.
  ELSEIF p_udia = abap_true.
    lv_dia = p_pint. lv_unid = 'dia(s)'.
  ELSEIF p_usem = abap_true.
    lv_sem = p_pint. lv_unid = 'semana(s)'.
  ELSEIF p_umes = abap_true.
    lv_mes = p_pint. lv_unid = 'mes(es)'.
  ENDIF.

* 1a execucao: data / hora da tela. Se ja passou, joga para daqui a um
* minuto (o JOB_CLOSE recusa data de inicio no passado).
  lv_dt = p_jdata.
  lv_tm = p_jhora.
  IF lv_dt IS INITIAL.
    lv_dt = sy-datum.
  ENDIF.
  IF lv_dt < sy-datum OR ( lv_dt = sy-datum AND lv_tm <= sy-uzeit ).
    lv_dt = sy-datum.
    lv_tm = sy-uzeit + 60.
    IF lv_tm < sy-uzeit.               " virou a meia-noite
      lv_dt = sy-datum + 1.
    ENDIF.
  ENDIF.

  CALL FUNCTION 'JOB_OPEN'
    EXPORTING
      jobname          = p_jname
    IMPORTING
      jobcount         = lv_jobcount
    EXCEPTIONS
      cant_create_job  = 1
      invalid_job_data = 2
      jobname_missing  = 3
      OTHERS           = 4.
  IF sy-subrc <> 0.
    MESSAGE 'Nao foi possivel abrir o job.' TYPE 'E'.
    RETURN.
  ENDIF.

* mesma selecao da tela; P_BG liga o modo headless + export do xlsx
  SUBMIT (sy-repid)
    WITH p_bukrs  = p_bukrs
    WITH s_contr  IN s_contr
    WITH s_dtctr  IN s_dtctr
    WITH s_dtpag  IN s_dtpag
    WITH s_safra  IN s_safra
    WITH s_forn   IN s_forn
    WITH s_cctr   IN s_cctr
    WITH s_cnpj   IN s_cnpj
    WITH s_cpf    IN s_cpf
    WITH s_ie     IN s_ie
    WITH s_stat   IN s_stat
    WITH s_audat  IN s_audat
    WITH s_vbeln  IN s_vbeln
    WITH s_vkorg  IN s_vkorg
    WITH s_vtweg  IN s_vtweg
    WITH s_spart  IN s_spart
    WITH s_vkbur  IN s_vkbur
    WITH s_vkgrp  IN s_vkgrp
    WITH s_auart  IN s_auart
    WITH s_anosf  IN s_anosf
    WITH s_wsd    IN s_wsd
    WITH s_matnr  IN s_matnr
    WITH s_matkl  IN s_matkl
    WITH p_mcruz  = p_mcruz
    WITH p_msep   = p_msep
    WITH p_x1arq  = p_x1arq
    WITH p_x2arq  = p_x2arq
    WITH p_mag    = p_mag
    WITH p_mwe    = p_mwe
    WITH p_mamb   = p_mamb
    WITH p_semven = p_semven
    WITH p_semctr = p_semctr
    WITH p_xdir   = p_xdir
    WITH p_xname  = p_xname
    WITH p_dsrv   = 'X'
    WITH p_dpc    = ' '
    WITH p_bg     = 'X'
    VIA JOB p_jname NUMBER lv_jobcount
    AND RETURN.

  CALL FUNCTION 'JOB_CLOSE'
    EXPORTING
      jobcount             = lv_jobcount
      jobname              = p_jname
      sdlstrtdt            = lv_dt
      sdlstrttm            = lv_tm
      prdmins              = lv_min
      prdhours             = lv_hor
      prddays              = lv_dia
      prdweeks             = lv_sem
      prdmonths            = lv_mes
    EXCEPTIONS
      cant_start_immediate = 1
      invalid_startdate    = 2
      jobname_missing      = 3
      job_close_failed     = 4
      job_nosteps          = 5
      job_notex            = 6
      lock_failed          = 7
      OTHERS               = 8.
  IF sy-subrc <> 0.
    MESSAGE 'Falha ao agendar o job (JOB_CLOSE).' TYPE 'E'.
    RETURN.
  ENDIF.

  DATA(lv_saida) = COND string(
    WHEN p_msep = abap_false THEN 'um arquivo (compra x venda cruzados)'
    WHEN p_x2arq = abap_true THEN 'dois arquivos (_COMPRAS e _VENDAS)'
    ELSE 'um arquivo com duas abas (compras e vendas)' ).

  IF p_pint > 0.
    MESSAGE |Job '{ p_jname }' agendado para { lv_dt DATE = USER } { lv_tm TIME = USER }, | &&
            |repetindo a cada { p_pint } { lv_unid }. Gera { lv_saida } em { p_xdir }.| TYPE 'S'.
  ELSE.
    MESSAGE |Job '{ p_jname }' agendado para { lv_dt DATE = USER } { lv_tm TIME = USER } | &&
            |(execucao unica). Gera { lv_saida } em { p_xdir }.| TYPE 'S'.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& CANCELAR_JOB  (remove os agendamentos FUTUROS; nao mexe no que ja rodou)
*&---------------------------------------------------------------------*
FORM cancelar_job.

  DATA lv_n TYPE i.

  SELECT * FROM tbtco INTO TABLE @DATA(lt_jobs)
         WHERE jobname = @p_jname
           AND status IN ( 'P', 'S', 'Y' ).   " P=agendado, S=liberado, Y=pronto
  IF lt_jobs IS INITIAL.
    MESSAGE 'Nenhum agendamento ativo do Job encontrado.' TYPE 'S'.
    RETURN.
  ENDIF.

  LOOP AT lt_jobs INTO DATA(ls_job).
    CALL FUNCTION 'BP_JOB_DELETE'
      EXPORTING
        jobcount                 = ls_job-jobcount
        jobname                  = ls_job-jobname
      EXCEPTIONS
        cant_delete_event_entry  = 1
        cant_delete_job          = 2
        cant_delete_joblog       = 3
        cant_delete_steps        = 4
        cant_delete_time_entry   = 5
        cant_derelease_successor = 6
        job_delete_failed        = 7
        job_does_not_exist       = 8
        no_delete_authority      = 9
        OTHERS                   = 10.
    IF sy-subrc = 0.
      lv_n = lv_n + 1.
    ENDIF.
  ENDLOOP.

  MESSAGE |{ lv_n } agendamento(s) do Job removido(s).| TYPE 'S'.

ENDFORM.

*&---------------------------------------------------------------------*
*& SITUACAO_JOB  (quantos agendamentos futuros e qual a proxima execucao)
*&---------------------------------------------------------------------*
FORM situacao_job.

  SELECT jobname, jobcount, status, sdlstrtdt, sdlstrttm
         FROM tbtco INTO TABLE @DATA(lt_j)
         WHERE jobname = @p_jname
           AND status IN ( 'P', 'S', 'Y', 'R' ).
  IF lt_j IS INITIAL.
    MESSAGE |Nenhum agendamento ativo do job '{ p_jname }' (detalhe na SM37).| TYPE 'S'.
    RETURN.
  ENDIF.

  SORT lt_j BY sdlstrtdt ASCENDING sdlstrttm ASCENDING.
  READ TABLE lt_j INTO DATA(ls_j) INDEX 1.

  MESSAGE |Job '{ p_jname }': { lines( lt_j ) } agendamento(s). Proxima execucao | &&
          |{ ls_j-sdlstrtdt DATE = USER } { ls_j-sdlstrttm TIME = USER }. Detalhe na SM37.| TYPE 'S'.

ENDFORM.

*---------------------------------------------------------------------*
* macro do cadastro de titulos (so para a lista abaixo ficar legivel)
*---------------------------------------------------------------------*
DEFINE lbl.
  INSERT VALUE ty_lbl( campo = &1 texto = &2 ) INTO TABLE gt_lbl.
END-OF-DEFINITION.

*&---------------------------------------------------------------------*
*& CARREGAR_LABELS
*&
*& Titulos das colunas da ZSD034 - os MESMOS textos do fieldcat da
*& ZSDR006_NEW (FORM F_FIELDCAT_INIT). Ficam aqui porque a captura em
*& memoria traz os DADOS, nao o fieldcat. Coluna nova que a ZSD034
*& ganhar e que nao estiver nesta lista sai com o nome tecnico do campo
*& (e continua aparecendo no relatorio).
*&---------------------------------------------------------------------*
FORM carregar_labels.

  lbl 'SEMAFORO'   'Semaforo'.
  lbl 'ANOSAFRA'   'Ano Safra'.
  lbl 'BSTDK'      'Data pedido'.
  lbl 'BSTKD'      'Nr. pedido'.
  lbl 'VBELN'      'Documento de vendas'.
  lbl 'VBELV'      'Documento (Precedente)'.
  lbl 'EDATU'      'Data 1a Remessa'.
  lbl 'KUNNR'      'Recebedor mercadoria'.
  lbl 'NAME1'      'Nome cliente'.
  lbl 'STREET_WE'  'Rua Recebedor'.
  lbl 'CITY2_WE'   'Bairro Recebedor'.
  lbl 'CITY1_WE'   'Cidade Recebedor'.
  lbl 'REGIO_WE'   'UF Recebedor'.
  lbl 'LZONE'      'Zona de transporte'.
  lbl 'ORT01'      'Cidade'.
  lbl 'ORT02'      'Bairro'.
  lbl 'REGIO'      'Estado'.
  lbl 'MATNR'      'Material'.
  lbl 'ARKTX'      'Denominacao'.
  lbl 'VKORG'      'Organizacao de Vendas'.
  lbl 'VTEXT'      'Nome Organizacao Vendas'.
  lbl 'VTWEG'      'Canal de distribuicao'.
  lbl 'VTEXT_V'    'Nome Canal distribuicao'.
  lbl 'SPART'      'Setor de Atividade'.
  lbl 'VTEXT_S'    'Nome Setor Atividade'.
  lbl 'VKBUR'      'Escritorio de vendas'.
  lbl 'BEZEI_V'    'Nome Escritorio vendas'.
  lbl 'VKGRP'      'Equipe de Vendas'.
  lbl 'BEZEI'      'Nome Equipe de Vendas'.
  lbl 'PRODH'      'Hierarquia'.
  lbl 'OBTEN'      'Obtentor'.
  lbl 'TECNO'      'Tecnologia'.
  lbl 'CULTI'      'Cultivar'.
  lbl 'BEZEI_2'    'Embalagem'.
  lbl 'LBTXT'      'Laboratorio/Escritorio'.
  lbl 'PSTYV'      'Categoria do item'.
  lbl 'NTGEW'      'Peso liquido do item'.
  lbl 'KWMENG'     'Quantidade Pendente'.
  lbl 'VRKME'      'Unidade de Venda'.
  lbl 'KBETR'      'Valor Unitario'.
  lbl 'KWERT'      'Valor Total'.
  lbl 'VDATU'      'Data Expedicao'.
  lbl 'VSART'      'Tipo de Frete'.
  lbl 'SPSTG_BEZ'  'Status do Pedido'.
  lbl 'NAME2'      'Nome Fantasia'.
  lbl 'TELF1'      '1o Nr. telefone'.
  lbl 'BSARK'      'Tipo de pedido'.
  lbl 'LFSTA'      'Status da Remessa'.
  lbl 'KBETR1'     'ZAT0-Germoplasma Tabela'.
  lbl 'KBETR2'     'ZAV0-Germoplasma Manual'.
  lbl 'KBETR3'     'ZAT6-TSI Tabela'.
  lbl 'KBETR4'     'ZAV6-TSI Manual'.
  lbl 'KBETR5'     'ZAT8-Royalties Tabela'.
  lbl 'KBETR6'     'ZAV8-Royalties Manual'.
  lbl 'KBETR7'     'ZA11-Royalties pago'.
  lbl 'KBETR8'     'ZAP6-Tratamento Pago'.
  lbl 'KBETR9'     'ZAT9'.
  lbl 'KBETR10'    'ZAP0-Germo Pago'.
  lbl 'KBETR11'    'ZA10-Frete'.
  lbl 'KBETR12'    'ZAV3-Var. Comercial'.
  lbl 'KBETR13'    'ZAT3-Variavel Comercial'.
  lbl 'KWERT_ZDED' 'ZDED - Deducao ICMS'.
  lbl 'QTD_OV'     'Quantidade OV'.
  lbl 'KUNNR2'     'Emissor da Ordem'.
  lbl 'NAME3'      'Nome do Emissor'.
  lbl 'ERDAT'      'Data de Criacao'.
  lbl 'VALDT'      'Data de vencimento'.
  lbl 'WERKS'      'Centro'.
  lbl 'INCO1'      'Incoterms'.
  lbl 'FBUDA'      'Data prestacao servico'.
  lbl 'BZIRK'      'Regional de vendas'.
  lbl 'PEDIDO_SF'  'Pedido SF'.

ENDFORM.
