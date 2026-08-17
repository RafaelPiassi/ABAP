# Projetos ABAP

Índice dos meus desenvolvimentos em **SAP ABAP** publicados como repositórios independentes. Cada um traz o fonte completo, um README técnico e a documentação em `.docx`/`.pdf`.

Escritos para **SAP ECC**, em português, com ALV via `CL_SALV_TABLE` ou `REUSE_ALV_GRID_DISPLAY`. Os códigos de empresa e centro nos exemplos são genéricos (`1000`, `1000`–`1026`).

**Desenvolvedor:** RPIASSI — Rafael Piassi

---

## Coleções por módulo

Objetos agrupados, um diretório por transação — cada um com fonte, README e `docs/`.

| Repositório | Objetos |
|---|---|
| [**sap-fi-reports**](https://github.com/RafaelPiassi/sap-fi-reports) — FI / CO | ZFI007 · ZFI009N · ZFI009R · ZFI018 · ZFI036 · ZFI037 · ZFI039 · ZFI040 · ZFI044 · ZFI045 · ZFI048 · ZFI053 · ZFI054 · ZFF67 · ZCO01 |
| [**sap-mm-reports**](https://github.com/RafaelPiassi/sap-mm-reports) — MM | ZMM113 · ZMM116 · ZMM118 · ZMM132 · ZMM198 · ZMMR033 · ZMMR117 · ZMIGO_CARGA_PEDIDO · ZQM022 |
| [**sap-sd-reports**](https://github.com/RafaelPiassi/sap-sd-reports) — SD | ZSD058 · ZSD117 · ZSD120 · ZSD126 · ZSD127 · ZSD128 · ZVA03 · ZMOSAIC · Z_NFDANFE_PORTRAIT |

## Repositórios individuais

Objetos maiores, cada um no seu próprio repositório.

### Basis e administração

Contas, jobs, certificados e a rotina diária de AMS — o que mantém o ambiente de pé.

| Objeto | O que faz |
|---|---|
| [ZBC001 — Relação de usuários](https://github.com/RafaelPiassi/sap-zbc001-relacao-usuarios) | Todos os usuários do mandante em ALV: endereço completo, alias de logon, tipo de licença contratual, bloqueio decodificado da `UFLAG` e tipo de usuário. |
| [ZBC002 — Saneamento de usuários](https://github.com/RafaelPiassi/sap-zbc002-saneamento-usuarios) | Validade máxima de 1 ano, sem logon há 90 dias, expirados. Ajuste individual ou em massa por `BAPI_USER_CHANGE`, carga de endereço por planilha e job diário com e-mail. |
| [ZBASISCHECK — Check list diário de Basis](https://github.com/RafaelPiassi/sap-zbasischeck-checklist-basis) | Automatiza os 33 itens do check list diário de AMS: 24 verificações automáticas, diagnóstico em ALV, e-mail HTML/CSV e job diário. |
| [ZJOBS — Monitor e limpeza de jobs](https://github.com/RafaelPiassi/sap-zjobs-monitor-jobs) | A SM37 com ação em massa: cancelar, deletar, iniciar, liberar e reagendar a seleção inteira. Simulação por padrão, proteção dos jobs do SAP, teto por execução e processamento em lotes com `COMMIT` a cada N. |
| [ZNFECERT — Monitor de certificados NF-e](https://github.com/RafaelPiassi/sap-znfecert-monitor-certificados-nfe) | Monitora os certificados digitais da `NFE_MNG_CERT` e avisa por e-mail 15 dias antes do vencimento, com job diário. |

### Localização Brasil

| Objeto | O que faz |
|---|---|
| [ZFI049 — Inscrição Estadual × Centro](https://github.com/RafaelPiassi/sap-zfi049-inscricao-estadual) | Confere e corrige IE, IM e CNPJ dos locais de negócios (`J_1BBRANCH`) de todos os centros de uma empresa, numa tela só. Aponta IE sem zeros à esquerda. |
| [ZFI050 — IE dos parceiros do centro](https://github.com/RafaelPiassi/sap-zfi050-ie-parceiros-centro) | Companheira da ZFI049: corrigir a `J_1BBRANCH` arruma o destinatário da NF-e, mas **não** o local de entrega, que vem de `T001W-KUNNR` → `KNA1-STCD3`. Iguala os dois via `CMD_EI_API`/`VMD_EI_API`. |

### Materiais e estoque

| Objeto | O que faz |
|---|---|
| [ZMM052 — Movimentação com estoque atual](https://github.com/RafaelPiassi/sap-zmm052-movimentacao-estoque) | Uma linha por movimento (estilo MB51) com o estoque atual em colunas e a origem → destino resolvida por centro, local de negócio e depósito. |
| [ZMM119 — Lotes em estoque](https://github.com/RafaelPiassi/sap-zmm119-lotes-em-estoque) | Consolida MB52, MB58 e MB51 numa tela só. Lê **todos** os tipos de estoque da `MCHB` — não só a utilização livre — e expõe a classificação do lote como colunas filtráveis. |
| [ZMM122 — Expansão de cadastros](https://github.com/RafaelPiassi/sap-zmm122-expansao-cadastros) | Estende material, cliente e fornecedor a novas organizações copiando de um registro de referência, com de-para de campos dinâmico. |

### Cadastros e ferramentas

| Objeto | O que faz |
|---|---|
| [ZRHIERARQUIA — Hierarquia de materiais](https://github.com/RafaelPiassi/sap-zrhierarquia-hierarquia-materiais) | Mantém `T179`/`T179T` em ALV editável sobre docking, substituindo a V/76 sem precisar liberar SM30. |
| [ZTABLEVIEW — Visualizador de tabelas](https://github.com/RafaelPiassi/sap-ztableview) | Comportamento da SE16N, porém exclusivamente para exibição — sem qualquer poder de edição. |

---

## Sobre o código

Todos os objetos seguem o mesmo padrão:

- **Cabeçalho documentado** — objetivo, tabelas SAP utilizadas e premissas do ambiente a confirmar antes de usar em produção;
- **README por repositório** — o que faz, por que existe, tela de seleção campo a campo, tabelas envolvidas e passo a passo de instalação (SE38 / SE93 / SE11);
- **Documentação técnica e funcional** em `.docx` e `.pdf`, gerada por um pipeline próprio (JSON → Word → PDF).

Nomes de objetos Z, tabelas customizadas e classes de classificação citadas nos READMEs são do ambiente onde foram desenvolvidos. Onde há dependência de objeto customizado, o README diz explicitamente o que precisa existir — e, quando a leitura é dinâmica, o report funciona sem ele.

## Licença

MIT em todos os repositórios.
