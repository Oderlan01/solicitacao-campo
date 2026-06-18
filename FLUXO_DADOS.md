# Fluxo de Dados e Dicionário de Variáveis — Ecossistema CAD ↔ Excel

Este documento é a **referência canônica** do ecossistema que integra os blocos do AutoCAD
a uma planilha Excel. Ele descreve o fluxo de dados, o dicionário único de variáveis e as
regras de negócio. Todas as peças de código (AutoLISP, Power Query, VBA) devem se alinhar a
este documento — a fonte da verdade das variáveis e dos domínios vive no AutoLISP
(`STW:DOMINIOS` em `stw-automacao.lsp`; dicionário canônico neste documento).

## Arquitetura (híbrida, centrada no LISP)

- **AutoLISP** é dono dos **domínios de validação** (`STW:DOMINIOS`). Exporta os dados dos
  blocos (CSV largo) e as listas de dropdown (`dominios.csv`). O **dicionário canônico** de
  variáveis (atributo CAD ↔ nome amigável ↔ grupo ↔ direção) é **este documento**.
- **Excel/Power Query/VBA** é a camada **interativa** (análise, dropdowns, fórmulas, cores,
  backup, IDs sequenciais). Consome o schema/domínios do LISP em vez de listas hard-coded.

```
                    STWExportar (LISP)             ┌──────────────────────────────────────┐
   ┌──────────────┐ atualiza tags + grava  ┌──────┴────────────┐  Power Query (1 por aba)  │
   │  BLOCO CAD   │ ─────────────────────► │ todos_atributos.csv│ ─► Aba COMPONENTES (núcleo, expande pai/filho)
   │  (AutoCAD)   │                        │ formato LARGO      │ ─► Aba MES3        (0K/0L/0M — etiqueta)
   │  atributos   │ ◄───────────────────── │ 1 linha por HANDLE │ ─► Aba ATRIBUTOS   (EAV Tipo+valor)
   │  ATL_STW_*   │  STWImportar (LISP)     │ cabeçalho dinâmico │                          │
   └──────────────┘                         └──────┬────────────┘ └────────────┬───────────┘
        ▲   │                                      ▲                           │ edição do usuário
        │   └──(STW:DOMINIOS)─► dominios.csv ──► aba oculta "Dominios" ──► dropdowns dependentes
        │                                          │ ExportarParaCAD (VBA)      ▼
        └──────────────────────────────────────────┴──── Componentes (1 linha-pai por HANDLE)
          STWImportar / ImportarCamposManuais (Excel → CAD)
```

Via alternativa de ingestão (legado): AutoCAD **DATAEXTRACTION** ("export dados") gera um
`.xlsx` (aba *Summary*) lido por `power-query-componentes.m`. Substituída pela via CSV direto
(`power-query-csv.m`), mas mantida porque a fonte MES3 ainda usa DATAEXTRACTION.

### Pontos-chave

- **Chave de identidade**: `HANDLE_CAD` (handle do bloco), estável e único por bloco no CSV.
- **CSV largo único**: 1 linha por bloco, 1 coluna por atributo; cabeçalho montado
  dinamicamente pela LISP a partir dos atributos visíveis (suporta atributos novos sem código).
- **Cada grupo lógico vira uma aba/query própria**. A query de Componentes remove de propósito
  as colunas dos outros grupos (`power-query-csv.m:75`) — é roteamento, não descarte.
- **A aba Componentes expande** 1 bloco em várias linhas (sufixos `TAG(1)`/`TAG(A)`, filhos
  laterais). Logo **não é 1:1 com HANDLE**: a gravação de volta usa só a **linha-pai** por
  HANDLE (1ª ocorrência), como o VBA já faz via `handlesVistos`.

## Dicionário Canônico de Variáveis

Dicionário canônico (referência única para LISP/PQ/VBA). Direção: **RW** = round-trip (edita no Excel e
grava no CAD) · **R** = somente leitura/derivado · **CTRL** = interno/identidade.

### Grupo NÚCLEO (Bloco CAD)

| Atributo CAD          | Nome amigável     | Direção | Observação                          |
|-----------------------|-------------------|---------|-------------------------------------|
| `0A_CATEGORIA`        | CATEGORIA         | RW      | dispara domínio de ACIONAMENTO      |
| `0B_SETOR`            | SETOR             | RW      |                                     |
| `0C_FAMILIA`          | FAMILIA           | RW      |                                     |
| `0D_MODELO`           | MODELO            | RW      |                                     |
| `0E_TAG`              | EQUIPAMENTO       | RW      | base da coluna TAG (expansão)       |
| `0F_TAG_AUXILIAR`     | EQUIPAMENTO AUX   | RW      |                                     |
| `0G_DESCRICAO_GERAL`  | DESCRICAO GERAL   | R       | composta no PQ (não regravar)       |
| `0H_ACIONAMENTO`      | ACIONAMENTO       | RW      | dropdown (domínio depende CATEGORIA)|
| `0I_DADOS_ENGENHARIA` | DADOS ENGENHARIA  | R       | split em POTÊNCIA/CORRENTE/TENSÃO   |
| `0J_SENSOR`           | SENSOR            | RW      |                                     |
| `CAIXA_DE_PASSAGEM`   | CAIXA DE PASSAGEM | RW      |                                     |
| `VALVULA_ABRE`        | VALVULA ABRE      | RW      | gera linha filho no PQ              |

### Grupo MES3 (etiqueta — aba/query própria)

| Atributo CAD    | Nome amigável | Direção |
|-----------------|---------------|---------|
| `0K_TAG_SIZE`   | TAG SIZE      | RW      |
| `0L_TAG_AMOUNT` | TAG AMOUNT    | RW      |
| `0M_NAME`       | NAME          | RW      |

Alimenta a aba **DADOS_MES3** (`power-query-mes3.m`) → tabela **MES3** com `id` sequencial.

### Grupo MANUAL (preenchido no Excel)

| Atributo CAD | Nome amigável | Direção | Volta ao CAD                      |
|--------------|---------------|---------|-----------------------------------|
| `INTERFACE`  | INTERFACE     | RW      | `campos_manuais.csv` (RN6)         |
| `CONEXAO`    | CONEXAO       | RW      | dropdown depende de ACIONAMENTO    |
| `ORIGEM`     | ORIGEM        | RW      |                                   |

> Obs.: na tabela analisada `INTERFACE` é o nome amigável; a coluna K/11 dos dropdowns.

### Grupo EAV (atributos genéricos Tipo + valor)

Conjunto **aberto** sem colunas fixas (ex.: `9A_TAG`, `4A_AUX1`, `4B_AUX2`). Como a LISP
exporta qualquer atributo visível dinamicamente, novos tipos viram colunas automaticamente.
O tratamento amigável deve ser data-driven (consultar o schema), nunca hard-coded.

### Grupo CTRL (identidade / interno)

| Atributo CAD    | Papel                                |
|-----------------|--------------------------------------|
| `HANDLE_CAD`    | chave de identidade (coluna do CSV)  |
| `ID_VISIVEL`    | = handle, sincronizado pela LISP     |
| `NOME_DO_BLOCO` | nome efetivo do bloco                |

## Domínios de Validação (RN4)

Espelha `STW:DOMINIOS`; exportado para `dominios.csv` por `STWExportarDominios` e consumido
pelos dropdowns do Excel (aba oculta `Dominios`, ver `Modulo_Dominios.bas`).

- **ACIONAMENTO** depende de **CATEGORIA**:
  - `MOTOR` → Soft Starter, Inversor, Partida Direta, Partida Direta + Inv, Inversor + PD,
    Partida Direta + Rev, Partida Inteligente, Soft Starter + Rev
  - demais (`*`) → Capacitivo, Indutivo, Magnético, Trava Seg., Laser, Temperatura,
    Ultrassônico, Contato, Acionamento, Feedback, Válvula, Simples Solenóide, Dupla Solenóide
- **INTERFACE** depende de **ACIONAMENTO**:
  - Inversor / Soft Starter / Partida Inteligente / Soft Starter + Rev / Inversor + PD /
    Partida Direta + Inv → I/O, Rede
  - Partida Direta / Partida Direta + Rev → I/O
  - Capacitivo / Indutivo / Magnético / Trava Seg. / Laser / Ultrassônico / Contato /
    Feedback → DI, AI, Circuito Elétrico, I/O Link
  - Temperatura → AI, I/O Link, RTD
  - Acionamento / Válvula / Dupla Solenóide / Simples Solenóide → AO, DO, Circuito Elétrico,
    I/O Link

## Regras de Negócio (RN1–RN6)

**RN1 — Identidade e chaves**
- `HANDLE_CAD` identifica o bloco; a LISP sincroniza em `ID_VISIVEL`.
- Componentes: chave de negócio = `TAG` (col H/8) + valor da col B/2; colunas B e H são
  protegidas (nunca sobrescritas no refresh do PQ).
- MES3: chave = `tag`; `id` sequencial incremental persistente.

**RN2 — Sincronização Componentes** (`Modulo1_Sincronizar.bas`)
- Backup físico timestamped antes (`SaveCopyAs` em `/Backup`).
- Memoriza fórmulas e valores (col 2..20) por chave; refresh do PQ; reintegra preservando o
  trabalho manual: fórmula restaurada (ouro), valor restaurado quando PQ veio vazio/`//`
  (ouro), divergência manual (vermelho), linha nova (laranja).
- Dedup: `TAG` repetida (`CountIf>1`) → célula TAG em vermelho.

**RN3 — Sincronização MES3** (`Modulo1_Sincronizar.bas`)
- Fonte `DADOS_MES3` (`power-query-mes3.m`): `0M_NAME→name`, `0L_TAG_AMOUNT→tag_amount`,
  `0K_TAG_SIZE→tag_size`, `0E_TAG→tag`, `0G_DESCRICAO_GERAL→description`; + `id`, `client_id`,
  `code`.
- Destino MES3 (`Tabela3`): `id, name, description, tag, tag_size, tag_amount`. Remove o que
  sumiu da fonte; adiciona novos com `id = max+1` (IDs estáveis).

**RN4 — Domínios de validação**: ver seção acima. Listas vêm do LISP via `dominios.csv`.

**RN5 — Expansão/derivação** (Power Query Componentes)
- Sufixos `TAG(1)`/`TAG(A)` → réplicas pai; colunas laterais → linhas filho; MOTOR primeiro.
- `DADOS ENGENHARIA` `P/C/T` → POTÊNCIA/CORRENTE/TENSÃO; `DESCRICAO GERAL` = `Categoria |
  Família Equipamento`. INTERFACE/CONEXÃO/ORIGEM são colunas manuais (null no PQ).

**RN6 — Volta ao CAD** (AutoLISP)
- `todos_atributos.csv` (largo) → `STWImportar`; ou `campos_manuais.csv`
  (`HANDLE;INTERFACE;CONEXAO;ORIGEM`) → `ImportarCamposManuais` (`automacao.lsp`).
- Só grava se valor ≠ "". Campos `R` (derivados) não devem ser regravados no bloco.

## Arquivos do ecossistema

| Arquivo                       | Papel                                                        |
|-------------------------------|--------------------------------------------------------------|
| `stw-automacao.lsp`           | LISP principal: domínios, export/import (CSV largo), propagação |
| `automacao.lsp`               | LISP legado: `ImportarCamposManuais`, `VerificarAtributos`   |
| `power-query-csv.m`           | PQ aba Componentes (fonte = CSV largo)                       |
| `power-query-componentes.m`   | PQ aba Componentes (fonte = DATAEXTRACTION, legado)          |
| `power-query-mes3.m`          | PQ fonte DADOS_MES3 (etiqueta MES3)                          |
| `Modulo_Sync.bas`             | VBA (reformulado): `AtualizarTudo` (sync rápido/seguro) + `BackupManual` |
| `Modulo_Dominios.bas`         | VBA: importa `dominios.csv` e fornece opções aos dropdowns   |
| `Planilha_Componentes.cls`    | VBA: `Worksheet_SelectionChange` (dropdowns via domínios)    |
| `Modulo3_SincronizarCAD.bas`  | VBA: write-back Excel→CSV por HANDLE (`ExportarParaCAD`, mapa só RW) |

## Importação no workbook (evita travar)

Ao colar/importar estes módulos no `.xlsm`, **substitua** os antigos — não mantenha dois
módulos com o mesmo nome nem duas macros com o mesmo nome (gera "Nome ambíguo detectado",
que trava o projeto inteiro). Em particular, ao adotar `Modulo_Sync.AtualizarTudo`, remova do
`Módulo1` as macros antigas `AtualizarE_Sincronizar_Total2`, `AtualizarMES3_Sequencial` e
`RealizarBackupFisico`.
