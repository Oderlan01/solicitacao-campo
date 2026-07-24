// ================================================================
// Power Query — Aba Componentes v2
// Fonte: Desktop\todos_atributos.csv (gerado pelo STWExportar / LISP)
//
// ARQUITETURA (hierarquia resolvida no Power Query — opcao 1B):
//   - Cada linha do CSV = 1 bloco do CAD (pai OU filho), com HANDLE proprio.
//   - Vinculo filho -> pai: coluna ID_PAI (= HANDLE do bloco pai).
//   - EQUIPAMENTO (tag mestre do grupo): self-join via ID_PAI; o pai usa o
//     proprio Atributo. ATRIBUTO (tag pura) = TAG apos o primeiro hifen
//     (ex.: MTE-RO-100 -> RO-100); sem hifen, a propria TAG.
//   - TAG ja vem PRONTA do LISP (nao concatenar).
//   - Sufixo (n) em colunas BASE (TAG(2), MODELO(2), CORRENTE(2)...):
//     gera REPLICA da linha que herda tudo e sobrepoe os campos (n).
//     SUBORDEM = n.
//   - Colunas NAO-base preenchidas (SENSOR_*, VAVULA_*, CELULA_DE_CARGA(n)...):
//     geram LINHA FILHA: CATEGORIA = nome da coluna, TAG = valor da celula.
//     Filhos laterais herdam o contexto (LOCAL/SETOR/FAMILIA/CAIXA); campos
//     tecnicos (MODELO/ACIONAMENTO/POTENCIA/...) so se o nome da coluna
//     contiver VAVULA/VALVULA (regra herdaTec).
//   - INDICE: UM POR BLOCO — a linha pai e todas as suas filhas (replicas,
//     laterais e blocos-filho via ID_PAI) recebem o MESMO numero, renumerado
//     sequencialmente (1, 2, 3...) apos a ordenacao.
//   - Colunas auxiliares do AutoCAD / metadados: lista ColunasExcluir (edite).
//
// Saida: 20 colunas fixas do gabarito de homologacao, Motor no topo do grupo.
// ================================================================
let
    // ------------------------------------------------------------
    // ETAPA 0: CONFIGURACAO (edite aqui)
    // ------------------------------------------------------------
    CaminhoCSV = "C:\Users\oderlan.colcenti\Desktop\todos_atributos.csv",

    // Colunas auxiliares do AutoCAD e metadados de desenho (descartadas).
    // Nomes em CAIXA ALTA (a Etapa 2 forca os cabecalhos para maiusculas).
    ColunasExcluir = {
        "2TAG_", "AUTOR", "DATA", "ESCALA", "REVISÃO", "ESTADO", "TIPO",
        "MUNICÍPIO", "PAÍS", "NOME_DA_EMPRESA", "NUM_CLI", "NUM_DESENHO_CLIENTE",
        "NUM_DESENHO_STW", "NUM_OBRA", "NUM_PROJ", "DETALHE_PROJETO1",
        "ID_AUTOMAÇÃO", "INSTALAÇÃO", "DESCRIÇÃO_GERAL", "ID_VISIVEL",
        "NOME_BLOCO", "NOME_DO_BLOCO", "NAME", "SUFIXO",
        "DESTINO_DIREITA", "DESTINO_ESQUERDA"
    },

    // Colunas BASE (nao geram filho lateral; suas variantes "(n)" sobrepoem
    // as replicas de sufixo). Tudo que nao for base nem excluido vira filho.
    ColunasBase = {
        "HANDLE", "ID_PAI", "LOCAL", "TAG", "EQUIPAMENTO AUX", "SETOR",
        "FAMILIA", "CATEGORIA", "MODELO", "ACIONAMENTO", "POTÊNCIA",
        "CORRENTE", "TENSÃO", "CAPACIDADE", "CAIXA_DE_PASSAGEM", "ÍNDICE_LINHA"
    },

    // ------------------------------------------------------------
    // ETAPA 1: CARGA (CSV ';' em ANSI 1252 — padrao do AutoCAD)
    // ------------------------------------------------------------
    Fonte = Csv.Document(
        File.Contents(CaminhoCSV),
        [Delimiter=";", Columns=null, Encoding=1252, QuoteStyle=QuoteStyle.None]
    ),
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),

    // ------------------------------------------------------------
    // ETAPA 2: BLINDAGEM CASE-SENSITIVE (cabecalhos em CAIXA ALTA)
    // ------------------------------------------------------------
    #"Maiúsculas" = Table.TransformColumnNames(#"Cabeçalhos Promovidos", Text.Upper),

    // ------------------------------------------------------------
    // ETAPA 3: PURGA CONFIGURAVEL (auxiliares do AutoCAD)
    // ------------------------------------------------------------
    #"Purgada" = Table.RemoveColumns(
        #"Maiúsculas",
        List.Intersect({Table.ColumnNames(#"Maiúsculas"), ColunasExcluir})
    ),

    // ------------------------------------------------------------
    // ETAPA 4: DICIONARIO DE TRADUCAO (chaves CAD -> nomes de relatorio)
    // ID_PAI NAO e renomeado nem exibido: fica interno, so para o self-join
    // (evita colisao com a antiga coluna manual ORIGEM, substituida por LOCAL).
    // ------------------------------------------------------------
    Renomes = {{"HANDLE_CAD", "HANDLE"}, {"TAG_AUXILIAR", "EQUIPAMENTO AUX"}},
    #"Renomeada" = Table.RenameColumns(
        #"Purgada",
        List.Select(Renomes, each List.Contains(Table.ColumnNames(#"Purgada"), _{0}))
    ),

    // Linhas sem HANDLE nao sao blocos validos
    #"Filtrada" = Table.SelectRows(
        #"Renomeada",
        each [HANDLE] <> null and Text.Trim(Text.From([HANDLE])) <> ""
    ),

    // ------------------------------------------------------------
    // ETAPA 6: INDEXADOR DE LEITURA + BUFFER
    // INDICE_LINHA = ordem de leitura do bloco (interno). O INDICE exibido
    // (um por bloco pai) e renumerado na Etapa 9.
    // ------------------------------------------------------------
    #"Indexada" = Table.AddIndexColumn(#"Filtrada", "ÍNDICE_LINHA", 1, 1, Int64.Type),
    Base = Table.Buffer(#"Indexada"),
    TodasColunas = Table.ColumnNames(Base),

    // ------------------------------------------------------------
    // ETAPAS 5 e 7: FUNCOES AUXILIARES
    // ------------------------------------------------------------
    // FixCase: trim -> corta do "(" em diante -> Proper Case
    FixCase = (texto) =>
        let
            t     = if texto = null then "" else Text.Trim(Text.From(texto)),
            pos   = Text.PositionOf(t, "("),
            limpo = if pos > 0 then Text.Start(t, pos) else t
        in Text.Proper(Text.Trim(limpo)),

    // GetCampo: coluna ausente OU celula null -> "" (nunca erro)
    GetCampo = (linha as record, nomeCol as text) =>
        if List.Contains(TodasColunas, nomeCol)
        then (let v = Record.Field(linha, nomeCol)
              in if v = null then "" else Text.Trim(Text.From(v)))
        else "",

    NomeLimpo = (col as text) =>
        if Text.Contains(col, "(") then Text.Trim(Text.BeforeDelimiter(col, "(")) else col,
    SufixoDe = (col as text) =>
        if Text.Contains(col, "(") then Text.BetweenDelimiters(col, "(", ")") else "",

    // Atributo (tag pura) = texto apos o primeiro hifen da TAG
    DerivarAtributo = (tag as text) =>
        if Text.Contains(tag, "-") then Text.AfterDelimiter(tag, "-") else tag,

    // ------------------------------------------------------------
    // LOOKUP DE PAIS: HANDLE -> {TAG do pai, INDICE_LINHA do pai}
    // ------------------------------------------------------------
    PaisDistintos = Table.Distinct(
        Table.SelectColumns(Base, {"HANDLE", "TAG", "ÍNDICE_LINHA"}),
        {"HANDLE"}),
    ListaPais = Table.ToRecords(PaisDistintos),
    RegPais = Record.FromList(
        List.Transform(ListaPais, each [TAG = Text.From([TAG] ?? ""), IDX = [ÍNDICE_LINHA]]),
        List.Transform(ListaPais, each Text.Upper(Text.Trim(Text.From([HANDLE]))))),
    BuscarPai = (h as text) => Record.FieldOrDefault(RegPais, h, null),

    // Colunas laterais = tudo que nao e base (pelo nome limpo, sem "(n)")
    ColsLaterais = List.Select(TodasColunas, each not List.Contains(ColunasBase, NomeLimpo(_))),

    // ------------------------------------------------------------
    // ETAPA 8: MOTOR DE HIERARQUIA (linha do bloco + replicas + filhos)
    // ------------------------------------------------------------
    #"Linhas Geradas" = Table.AddColumn(Base, "LISTA_FINAL", each
        let
            L = _,

            // Campo base com sobreposicao de sufixo: BASE(n) > BASE
            GetSuf = (nomeBase as text, suf as text) =>
                let
                    v = if suf = "" then ""
                        else (let a = GetCampo(L, nomeBase & "(" & suf & ")")
                              in if a <> "" then a
                                 else GetCampo(L, nomeBase & " (" & suf & ")"))
                in if v <> "" then v else GetCampo(L, nomeBase),

            idPai       = GetCampo(L, "ID_PAI"),
            pai         = if idPai <> "" then BuscarPai(Text.Upper(idPai)) else null,
            tagBase     = GetCampo(L, "TAG"),
            equipMestre = Text.Upper(if pai <> null then DerivarAtributo(pai[TAG])
                                     else DerivarAtributo(tagBase)),
            idxGrupo    = if pai <> null then pai[IDX] else L[ÍNDICE_LINHA],

            // Fabrica de registro: 20 colunas fixas + controle interno
            Registro = (categoria as text, tagLinha as text, suf as text, herdaTec as logical) =>
                [
                    ÍNDICE_LINHA         = L[ÍNDICE_LINHA],
                    ÍNDICE_GRUPO         = idxGrupo,
                    HANDLE               = GetCampo(L, "HANDLE"),
                    LOCAL                = GetCampo(L, "LOCAL"),
                    EQUIPAMENTO          = equipMestre,
                    ATRIBUTO             = Text.Upper(DerivarAtributo(tagLinha)),
                    TAG                  = Text.Upper(tagLinha),
                    #"EQUIPAMENTO AUX"   = Text.Upper(GetSuf("EQUIPAMENTO AUX", suf)),
                    SETOR                = GetCampo(L, "SETOR"),
                    FAMILIA              = FixCase(GetCampo(L, "FAMILIA")),
                    CATEGORIA            = FixCase(categoria),
                    MODELO               = if herdaTec then FixCase(GetSuf("MODELO", suf)) else "",
                    #"DESCRIÇÃO GERAL"   = FixCase(categoria) & " | " &
                                           FixCase(GetCampo(L, "FAMILIA")) & " " &
                                           Text.Upper(tagLinha),
                    ACIONAMENTO          = if herdaTec then FixCase(GetSuf("ACIONAMENTO", suf)) else "",
                    INTERFACE            = null,
                    POTÊNCIA             = if herdaTec then GetSuf("POTÊNCIA", suf) else "",
                    CORRENTE             = if herdaTec then GetSuf("CORRENTE", suf) else "",
                    TENSÃO               = if herdaTec then GetSuf("TENSÃO", suf) else "",
                    CAPACIDADE           = if herdaTec then GetSuf("CAPACIDADE", suf) else "",
                    #"CAIXA DE PASSAGEM" = FixCase(GetCampo(L, "CAIXA_DE_PASSAGEM")),
                    CONEXÃO              = null,
                    SUBORDEM             = suf
                ],

            // 1) Linha do proprio bloco (pai ou filho via ID_PAI)
            LinhaBloco =
                if tagBase <> "" or GetCampo(L, "CATEGORIA") <> ""
                then { Registro(GetCampo(L, "CATEGORIA"), tagBase, "", true) }
                else {},

            // 2) Replicas de sufixo: uma por TAG(n) preenchida; herda tudo e
            //    sobrepoe CATEGORIA(n)/MODELO(n)/ACIONAMENTO(n)/eletrica(n)
            SufsTag = List.Select(
                List.Transform(
                    List.Select(TodasColunas, each Text.StartsWith(_, "TAG(")),
                    each SufixoDe(_)),
                each GetCampo(L, "TAG(" & _ & ")") <> ""),
            Replicas = List.Transform(SufsTag, (n) =>
                Registro(GetSuf("CATEGORIA", n), GetCampo(L, "TAG(" & n & ")"), n, true)),

            // 3) Filhos laterais: coluna nao-base preenchida
            //    CATEGORIA = nome da coluna; TAG = valor da celula
            //    herdaTec so para VAVULA/VALVULA
            Laterais = List.Transform(
                List.Select(ColsLaterais, each GetCampo(L, _) <> ""),
                (col) =>
                    let
                        nm     = NomeLimpo(col),
                        n      = SufixoDe(col),
                        nmUp   = Text.Upper(nm),
                        ehValv = Text.Contains(nmUp, "VÁVULA") or
                                 Text.Contains(nmUp, "VÁLVULA") or
                                 Text.Contains(nmUp, "VALVULA")
                    in Registro(Text.Replace(nm, "_", " "), GetCampo(L, col), n, ehValv))
        in
            LinhaBloco & Replicas & Laterais),

    #"Tabela Expandida" = Table.FromRecords(
        List.Combine(Table.Column(#"Linhas Geradas", "LISTA_FINAL"))),

    // ------------------------------------------------------------
    // ETAPA 9: ORDENACAO + INDICE UNICO POR BLOCO
    // ------------------------------------------------------------
    #"Peso Adicionado" = Table.AddColumn(#"Tabela Expandida", "PESO",
        each if Text.Contains(Text.Upper(Text.From([CATEGORIA] ?? "")), "MOTOR") then 0 else 1),
    #"Ordenado" = Table.Sort(#"Peso Adicionado", {
        {"ÍNDICE_GRUPO", Order.Ascending},
        {"PESO",         Order.Ascending},
        {"ÍNDICE_LINHA", Order.Ascending},
        {"CATEGORIA",    Order.Ascending},
        {"SUBORDEM",     Order.Ascending}
    }),

    // Renumera o INDICE: um numero sequencial POR BLOCO (pai + filhas juntos)
    GruposOrdenados = List.Buffer(List.Distinct(Table.Column(#"Ordenado", "ÍNDICE_GRUPO"))),
    #"Índice Por Bloco" = Table.AddColumn(#"Ordenado", "ÍNDICE",
        each List.PositionOf(GruposOrdenados, [ÍNDICE_GRUPO]) + 1, Int64.Type),

    // ------------------------------------------------------------
    // ETAPA 10: CARGA FINAL — 20 COLUNAS FIXAS DO GABARITO
    // ------------------------------------------------------------
    #"Seleção Final" = Table.SelectColumns(#"Índice Por Bloco", {
        "ÍNDICE", "HANDLE", "LOCAL", "EQUIPAMENTO", "ATRIBUTO", "TAG",
        "EQUIPAMENTO AUX", "SETOR", "FAMILIA", "CATEGORIA", "MODELO",
        "DESCRIÇÃO GERAL", "ACIONAMENTO", "INTERFACE", "POTÊNCIA", "CORRENTE",
        "TENSÃO", "CAPACIDADE", "CAIXA DE PASSAGEM", "CONEXÃO"
    })
in
    #"Seleção Final"
</content>
