// ================================================================
// Power Query — Aba Componentes (compoteste.xlsm)
// Fonte: Desktop\todos_atributos.csv (gerado pelo STWExportar)
//
// Fluxo (sem DATAEXTRACTION manual):
//   1. AutoCAD: STWExportar      -> Desktop\todos_atributos.csv
//   2. Excel:   Atualizar query  <- Desktop\todos_atributos.csv
//
// Logica de expansao preservada:
//   - Sufixos TAG(1), TAG(A), ACIONAMENTO(1)... geram linhas pai extras
//   - Colunas laterais (VALVULA ABRE, SENSOR...) geram linhas filho
//   - Estrutura final identica a versao com DATAEXTRACTION
// ================================================================
let
    // ---------------------------------------------------------------
    // 1. CARREGAMENTO DO CSV
    // ---------------------------------------------------------------
    CaminhoCSV = "C:\Users\oderlan.colcenti\Desktop\todos_atributos.csv",
    Fonte = Csv.Document(
        File.Contents(CaminhoCSV),
        [Delimiter=";", Columns=null, Encoding=1252, QuoteStyle=QuoteStyle.None]
    ),
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),

    // Normaliza nomes de colunas: substitui _ por espaço
    // (igual ao passo Text.Replace do Power Query original com DATAEXTRACTION)
    #"Nomes Normalizados" = Table.TransformColumnNames(
        #"Cabeçalhos Promovidos",
        each Text.Replace(_, "_", " ")
    ),

    // Filtro inicial: remove linhas sem HANDLE CAD
    #"Filtro Inicial" = Table.SelectRows(
        #"Nomes Normalizados",
        each [#"HANDLE CAD"] <> null and Text.Trim(Text.From([#"HANDLE CAD"])) <> ""
    ),
    #"Fonte Bufferizada" = Table.Buffer(#"Filtro Inicial"),

    // ---------------------------------------------------------------
    // 2. RENOMEAR COLUNAS (tag normalizado -> nome amigavel)
    // ---------------------------------------------------------------

    // Remove atributos que colidem com nomes amigaveis:
    //   NOME_DO_BLOCO (atributo do bloco) -> "NOME DO BLOCO" apos normalizacao,
    //   conflita com a renomeacao de "Nome Bloco" -> "NOME DO BLOCO"
    //   ID_VISIVEL (atributo do bloco) -> redundante, HANDLE ja vem de HANDLE_CAD
    ColsPre = Table.ColumnNames(#"Fonte Bufferizada"),
    #"Atributos Redundantes Removidos" = Table.RemoveColumns(
        #"Fonte Bufferizada",
        List.Intersect({ColsPre, {"NOME DO BLOCO", "ID VISIVEL"}})
    ),

    DicionarioNomes = {
        {"HANDLE CAD",          "HANDLE"},
        {"Nome Bloco",          "NOME DO BLOCO"},
        {"0A CATEGORIA",        "CATEGORIA"},
        {"0B SETOR",            "SETOR"},
        {"0C FAMILIA",          "FAMILIA"},
        {"0D MODELO",           "MODELO"},
        {"0E TAG",              "EQUIPAMENTO"},
        {"0F TAG AUXILIAR",     "EQUIPAMENTO AUX"},
        {"0G DESCRICAO GERAL",  "DESCRIÇÃO GERAL"},
        {"0G DESCRI" & Character.FromNumber(199) & Character.FromNumber(195) & "O GERAL", "DESCRIÇÃO GERAL"},
        {"0H ACIONAMENTO",      "ACIONAMENTO"},
        {"0I DADOS ENGENHARIA", "DADOS ENGENHARIA"},
        {"CAIXA DE PASSAGEM",   "CAIXA DE PASSAGEM"}
    },
    ColsAtuais = Table.ColumnNames(#"Atributos Redundantes Removidos"),
    #"Colunas Renomeadas" = Table.RenameColumns(
        #"Atributos Redundantes Removidos",
        List.Select(DicionarioNomes, each List.Contains(ColsAtuais, _{0}))
    ),

    // Remove colunas que nao sao usadas na tabela final (se existirem)
    ColunasParaRemover = {"0J SENSOR", "0K TAG SIZE", "0L TAG AMOUNT", "0M NAME", "9A TAG ", "4A AUX1", "4B AUX2"},
    #"Colunas Removidas" = Table.RemoveColumns(
        #"Colunas Renomeadas",
        List.Intersect({Table.ColumnNames(#"Colunas Renomeadas"), ColunasParaRemover})
    ),

    // Garante que todas as colunas base existam (null se ausente no CSV)
    // Evita erros ao acessar Linha[CAMPO] na logica de expansao
    ColsBase = {
        "EQUIPAMENTO", "EQUIPAMENTO AUX", "SETOR", "FAMILIA", "CATEGORIA",
        "DESCRIÇÃO GERAL", "ACIONAMENTO", "DADOS ENGENHARIA", "MODELO",
        "CAIXA DE PASSAGEM", "HANDLE"
    },
    #"Colunas Garantidas" = List.Accumulate(
        List.Select(ColsBase, each not List.Contains(Table.ColumnNames(#"Colunas Removidas"), _)),
        #"Colunas Removidas",
        (acc, col) => Table.AddColumn(acc, col, each null)
    ),

    // ---------------------------------------------------------------
    // 3. HELPERS E INICIALIZACAO
    // ---------------------------------------------------------------
    FixCase = (texto) =>
        let
            t     = if texto = null then "" else Text.Trim(Text.From(texto)),
            pos   = Text.PositionOf(t, "("),
            limpo = if pos > 0 then Text.Start(t, pos) else t
        in Text.Proper(Text.Trim(limpo)),

    // TAG inicializada (se nao existir como coluna, usa EQUIPAMENTO)
    #"TAG Inicializada" = if List.Contains(Table.ColumnNames(#"Colunas Garantidas"), "TAG")
        then #"Colunas Garantidas"
        else Table.AddColumn(#"Colunas Garantidas", "TAG", each [EQUIPAMENTO], type text),

    #"ÍNDICE Adicionado" = Table.AddIndexColumn(#"TAG Inicializada", "ÍNDICE", 1, 1, Int64.Type),

    TodasColunas     = Table.ColumnNames(#"ÍNDICE Adicionado"),
    ColunasBaseSufixo = {"TAG", "ACIONAMENTO", "DADOS ENGENHARIA", "MODELO", "EQUIPAMENTO AUX"},

    // ---------------------------------------------------------------
    // 4. LOGICA DE EXPANSAO PAI/FILHO
    //
    // Para cada linha do CSV (= 1 bloco AutoCAD):
    //   A) Sufixos TAG(1), TAG(A)... → cria cópia da linha pai
    //      substituindo TAG, ACIONAMENTO, MODELO, etc. pelo valor com sufixo
    //   B) Colunas laterais (VALVULA ABRE, SENSOR...)  → cria linhas filho
    //      com CATEGORIA = nome da coluna e TAG = valor da celula
    // ---------------------------------------------------------------
    #"Linhas Geradas" = Table.AddColumn(#"ÍNDICE Adicionado", "Lista Final", each
        let
            Linha = _,

            // --- Detecta colunas de sufixo: TAG(A), TAG(1), TAG(B)... ---
            ColsTagSufixo    = List.Select(TodasColunas, each Text.StartsWith(_, "TAG(")),
            SufixosValidos   = List.Select(ColsTagSufixo, each
                Record.Field(Linha, _) <> null and
                Text.From(Record.Field(Linha, _)) <> ""
            ),

            // --- Detecta colunas filho laterais ---
            // Exclui colunas base, de identificacao e de sufixo
            ColsFilhosLaterais = List.Select(TodasColunas, each
                let
                    NomeLimpo = if Text.Contains(_, "(")
                                then Text.Trim(Text.BeforeDelimiter(_, "("))
                                else _
                in not List.Contains(ColunasBaseSufixo, NomeLimpo) and
                   not List.Contains(
                       {"EQUIPAMENTO", "SETOR", "FAMILIA", "CATEGORIA", "DESCRIÇÃO GERAL",
                        "ÍNDICE", "CAIXA DE PASSAGEM", "ID VISIVEL", "NOME DO BLOCO",
                        "HANDLE", "ACIONAMENTO", "MODELO", "DADOS ENGENHARIA",
                        "EQUIPAMENTO AUX", "TAG", "ORDEM INTERNA", "SUBORDEM", "Tipo"},
                       NomeLimpo)
            ),
            FilhosLateraisValidos = List.Select(ColsFilhosLaterais, each
                Record.Field(Linha, _) <> null and
                Text.From(Record.Field(Linha, _)) <> ""
            ),

            // Lista de indices: "ORIGINAL" ou sufixos encontrados
            IndicesSufixo = if List.IsEmpty(SufixosValidos)
                then {"ORIGINAL"}
                else List.Transform(SufixosValidos, each Text.BetweenDelimiters(_, "(", ")")),

            // --- BLOCO A: LINHAS PAI (uma por sufixo ou a linha original) ---
            ReplicasPai = List.Transform(IndicesSufixo, (idx) =>
                let
                    s = if idx = "ORIGINAL" then "" else "(" & idx & ")",

                    // Busca o valor da coluna com sufixo;
                    // se nao encontrar, retorna o valor original
                    GetAtributo = (nomeBase, original) =>
                        let
                            c1       = nomeBase & s,
                            c2       = nomeBase & " " & s,
                            valSufixo =
                                if idx = "ORIGINAL" then original
                                else if List.Contains(TodasColunas, c1) then Record.Field(Linha, c1)
                                else if List.Contains(TodasColunas, c2) then Record.Field(Linha, c2)
                                else null
                        in if valSufixo <> null and valSufixo <> "" then valSufixo else original,

                    CategoriaDinamica = GetAtributo("CATEGORIA", Linha[CATEGORIA]),
                    DescPai     = FixCase(CategoriaDinamica) & " | " &
                                  FixCase(Linha[FAMILIA]) & " " &
                                  Text.Upper(Text.Trim(Text.From(Linha[EQUIPAMENTO] ?? ""))),
                    PossuiExtras = List.Count(SufixosValidos) > 0 or List.Count(FilhosLateraisValidos) > 0,
                    EhMotorPai   = Text.Contains(Text.Upper(Text.From(CategoriaDinamica ?? "")), "MOTOR"),
                    TipoResultado = if idx = "ORIGINAL" and PossuiExtras and not EhMotorPai
                                    then "Excluir" else "Manter"
                in [
                    ÍNDICE              = Linha[ÍNDICE],
                    HANDLE              = Linha[HANDLE],
                    EQUIPAMENTO         = Text.Upper(Text.From(Linha[EQUIPAMENTO])),
                    #"EQUIPAMENTO AUX"  = Text.Upper(Text.From(GetAtributo("EQUIPAMENTO AUX", Linha[#"EQUIPAMENTO AUX"]) ?? "")),
                    SETOR               = Linha[SETOR],
                    FAMILIA             = Linha[FAMILIA],
                    CATEGORIA           = CategoriaDinamica,
                    TAG                 = Text.Upper(Text.From(GetAtributo("TAG", Linha[TAG]))),
                    #"DESCRIÇÃO GERAL"  = DescPai,
                    ACIONAMENTO         = GetAtributo("ACIONAMENTO", Linha[ACIONAMENTO]),
                    #"DADOS ENGENHARIA" = GetAtributo("DADOS ENGENHARIA", Linha[#"DADOS ENGENHARIA"]),
                    MODELO              = GetAtributo("MODELO", Linha[MODELO]),
                    #"CAIXA DE PASSAGEM" = Linha[#"CAIXA DE PASSAGEM"],
                    ORDEM_INTERNA       = 0,
                    SUBORDEM            = if idx = "ORIGINAL" then "" else idx,
                    Tipo                = TipoResultado
                ]
            ),

            // --- BLOCO B: LINHAS FILHO (colunas laterais) ---
            Filhos = List.Transform(ColsFilhosLaterais, (col) =>
                let
                    val         = Record.Field(Linha, col),
                    colUpper    = Text.Upper(col),
                    EhValvula   = Text.Contains(colUpper, "VALVULA") or
                                  Text.Contains(colUpper, "VÁLVULA"),
                    NomeCatFilho = if Text.Contains(col, "(")
                                   then Text.Trim(Text.BeforeDelimiter(col, "("))
                                   else col,
                    IdxFilho    = if Text.Contains(col, "(")
                                  then Text.BetweenDelimiters(col, "(", ")")
                                  else "",
                    DescFilho   = FixCase(NomeCatFilho) & " | " &
                                  FixCase(Linha[FAMILIA]) & " " &
                                  Text.Upper(Text.Trim(Text.From(Linha[EQUIPAMENTO] ?? "")))
                in if val <> null and Text.Trim(Text.From(val)) <> "" then [
                    ÍNDICE              = Linha[ÍNDICE],
                    HANDLE              = Linha[HANDLE],
                    EQUIPAMENTO         = Text.Upper(Text.From(Linha[EQUIPAMENTO])),
                    #"EQUIPAMENTO AUX"  = Text.Upper(Text.From(Linha[#"EQUIPAMENTO AUX"] ?? "")),
                    SETOR               = Linha[SETOR],
                    FAMILIA             = Linha[FAMILIA],
                    CATEGORIA           = NomeCatFilho,
                    TAG                 = Text.Upper(Text.From(val)),
                    #"DESCRIÇÃO GERAL"  = DescFilho,
                    ACIONAMENTO         = if EhValvula then Linha[ACIONAMENTO] else null,
                    #"DADOS ENGENHARIA" = if EhValvula then Linha[#"DADOS ENGENHARIA"] else null,
                    MODELO              = if EhValvula then Linha[MODELO] else null,
                    #"CAIXA DE PASSAGEM" = Linha[#"CAIXA DE PASSAGEM"],
                    ORDEM_INTERNA       = 1,
                    SUBORDEM            = IdxFilho,
                    Tipo                = "Manter"
                ] else null
            ),

            Result = List.Select(ReplicasPai & Filhos, each _ <> null)
        in Result
    ),

    // ---------------------------------------------------------------
    // 5. EXPANSAO E FILTRO
    // ---------------------------------------------------------------
    #"Tabela Expandida" = Table.FromRecords(
        List.Combine(Table.Column(#"Linhas Geradas", "Lista Final"))
    ),
    #"Filtro Manter" = Table.SelectRows(#"Tabela Expandida", each ([Tipo] = "Manter")),

    // ---------------------------------------------------------------
    // 6. DADOS ENGENHARIA -> POTÊNCIA / CORRENTE / TENSÃO
    // ---------------------------------------------------------------
    #"Engenharia Dividida" = Table.AddColumn(#"Filtro Manter", "EngSplit", each
        let
            texto  = Text.From([#"DADOS ENGENHARIA"] ?? ""),
            partes = Text.Split(texto, "/"),
            valido = List.Count(partes) >= 3
        in if texto = "" or texto = "//"
           then [P="--", C="--", T="--"]
           else if valido
           then [P=partes{0}, C=partes{1}, T=partes{2}]
           else [P=texto,     C="--",       T="--"]
    ),
    #"Colunas Engenharia" = Table.ExpandRecordColumn(
        #"Engenharia Dividida", "EngSplit",
        {"P", "C", "T"}, {"POTÊNCIA", "CORRENTE", "TENSÃO"}
    ),

    // ---------------------------------------------------------------
    // 7. TRATAMENTO DE TEXTO
    // ---------------------------------------------------------------
    ColsTratamento  = {"CATEGORIA", "SETOR", "FAMILIA", "ACIONAMENTO", "MODELO", "CAIXA DE PASSAGEM"},
    ColsParaTratar  = List.Intersect({Table.ColumnNames(#"Colunas Engenharia"), ColsTratamento}),
    #"Tratamento Texto" = Table.TransformColumns(
        #"Colunas Engenharia",
        List.Transform(ColsParaTratar, each {_, each FixCase(_)})
    ),

    // ---------------------------------------------------------------
    // 8. COLUNAS MANUAIS (Excel-only: preenchidas pelo usuario)
    // ---------------------------------------------------------------
    #"Colunas Extras" = Table.AddColumn(
        Table.AddColumn(
            Table.AddColumn(#"Tratamento Texto", "INTERFACE", each null),
            "CONEXÃO", each null
        ),
        "ORIGEM", each null
    ),

    // ---------------------------------------------------------------
    // 9. ORDENACAO
    // ---------------------------------------------------------------
    #"Peso Temporario" = Table.AddColumn(#"Colunas Extras", "Peso",
        each if Text.Contains(Text.Upper(Text.From([CATEGORIA] ?? "")), "MOTOR") then 0 else 1
    ),
    #"Ordenado" = Table.Sort(#"Peso Temporario", {
        {"ÍNDICE",    Order.Ascending},
        {"Peso",      Order.Ascending},
        {"CATEGORIA", Order.Ascending},
        {"SUBORDEM",  Order.Ascending}
    }),

    // ---------------------------------------------------------------
    // 10. SELECAO FINAL — mesma estrutura da versao com DATAEXTRACTION
    // ---------------------------------------------------------------
    ColsFinais = {
        "ÍNDICE", "HANDLE", "EQUIPAMENTO", "EQUIPAMENTO AUX", "SETOR", "FAMILIA",
        "CATEGORIA", "TAG", "DESCRIÇÃO GERAL", "ACIONAMENTO", "INTERFACE",
        "POTÊNCIA", "CORRENTE", "TENSÃO", "MODELO", "CAIXA DE PASSAGEM", "CONEXÃO", "ORIGEM"
    },
    #"Seleção Final" = Table.SelectColumns(
        #"Ordenado",
        List.Intersect({ColsFinais, Table.ColumnNames(#"Ordenado")})
    )
in
    #"Seleção Final"
