// ================================================================
// Power Query — Aba Componentes (compoteste.xlsm)
//
// MUDANÇAS em relação à versão anterior:
//   1. ID VISIVEL e NOME DO BLOCO excluídos de ColsFilhosLaterais
//      (evita geração de linhas espúrias com CATEGORIA="ID VISIVEL")
//   2. Campo HANDLE adicionado nas linhas pai (ReplicasPai)
//   3. Campo HANDLE adicionado nas linhas filhos (Filhos)
//   4. Coluna HANDLE incluída na Seleção Final
//
// Como usar: Cole este código no Editor Avançado da query Componentes
// ================================================================
let
    // 1. CARREGAMENTO E PREPARAÇÃO
    Fonte = Excel.Workbook(File.Contents("C:\Users\oderlan.colcenti\Desktop\teste.xlsx"), null, true),
    Summary_Sheet = Fonte{[Item="Summary",Kind="Sheet"]}[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Summary_Sheet, [PromoteAllScalars=true]),
    #"Nomes Ajustados" = Table.TransformColumnNames(#"Cabeçalhos Promovidos", each Text.Replace(_, "_", " ")),
    #"Filtro Inicial" = Table.SelectRows(#"Nomes Ajustados", each ([0A CATEGORIA] <> null)),
    #"Fonte Bufferizada" = Table.Buffer(#"Filtro Inicial"),

    ColunasParaRemover = {"0J SENSOR", "0K TAG SIZE", "0L TAG AMOUNT", "0M NAME", "9A TAG ", "4A AUX1", "4B AUX2"},
    #"Colunas Removidas" = Table.RemoveColumns(#"Fonte Bufferizada", List.Intersect({Table.ColumnNames(#"Fonte Bufferizada"), ColunasParaRemover})),

    DicionarioNomes = {
        {"0F TAG AUXILIAR", "EQUIPAMENTO AUX"}, {"0B SETOR", "SETOR"}, {"0C FAMILIA", "FAMILIA"},
        {"0D MODELO", "MODELO"}, {"0E TAG", "EQUIPAMENTO"}, {"0G DESCRIÇÃO GERAL", "DESCRIÇÃO GERAL"},
        {"0I DADOS ENGENHARIA", "DADOS ENGENHARIA"}, {"0H ACIONAMENTO", "ACIONAMENTO"}, {"0A CATEGORIA", "CATEGORIA"}
    },
    #"Colunas Renomeadas" = Table.RenameColumns(#"Colunas Removidas", List.Select(DicionarioNomes, each List.Contains(Table.ColumnNames(#"Colunas Removidas"), _{0}))),

    FixCase = (texto) =>
        let
            t = if texto = null then "" else Text.Trim(Text.From(texto)),
            pos = Text.PositionOf(t, "("),
            limpo = if pos > 0 then Text.Start(t, pos) else t,
            Final = Text.Proper(Text.Trim(limpo))
        in Final,

    #"TAG Inicializada" = if List.Contains(Table.ColumnNames(#"Colunas Renomeadas"), "TAG") then #"Colunas Renomeadas" else Table.AddColumn(#"Colunas Renomeadas", "TAG", each [EQUIPAMENTO], type text),
    #"ÍNDICE Adicionado" = Table.AddIndexColumn(#"TAG Inicializada", "ÍNDICE", 1, 1, Int64.Type),

    TodasColunas = Table.ColumnNames(#"ÍNDICE Adicionado"),
    ColunasBaseSufixo = {"TAG", "ACIONAMENTO", "DADOS ENGENHARIA", "MODELO","EQUIPAMENTO AUX"},

    // 2. LÓGICA DE TRANSPOSIÇÃO
    #"Linhas Geradas" = Table.AddColumn(#"ÍNDICE Adicionado", "Lista_Final", each
        let
            Linha = _,

            // --- IDENTIFICAÇÃO DE DADOS EXTRAS ---
            ColsTagSufixo = List.Select(TodasColunas, each Text.StartsWith(_, "TAG(")),
            SufixosValidos = List.Select(ColsTagSufixo, each Record.Field(Linha, _) <> null and Text.From(Record.Field(Linha, _)) <> ""),

            // MUDANÇA 1: "ID VISIVEL" e "NOME DO BLOCO" excluídos para não virarem linhas filhos
            ColsFilhosLaterais = List.Select(TodasColunas, each
                let NomeLimpo = if Text.Contains(_, "(") then Text.Trim(Text.BeforeDelimiter(_, "(")) else _
                in not List.Contains(ColunasBaseSufixo, NomeLimpo) and
                   not List.Contains(
                       {"EQUIPAMENTO", "SETOR", "FAMILIA", "CATEGORIA", "DESCRIÇÃO GERAL",
                        "ÍNDICE", "CAIXA DE PASSAGEM", "ID VISIVEL", "NOME DO BLOCO"},
                       NomeLimpo)
            ),
            FilhosLateraisValidos = List.Select(ColsFilhosLaterais, each Record.Field(Linha, _) <> null and Text.From(Record.Field(Linha, _)) <> ""),

            IndicesSufixo = if List.IsEmpty(SufixosValidos) then {"ORIGINAL"} else List.Transform(SufixosValidos, each Text.BetweenDelimiters(_, "(", ")")),

            // --- BLOCO A: LINHAS PAI ---
            ReplicasPai = List.Transform(IndicesSufixo, (idx) =>
                let
                    s = if idx = "ORIGINAL" then "" else "(" & idx & ")",
                    GetAtributo = (nomeBase, original) =>
                        let
                            c1 = nomeBase & s, c2 = nomeBase & " " & s,
                            valSufixo = if idx = "ORIGINAL" then original
                                        else if List.Contains(TodasColunas, c1) then Record.Field(Linha, c1)
                                        else if List.Contains(TodasColunas, c2) then Record.Field(Linha, c2)
                                        else null
                        in if valSufixo <> null and valSufixo <> "" then valSufixo else original,

                    CategoriaDinamica = GetAtributo("CATEGORIA", Linha[CATEGORIA]),
                    DescPai = FixCase(CategoriaDinamica) & " | " & FixCase(Linha[FAMILIA]) & " " & Text.Upper(Text.Trim(Text.From(Linha[EQUIPAMENTO] ?? ""))),
                    PossuiExtras = List.Count(SufixosValidos) > 0 or List.Count(FilhosLateraisValidos) > 0,
                    EhMotorPai = Text.Contains(Text.Upper(Text.From(CategoriaDinamica ?? "")), "MOTOR"),
                    TipoResultado = if idx = "ORIGINAL" and PossuiExtras and not EhMotorPai then "Excluir" else "Manter"
                in [
                    ÍNDICE = Linha[ÍNDICE],
                    EQUIPAMENTO = Text.Upper(Text.From(Linha[EQUIPAMENTO])),
                    EQUIPAMENTO AUX = Text.Upper(Text.From(GetAtributo("EQUIPAMENTO AUX", Linha[EQUIPAMENTO AUX]) ?? "")),
                    SETOR = Linha[SETOR],
                    FAMILIA = Linha[FAMILIA],
                    CATEGORIA = CategoriaDinamica,
                    TAG = Text.Upper(Text.From(GetAtributo("TAG", Linha[TAG]))),
                    DESCRIÇÃO GERAL = DescPai,
                    ACIONAMENTO = GetAtributo("ACIONAMENTO", Linha[ACIONAMENTO]),
                    DADOS ENGENHARIA = GetAtributo("DADOS ENGENHARIA", Linha[DADOS ENGENHARIA]),
                    MODELO = GetAtributo("MODELO", Linha[MODELO]),
                    CAIXA DE PASSAGEM = Linha[CAIXA DE PASSAGEM],
                    // MUDANÇA 2: HANDLE adicionado nas linhas pai
                    HANDLE = Text.Upper(Text.From(Linha[ID VISIVEL] ?? "")),
                    ORDEM_INTERNA = 0,
                    SUBORDEM = if idx = "ORIGINAL" then "" else idx,
                    Tipo = TipoResultado
                ]
            ),

            // --- BLOCO B: LINHAS FILHOS (SENSORES/INSTRUMENTOS) ---
            Filhos = List.Transform(ColsFilhosLaterais, (col) =>
                let
                    val = Record.Field(Linha, col),
                    EhValvulaFilho = Text.Contains(Text.Upper(col), "VÁLVULA"),
                    NomeCatFilho = if Text.Contains(col, "(") then Text.Trim(Text.BeforeDelimiter(col, "(")) else col,
                    IdxFilho = if Text.Contains(col, "(") then Text.BetweenDelimiters(col, "(", ")") else "",
                    DescFilho = FixCase(NomeCatFilho) & " | " & FixCase(Linha[FAMILIA]) & " " & Text.Upper(Text.Trim(Text.From(Linha[EQUIPAMENTO] ?? "")))
                in if val <> null and Text.Trim(Text.From(val)) <> "" then [
                    ÍNDICE = Linha[ÍNDICE],
                    EQUIPAMENTO = Text.Upper(Text.From(Linha[EQUIPAMENTO])),
                    EQUIPAMENTO AUX = Text.Upper(Text.From(Linha[EQUIPAMENTO AUX] ?? "")),
                    SETOR = Linha[SETOR],
                    FAMILIA = Linha[FAMILIA],
                    CATEGORIA = NomeCatFilho,
                    TAG = Text.Upper(Text.From(val)),
                    DESCRIÇÃO GERAL = DescFilho,
                    ACIONAMENTO = if EhValvulaFilho then Linha[ACIONAMENTO] else null,
                    DADOS ENGENHARIA = if EhValvulaFilho then Linha[DADOS ENGENHARIA] else null,
                    MODELO = if EhValvulaFilho then Linha[MODELO] else null,
                    CAIXA DE PASSAGEM = Linha[CAIXA DE PASSAGEM],
                    // MUDANÇA 3: HANDLE adicionado nas linhas filhos (mesmo bloco pai)
                    HANDLE = Text.Upper(Text.From(Linha[ID VISIVEL] ?? "")),
                    ORDEM_INTERNA = 1,
                    SUBORDEM = IdxFilho,
                    Tipo = "Manter"
                ] else null
            ),
            Result = List.Select(ReplicasPai & Filhos, each _ <> null)
        in Result
    ),

    // 3. EXPANSÃO E COMPLEMENTO DE ENGENHARIA
    #"Tabela Expandida" = Table.FromRecords(List.Combine(Table.Column(#"Linhas Geradas", "Lista_Final"))),
    #"Filtro Manter" = Table.SelectRows(#"Tabela Expandida", each ([Tipo] = "Manter")),

    #"Engenharia Dividida" = Table.AddColumn(#"Filtro Manter", "EngSplit", each
        let
            texto = Text.From([DADOS ENGENHARIA] ?? ""),
            partes = Text.Split(texto, "/"),
            valid = List.Count(partes) >= 3
        in if texto = "" or texto = "//" then [P="--", C="--", T="--"]
           else if valid then [P=partes{0}, C=partes{1}, T=partes{2}]
           else [P=texto, C="--", T="--"]),
    #"Colunas Engenharia" = Table.ExpandRecordColumn(#"Engenharia Dividida", "EngSplit", {"P", "C", "T"}, {"POTÊNCIA", "CORRENTE", "TENSÃO"}),

    // 4. TRATAMENTOS DE TEXTO
    #"Tratamento Texto" = Table.TransformColumns(#"Colunas Engenharia", {
        {"CATEGORIA", each FixCase(_)}, {"SETOR", each FixCase(_)}, {"FAMILIA", each FixCase(_)},
        {"ACIONAMENTO", each FixCase(_)}, {"MODELO", each FixCase(_)}, {"CAIXA DE PASSAGEM", each FixCase(_)}
    }),

    #"Colunas Extras" = Table.AddColumn(Table.AddColumn(Table.AddColumn(#"Tratamento Texto", "INTERFACE", each null), "CONEXÃO", each null), "ORIGEM", each null),

    // 5. ORDENAÇÃO PRIORITÁRIA
    #"Peso Temporario" = Table.AddColumn(#"Colunas Extras", "Peso", each if Text.Contains(Text.Upper([CATEGORIA]), "MOTOR") then 0 else 1),

    #"Ordenado" = Table.Sort(#"Peso Temporario", {
        {"ÍNDICE", Order.Ascending},
        {"Peso", Order.Ascending},
        {"CATEGORIA", Order.Ascending},
        {"SUBORDEM", Order.Ascending}
    }),

    // MUDANÇA 4: HANDLE incluído na seleção final (posição após ÍNDICE)
    #"Seleção Final" = Table.SelectColumns(#"Ordenado", {
        "ÍNDICE", "HANDLE", "EQUIPAMENTO", "EQUIPAMENTO AUX", "SETOR", "FAMILIA",
        "CATEGORIA", "TAG", "DESCRIÇÃO GERAL", "ACIONAMENTO", "INTERFACE",
        "POTÊNCIA", "CORRENTE", "TENSÃO", "MODELO", "CAIXA DE PASSAGEM", "CONEXÃO", "ORIGEM"
    })
in
    #"Seleção Final"
