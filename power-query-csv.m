// ================================================================
// Power Query — Aba Componentes
// Fonte: Desktop\todos_atributos.csv (gerado pelo STWExportar)
//
// Fluxo simplificado (sem DATAEXTRACTION):
//   1. No AutoCAD: execute STWExportar  -> gera todos_atributos.csv
//   2. No Excel:   clique em Atualizar na query Componentes
//
// Estrutura final identica a versao anterior (mesmas colunas).
// INTERFACE, CONEXAO e ORIGEM permanecem como campos manuais.
// ================================================================
let
    // ---------------------------------------------------------------
    // 1. CARREGAMENTO DO CSV
    // Ajuste o caminho se necessario (mesmo usuario do automacao.lsp)
    // ---------------------------------------------------------------
    CaminhoCSV = "C:\Users\oderlan.colcenti\Desktop\todos_atributos.csv",
    Fonte = Csv.Document(
        File.Contents(CaminhoCSV),
        [Delimiter=";", Columns=null, Encoding=1252, QuoteStyle=QuoteStyle.None]
    ),
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),

    // ---------------------------------------------------------------
    // 2. FILTRO INICIAL — remove linhas sem HANDLE valido
    // ---------------------------------------------------------------
    #"Filtro Inicial" = Table.SelectRows(
        #"Cabeçalhos Promovidos",
        each [HANDLE_CAD] <> null and Text.Trim(Text.From([HANDLE_CAD])) <> ""
    ),
    #"Fonte Bufferizada" = Table.Buffer(#"Filtro Inicial"),

    // ---------------------------------------------------------------
    // 3. RENOMEAR COLUNAS: tag AutoCAD -> nome amigavel
    // Inclui variantes com e sem acentuacao no nome do atributo
    // ---------------------------------------------------------------
    DicionarioNomes = {
        {"HANDLE_CAD",          "HANDLE"},
        {"Nome_Bloco",          "NOME DO BLOCO"},
        {"0A_CATEGORIA",        "CATEGORIA"},
        {"0B_SETOR",            "SETOR"},
        {"0C_FAMILIA",          "FAMILIA"},
        {"0D_MODELO",           "MODELO"},
        {"0E_TAG",              "EQUIPAMENTO"},
        {"0F_TAG_AUXILIAR",     "EQUIPAMENTO AUX"},
        {"0G_DESCRICAO_GERAL",  "_DESC_CSV"},
        {"0G_DESCRI" & Character.FromNumber(199) & Character.FromNumber(195) & "O_GERAL", "_DESC_CSV"},
        {"0H_ACIONAMENTO",      "ACIONAMENTO"},
        {"0I_DADOS_ENGENHARIA", "DADOS ENGENHARIA"},
        {"CAIXA_DE_PASSAGEM",   "CAIXA DE PASSAGEM"}
    },
    ColsAtuais = Table.ColumnNames(#"Fonte Bufferizada"),
    #"Colunas Renomeadas" = Table.RenameColumns(
        #"Fonte Bufferizada",
        List.Select(DicionarioNomes, each List.Contains(ColsAtuais, _{0}))
    ),

    // ---------------------------------------------------------------
    // 4. INDICE de linha
    // ---------------------------------------------------------------
    #"ÍNDICE Adicionado" = Table.AddIndexColumn(
        #"Colunas Renomeadas", "ÍNDICE", 1, 1, Int64.Type
    ),

    // ---------------------------------------------------------------
    // 5. HELPER FixCase
    //    Remove sufixo entre parenteses e aplica Title Case
    // ---------------------------------------------------------------
    FixCase = (texto) =>
        let
            t     = if texto = null then "" else Text.Trim(Text.From(texto)),
            pos   = Text.PositionOf(t, "("),
            limpo = if pos > 0 then Text.Start(t, pos) else t
        in Text.Proper(Text.Trim(limpo)),

    // ---------------------------------------------------------------
    // 6. TAG = EQUIPAMENTO em maiusculas
    //    Mantém compatibilidade com a estrutura anterior da tabela
    // ---------------------------------------------------------------
    ColsPos6 = Table.ColumnNames(#"ÍNDICE Adicionado"),
    #"TAG Adicionada" = Table.AddColumn(
        #"ÍNDICE Adicionado", "TAG",
        each Text.Upper(Text.Trim(Text.From(
            if List.Contains(ColsPos6, "EQUIPAMENTO") then [EQUIPAMENTO] else null
        ) ?? "")),
        type text
    ),

    // ---------------------------------------------------------------
    // 7. DESCRICAO GERAL calculada
    //    Formula: FixCase(CATEGORIA) | FixCase(FAMILIA) EQUIPAMENTO
    //    Substitui o valor vindo do CSV (recalculo garante padrao)
    // ---------------------------------------------------------------
    ColsPos7 = Table.ColumnNames(#"TAG Adicionada"),
    #"Desc Calculada" = Table.AddColumn(
        #"TAG Adicionada", "DESCRIÇÃO GERAL",
        each
            let
                cat   = FixCase(if List.Contains(ColsPos7, "CATEGORIA") then [CATEGORIA] else ""),
                fam   = FixCase(if List.Contains(ColsPos7, "FAMILIA")   then [FAMILIA]   else ""),
                equip = Text.Upper(Text.Trim(Text.From([TAG] ?? "")))
            in cat & " | " & fam & " " & equip,
        type text
    ),
    // Remove coluna _DESC_CSV (descrição original do CSV) se existir
    #"Desc Limpa" = if List.Contains(Table.ColumnNames(#"Desc Calculada"), "_DESC_CSV")
        then Table.RemoveColumns(#"Desc Calculada", {"_DESC_CSV"})
        else #"Desc Calculada",

    // ---------------------------------------------------------------
    // 8. DADOS ENGENHARIA -> POTENCIA / CORRENTE / TENSAO
    //    Formato esperado no atributo: "valor/valor/valor"
    // ---------------------------------------------------------------
    ColsPos8 = Table.ColumnNames(#"Desc Limpa"),
    #"Engenharia Dividida" = Table.AddColumn(
        #"Desc Limpa", "EngSplit",
        each
            let
                texto  = Text.From(
                    if List.Contains(ColsPos8, "DADOS ENGENHARIA")
                    then [DADOS ENGENHARIA] else null
                ) ?? "",
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
    // 9. TRATAMENTO DE TEXTO — FixCase nas colunas de texto
    //    Aplica apenas nas colunas que existirem no CSV
    // ---------------------------------------------------------------
    ColsTratamento = {
        "CATEGORIA", "SETOR", "FAMILIA", "ACIONAMENTO", "MODELO", "CAIXA DE PASSAGEM"
    },
    ColsParaTratar = List.Intersect({
        Table.ColumnNames(#"Colunas Engenharia"),
        ColsTratamento
    }),
    #"Tratamento Texto" = Table.TransformColumns(
        #"Colunas Engenharia",
        List.Transform(ColsParaTratar, each {_, each FixCase(_)})
    ),

    // ---------------------------------------------------------------
    // 10. COLUNAS MANUAIS (Excel-only)
    //     Preenchidas pelo usuario na tabela — nao sincronizadas com CAD
    // ---------------------------------------------------------------
    #"Colunas Extras" = Table.AddColumn(
        Table.AddColumn(
            Table.AddColumn(#"Tratamento Texto", "INTERFACE", each null),
            "CONEXÃO", each null
        ),
        "ORIGEM", each null
    ),

    // ---------------------------------------------------------------
    // 11. ORDENACAO
    // ---------------------------------------------------------------
    #"Ordenado" = Table.Sort(#"Colunas Extras", {
        {"ÍNDICE",    Order.Ascending},
        {"CATEGORIA", Order.Ascending}
    }),

    // ---------------------------------------------------------------
    // 12. SELECAO FINAL — mesma estrutura da versao com DATAEXTRACTION
    // ---------------------------------------------------------------
    ColsFinais = {
        "ÍNDICE", "HANDLE", "EQUIPAMENTO", "EQUIPAMENTO AUX", "SETOR", "FAMILIA",
        "CATEGORIA", "TAG", "DESCRIÇÃO GERAL", "ACIONAMENTO", "INTERFACE",
        "POTÊNCIA", "CORRENTE", "TENSÃO", "MODELO", "CAIXA DE PASSAGEM", "CONEXÃO", "ORIGEM"
    },
    ColsDisponiveis = List.Intersect({ColsFinais, Table.ColumnNames(#"Ordenado")}),
    #"Seleção Final" = Table.SelectColumns(#"Ordenado", ColsDisponiveis)
in
    #"Seleção Final"
