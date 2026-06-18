// ================================================================
// Power Query — Aba DADOS_MES3 (fonte da etiqueta MES3)
//
// Le o Summary do .xlsx do projeto (gerado por DATAEXTRACTION / "export
// dados" do AutoCAD) e extrai apenas os campos do grupo MES3 do schema:
//   0M_NAME           -> name
//   0L_TAG_AMOUNT     -> tag_amount
//   0K_TAG_SIZE       -> tag_size
//   0E_TAG            -> tag
//   0G_DESCRICAO_GERAL-> description
// Adiciona id / client_id / code (preenchidos pelo VBA AtualizarMES3_Sequencial,
// que mantem id sequencial persistente na tabela MES3/Tabela3).
//
// Ajuste CaminhoFonte para o .xlsx do projeto corrente.
// Ver FLUXO_DADOS.md (RN3, grupo MES3).
// ================================================================
let
    CaminhoFonte = "\\10.1.1.3\Servidor\Clientes\Safeeds - Cascavel-PR\7540.5.0 - Implementação Líquidos, Micro Encapsulados e Semi Sólidos\2 - DOCUMENTAÇÃO\ELÉTRICA\Documentos\Fluxogramas e P&ID\Safeeds_Liquidos.xlsx",
    Fonte = Excel.Workbook(File.Contents(CaminhoFonte), null, true),
    Summary_Sheet = Fonte{[Item="Summary",Kind="Sheet"]}[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Summary_Sheet, [PromoteAllScalars=true]),
    #"Outras Colunas Removidas" = Table.SelectColumns(#"Cabeçalhos Promovidos",{"0M_NAME", "0L_TAG_AMOUNT", "0K_TAG_SIZE", "0E_TAG", "0G_DESCRIÇÃO_GERAL"}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Outras Colunas Removidas",{{"0M_NAME", "name"}, {"0L_TAG_AMOUNT", "tag_amount"}, {"0K_TAG_SIZE", "tag_size"}, {"0E_TAG", "tag"}, {"0G_DESCRIÇÃO_GERAL", "description"}}),
    #"Personalização Adicionada" = Table.AddColumn(#"Colunas Renomeadas", "id", each null),
    #"Personalização Adicionada1" = Table.AddColumn(#"Personalização Adicionada", "client_id", each null),
    #"Personalização Adicionada2" = Table.AddColumn(#"Personalização Adicionada1", "code", each null),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Personalização Adicionada2",{"id", "name", "description", "tag", "tag_size", "tag_amount", "client_id", "code"})
in
    #"Colunas Reordenadas"
