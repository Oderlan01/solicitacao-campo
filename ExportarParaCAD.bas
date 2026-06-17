Attribute VB_Name = "Modulo_ExportarParaCAD"
Option Explicit

' ================================================================
' ExportarParaCAD
' Exporta os campos editaveis da Tabela_Componentes para
' Desktop\campos_manuais.csv, pronto para ser importado no
' AutoCAD com o comando ImportarCamposManuais (automacao.lsp).
'
' Regras:
'   - Apenas colunas com atributo correspondente no AutoCAD
'     sao incluidas no CSV.
'   - Colunas Excel-only (INTERFACE, CONEXAO, ORIGEM) sao
'     ignoradas automaticamente.
'   - Apenas a primeira ocorrencia de cada HANDLE e exportada
'     (linha pai). Linhas filhos compartilham o mesmo HANDLE
'     e sao puladas para evitar sobreescrita indevida.
'
' Como usar:
'   1. Edite os valores na aba Componentes
'   2. Execute esta macro (botao ou Alt+F8 -> ExportarParaCAD)
'   3. No AutoCAD, execute o comando: ImportarCamposManuais
' ================================================================

Sub ExportarParaCAD()

    Dim ws       As Worksheet
    Dim lo       As ListObject
    Dim caminho  As String
    Dim numArq   As Integer
    Dim i        As Long
    Dim c        As Long
    Dim header   As String
    Dim linha    As String
    Dim handle   As String
    Dim val      As String
    Dim exportados As Long
    Dim semHandle  As Long

    ' --- Localizar a tabela ---
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Componentes")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Aba 'Componentes' nao encontrada.", vbCritical, "Erro"
        Exit Sub
    End If

    On Error Resume Next
    Set lo = ws.ListObjects("Tabela_Componentes")
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Tabela 'Tabela_Componentes' nao encontrada.", vbCritical, "Erro"
        Exit Sub
    End If

    ' --- Mapeamento: nome da coluna em Componentes -> tag do atributo no AutoCAD ---
    ' Adicione ou remova entradas conforme os atributos dos seus blocos.
    Dim mapa As Object
    Set mapa = CreateObject("Scripting.Dictionary")
    mapa.Add "ACIONAMENTO",       "0H_ACIONAMENTO"
    mapa.Add "MODELO",            "0D_MODELO"
    mapa.Add "CATEGORIA",         "0A_CATEGORIA"
    mapa.Add "SETOR",             "0B_SETOR"
    mapa.Add "FAMILIA",           "0C_FAMILIA"
    mapa.Add "EQUIPAMENTO AUX",   "0F_TAG_AUXILIAR"
    mapa.Add "CAIXA DE PASSAGEM", "CAIXA_DE_PASSAGEM"
    mapa.Add "DESCRICAO GERAL",   "0G_DESCRICAO_GERAL"
    mapa.Add "DESCRI" & Chr(199) & Chr(195) & "O GERAL", "0G_DESCRI" & Chr(199) & Chr(195) & "O_GERAL"
    ' Chr(199)=C cedilha, Chr(195)=A til  ->  "DESCRICAO"/"DESCRIÇÃO"

    ' --- Descobrir quais colunas do mapa existem na tabela ---
    Dim nomesCols() As String
    Dim tagsCols()  As String
    Dim idxCols()   As Long
    Dim nCols       As Long
    nCols = 0

    Dim chave As Variant
    Dim colTmp As ListColumn
    For Each chave In mapa.Keys
        Set colTmp = Nothing
        On Error Resume Next
        Set colTmp = lo.ListColumns(CStr(chave))
        On Error GoTo 0
        If Not colTmp Is Nothing Then
            ReDim Preserve nomesCols(nCols)
            ReDim Preserve tagsCols(nCols)
            ReDim Preserve idxCols(nCols)
            nomesCols(nCols) = CStr(chave)
            tagsCols(nCols)  = CStr(mapa(chave))
            idxCols(nCols)   = colTmp.Index
            nCols = nCols + 1
        End If
    Next chave

    If nCols = 0 Then
        MsgBox "Nenhuma coluna mapeavel encontrada na tabela." & vbNewLine & _
               "Verifique o mapeamento no codigo VBA.", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' --- Coluna HANDLE (obrigatoria) ---
    Dim colHandle As Long
    colHandle = 0
    On Error Resume Next
    colHandle = lo.ListColumns("HANDLE").Index
    On Error GoTo 0
    If colHandle = 0 Then
        MsgBox "Coluna HANDLE nao encontrada em Tabela_Componentes." & vbNewLine & _
               "Atualize o Power Query com o arquivo power-query-componentes.m e recarregue.", _
               vbCritical, "Erro"
        Exit Sub
    End If

    ' --- Gerar CSV ---
    caminho = Environ("USERPROFILE") & "\Desktop\campos_manuais.csv"
    Application.ScreenUpdating = False

    On Error GoTo ErroArquivo
    numArq = FreeFile
    Open caminho For Output As #numArq

    ' Cabecalho: HANDLE seguido dos tags AutoCAD
    header = "HANDLE"
    For c = 0 To nCols - 1
        header = header & ";" & tagsCols(c)
    Next c
    Print #numArq, header

    ' Dados: uma linha por bloco (apenas primeira ocorrencia do HANDLE)
    Dim handlesVistos As Object
    Set handlesVistos = CreateObject("Scripting.Dictionary")
    exportados = 0
    semHandle  = 0

    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            handle = Trim(CStr(lo.DataBodyRange.Cells(i, colHandle).Value))

            If handle = "" Then
                semHandle = semHandle + 1
            ElseIf Not handlesVistos.Exists(handle) Then
                handlesVistos.Add handle, True
                linha = handle
                For c = 0 To nCols - 1
                    val = Trim(CStr(lo.DataBodyRange.Cells(i, idxCols(c)).Value))
                    ' Substituir ponto-e-virgula dentro do valor para nao quebrar o CSV
                    val = Replace(val, ";", ",")
                    linha = linha & ";" & val
                Next c
                Print #numArq, linha
                exportados = exportados + 1
            End If
        Next i
    End If

    Close #numArq
    Application.ScreenUpdating = True

    Dim msg As String
    msg = exportados & " bloco(s) exportado(s) para:" & vbNewLine & caminho
    If semHandle > 0 Then
        msg = msg & vbNewLine & vbNewLine & _
              "Atencao: " & semHandle & " linha(s) sem HANDLE ignoradas" & vbNewLine & _
              "(linhas filhos compartilham o handle do bloco pai)."
    End If
    MsgBox msg, vbInformation, "Exportar para CAD"
    Exit Sub

ErroArquivo:
    Close #numArq
    Application.ScreenUpdating = True
    MsgBox "Erro ao gravar o arquivo:" & vbNewLine & Err.Description, vbCritical, "Erro"
End Sub
