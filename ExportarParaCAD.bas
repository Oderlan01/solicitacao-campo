Attribute VB_Name = "Modulo_ExportarParaCAD"
Option Explicit

' ================================================================
' AtualizarBaseParaCAD
' Le o Desktop\todos_atributos.csv (gerado pelo c:ExportarTodos do
' AutoCAD), atualiza apenas os campos mapeados usando a coluna
' HANDLE como chave, e grava o arquivo de volta sem alterar sua
' estrutura ou colunas.
'
' Fluxo:
'   1. Edicao dos valores em Componentes (Tabela_Componentes)
'   2. Executar esta macro (botao ou Alt+F8 -> AtualizarBaseParaCAD)
'   3. No AutoCAD, executar o comando: ImportarTodos
'
' Pre-requisito: o arquivo todos_atributos.csv deve existir no
' Desktop, gerado previamente pelo comando ExportarTodos no AutoCAD.
' ================================================================

Sub AtualizarBaseParaCAD()

    Dim ws          As Worksheet
    Dim lo          As ListObject
    Dim caminho     As String
    Dim numArq      As Integer
    Dim i           As Long
    Dim j           As Long

    ' --- Localizar a tabela Componentes ---
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

    ' --- Caminho do CSV base gerado pelo AutoCAD ---
    caminho = Environ("USERPROFILE") & "\Desktop\todos_atributos.csv"
    If Dir(caminho) = "" Then
        MsgBox "Arquivo nao encontrado:" & vbNewLine & caminho & vbNewLine & vbNewLine & _
               "Execute primeiro o comando ExportarTodos no AutoCAD.", vbCritical, "Erro"
        Exit Sub
    End If

    ' --- Mapeamento: coluna em Componentes -> tag AutoCAD (= coluna no CSV) ---
    Dim mapa As Object
    Set mapa = CreateObject("Scripting.Dictionary")
    mapa.CompareMode = vbTextCompare
    mapa.Add "ACIONAMENTO",       "0H_ACIONAMENTO"
    mapa.Add "MODELO",            "0D_MODELO"
    mapa.Add "CATEGORIA",         "0A_CATEGORIA"
    mapa.Add "SETOR",             "0B_SETOR"
    mapa.Add "FAMILIA",           "0C_FAMILIA"
    mapa.Add "EQUIPAMENTO AUX",   "0F_TAG_AUXILIAR"
    mapa.Add "CAIXA DE PASSAGEM", "CAIXA_DE_PASSAGEM"
    ' DESCRICAO com e sem acentuacao (Chr(199)=C-cedilha, Chr(195)=A-til)
    mapa.Add "DESCRICAO GERAL",                           "0G_DESCRICAO_GERAL"
    mapa.Add "DESCRI" & Chr(199) & Chr(195) & "O GERAL", "0G_DESCRI" & Chr(199) & Chr(195) & "O_GERAL"

    Application.ScreenUpdating = False

    ' -------------------------------------------------------
    ' PASSO 1: Ler o CSV completo em um array de strings
    ' -------------------------------------------------------
    Dim buf()       As String
    Dim totalLinhas As Long
    totalLinhas = 0
    ReDim buf(0)

    On Error GoTo ErroArquivo
    numArq = FreeFile
    Open caminho For Input As #numArq
    Dim txtLinha As String
    Do While Not EOF(numArq)
        Line Input #numArq, txtLinha
        ReDim Preserve buf(totalLinhas)
        buf(totalLinhas) = txtLinha
        totalLinhas = totalLinhas + 1
    Loop
    Close #numArq
    On Error GoTo 0

    If totalLinhas < 2 Then
        MsgBox "Arquivo CSV vazio ou sem dados:" & vbNewLine & caminho, vbExclamation, "Aviso"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' -------------------------------------------------------
    ' PASSO 2: Parsear cabecalho e mapear nome -> indice
    ' -------------------------------------------------------
    Dim cabCols()  As String
    cabCols = Split(buf(0), ";")
    Dim nCsvCols   As Long
    nCsvCols = UBound(cabCols) + 1

    ' Limpar \r das celulas do cabecalho e construir dicionario
    Dim idxCsvCol As Object
    Set idxCsvCol = CreateObject("Scripting.Dictionary")
    idxCsvCol.CompareMode = vbTextCompare
    For j = 0 To UBound(cabCols)
        Dim nomeCol As String
        nomeCol = Trim(cabCols(j))
        If Len(nomeCol) > 0 And Right(nomeCol, 1) = Chr(13) Then
            nomeCol = Left(nomeCol, Len(nomeCol) - 1)
        End If
        cabCols(j) = nomeCol
        If nomeCol <> "" And Not idxCsvCol.Exists(nomeCol) Then
            idxCsvCol.Add nomeCol, j
        End If
    Next j

    If Not idxCsvCol.Exists("HANDLE_CAD") Then
        MsgBox "Coluna HANDLE_CAD nao encontrada no CSV." & vbNewLine & _
               "Verifique se o arquivo foi gerado pelo comando ExportarTodos do AutoCAD.", _
               vbCritical, "Erro"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Dim idxHandleCSV As Long
    idxHandleCSV = idxCsvCol("HANDLE_CAD")

    ' -------------------------------------------------------
    ' PASSO 3: Construir dicionario HANDLE -> indice da linha
    ' -------------------------------------------------------
    Dim handleParaLinha As Object
    Set handleParaLinha = CreateObject("Scripting.Dictionary")
    handleParaLinha.CompareMode = vbTextCompare

    Dim celulas() As String
    For i = 1 To totalLinhas - 1
        celulas = Split(buf(i), ";")
        If UBound(celulas) >= idxHandleCSV Then
            Dim hndCSV As String
            hndCSV = UCase(Trim(celulas(idxHandleCSV)))
            If hndCSV <> "" And Not handleParaLinha.Exists(hndCSV) Then
                handleParaLinha.Add hndCSV, i
            End If
        End If
    Next i

    ' -------------------------------------------------------
    ' PASSO 4: Descobrir coluna HANDLE na tabela Componentes
    ' -------------------------------------------------------
    Dim colHandleExcel As Long
    colHandleExcel = 0
    On Error Resume Next
    colHandleExcel = lo.ListColumns("HANDLE").Index
    On Error GoTo 0
    If colHandleExcel = 0 Then
        MsgBox "Coluna HANDLE nao encontrada em Tabela_Componentes." & vbNewLine & _
               "Atualize o Power Query com power-query-componentes.m e recarregue.", _
               vbCritical, "Erro"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' -------------------------------------------------------
    ' PASSO 5: Montar lista dos pares (Excel col, CSV col)
    '          presentes nos dois lados
    ' -------------------------------------------------------
    Dim chavesExcel() As String
    Dim idxExcelCol() As Long
    Dim idxCsvDest()  As Long
    Dim nPares        As Long
    nPares = 0

    Dim chave  As Variant
    Dim colTmp As ListColumn
    For Each chave In mapa.Keys
        Set colTmp = Nothing
        On Error Resume Next
        Set colTmp = lo.ListColumns(CStr(chave))
        On Error GoTo 0
        If Not colTmp Is Nothing Then
            Dim tagDest As String
            tagDest = CStr(mapa(chave))
            If idxCsvCol.Exists(tagDest) Then
                ReDim Preserve chavesExcel(nPares)
                ReDim Preserve idxExcelCol(nPares)
                ReDim Preserve idxCsvDest(nPares)
                chavesExcel(nPares) = CStr(chave)
                idxExcelCol(nPares) = colTmp.Index
                idxCsvDest(nPares)  = idxCsvCol(tagDest)
                nPares = nPares + 1
            End If
        End If
    Next chave

    If nPares = 0 Then
        MsgBox "Nenhum par de colunas encontrado nos dois arquivos." & vbNewLine & _
               "Verifique o mapeamento no codigo VBA e as colunas do CSV.", _
               vbExclamation, "Aviso"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' -------------------------------------------------------
    ' PASSO 6: Percorrer Componentes e atualizar o buffer CSV
    ' -------------------------------------------------------
    Dim handlesVistos As Object
    Set handlesVistos = CreateObject("Scripting.Dictionary")
    handlesVistos.CompareMode = vbTextCompare

    Dim atualizados   As Long
    Dim naoEncontrados As Long
    atualizados    = 0
    naoEncontrados = 0

    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            Dim hndExcel As String
            hndExcel = UCase(Trim(CStr(lo.DataBodyRange.Cells(i, colHandleExcel).Value)))

            If hndExcel = "" Then GoTo ProximaLinha
            If handlesVistos.Exists(hndExcel) Then GoTo ProximaLinha
            handlesVistos.Add hndExcel, True

            If Not handleParaLinha.Exists(hndExcel) Then
                naoEncontrados = naoEncontrados + 1
                GoTo ProximaLinha
            End If

            Dim idxLinhaCsv As Long
            idxLinhaCsv = handleParaLinha(hndExcel)

            ' Dividir a linha CSV em celulas
            celulas = Split(buf(idxLinhaCsv), ";")

            ' Garantir tamanho suficiente
            If UBound(celulas) + 1 < nCsvCols Then
                ReDim Preserve celulas(nCsvCols - 1)
            End If

            ' Aplicar cada campo mapeado
            Dim p As Long
            For p = 0 To nPares - 1
                Dim valExcel As String
                valExcel = Trim(CStr(lo.DataBodyRange.Cells(i, idxExcelCol(p)).Value))
                If valExcel <> "" Then
                    valExcel = Replace(valExcel, ";", ",")   ' protege o delimitador CSV
                    valExcel = Replace(valExcel, Chr(13), "") ' remove \r residual

                    ' Valor atual da celula CSV (limpo de \r)
                    Dim valAtual As String
                    valAtual = celulas(idxCsvDest(p))
                    If Len(valAtual) > 0 And Right(valAtual, 1) = Chr(13) Then
                        valAtual = Left(valAtual, Len(valAtual) - 1)
                    End If

                    If valExcel <> valAtual Then
                        celulas(idxCsvDest(p)) = valExcel
                        atualizados = atualizados + 1
                    End If
                End If
            Next p

            ' Remontar a linha no buffer
            buf(idxLinhaCsv) = Join(celulas, ";")

ProximaLinha:
        Next i
    End If

    ' -------------------------------------------------------
    ' PASSO 7: Gravar o CSV de volta (mesma estrutura)
    ' -------------------------------------------------------
    On Error GoTo ErroArquivo
    numArq = FreeFile
    Open caminho For Output As #numArq
    For i = 0 To totalLinhas - 1
        Print #numArq, buf(i)
    Next i
    Close #numArq
    On Error GoTo 0

    Application.ScreenUpdating = True

    Dim msg As String
    msg = atualizados & " campo(s) atualizado(s) em:" & vbNewLine & caminho
    If naoEncontrados > 0 Then
        msg = msg & vbNewLine & vbNewLine & _
              "Atencao: " & naoEncontrados & " handle(s) da planilha nao encontrado(s) no CSV."
    End If
    msg = msg & vbNewLine & vbNewLine & "Execute o comando ImportarTodos no AutoCAD."
    MsgBox msg, vbInformation, "Atualizar Base para CAD"
    Exit Sub

ErroArquivo:
    On Error GoTo 0
    Close #numArq
    Application.ScreenUpdating = True
    MsgBox "Erro ao acessar o arquivo:" & vbNewLine & Err.Description, vbCritical, "Erro"
End Sub
