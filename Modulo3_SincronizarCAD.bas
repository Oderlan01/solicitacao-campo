Attribute VB_Name = "Modulo3_SincronizarCAD"
Option Explicit

' ================================================================
' MODULO 3 — SincronizarCAD
'
' Duas macros bidirecionais usando UM UNICO arquivo:
'   Desktop\todos_atributos.csv
'
' ImportarDoCAD:   le CSV exportado pelo AutoCAD -> atualiza Excel
' ExportarParaCAD: le Excel -> atualiza CSV -> AutoCAD importa
'
' Mapeamento central (coluna Excel <-> tag AutoCAD):
'   Edite ObterMapa() para adicionar/remover colunas
' ================================================================

' --- Mapeamento central bidirecional ---
' Chave = nome da coluna na Tabela_Componentes (Excel)
' Valor = nome do atributo no bloco AutoCAD (= coluna no CSV)
Function ObterMapa() As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    ' Apenas campos RW (round-trip). Alinhado ao dicionario em FLUXO_DADOS.md.
    ' DESCRICAO GERAL e DADOS ENGENHARIA sao derivados (R) e NAO entram aqui,
    ' para nao corromper o bloco na volta ao CAD.
    m.Add "CATEGORIA",         "0A_CATEGORIA"
    m.Add "SETOR",             "0B_SETOR"
    m.Add "FAMILIA",           "0C_FAMILIA"
    m.Add "MODELO",            "0D_MODELO"
    m.Add "TAG",               "0E_TAG"
    m.Add "EQUIPAMENTO AUX",   "0F_TAG_AUXILIAR"
    m.Add "ACIONAMENTO",       "0H_ACIONAMENTO"
    m.Add "SENSOR",            "0J_SENSOR"
    m.Add "CAIXA DE PASSAGEM", "CAIXA_DE_PASSAGEM"
    m.Add "VALVULA ABRE",      "VALVULA_ABRE"
    Set ObterMapa = m
End Function

' Helper: remove \r do final da string (CRLF do Windows/Excel)
Private Function StripCR(s As String) As String
    If Len(s) > 0 And Right(s, 1) = Chr(13) Then
        StripCR = Left(s, Len(s) - 1)
    Else
        StripCR = s
    End If
End Function

' ----------------------------------------------------------------
' IMPORTAR DO CAD -> EXCEL
' Le Desktop\todos_atributos.csv (exportado pelo AutoCAD)
' e atualiza Tabela_Componentes pelo HANDLE.
' Celulas atualizadas ficam em verde para identificacao.
' ----------------------------------------------------------------
Sub ImportarDoCAD()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim mapa As Object
    Dim mapaReverso As Object
    Dim caminho As String
    Dim numArq As Integer
    Dim cabecalho As String
    Dim linha As String
    Dim listaCols() As String
    Dim colHandle As Long
    Dim colNome As Long
    Dim i As Long, c As Long
    Dim atualizados As Long
    Dim naoEncontrados As Long

    caminho = Environ("USERPROFILE") & "\Desktop\todos_atributos.csv"

    If Dir(caminho) = "" Then
        MsgBox "Arquivo nao encontrado:" & vbNewLine & caminho & vbNewLine & vbNewLine & _
               "Execute primeiro o comando STWExportar no AutoCAD.", _
               vbExclamation, "Arquivo nao encontrado"
        Exit Sub
    End If

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

    colHandle = 0 : colNome = 0
    On Error Resume Next
    colHandle = lo.ListColumns("HANDLE").Index
    colNome   = lo.ListColumns("NOME_DO_BLOCO").Index
    On Error GoTo 0
    If colHandle = 0 Then
        MsgBox "Coluna HANDLE nao encontrada em Tabela_Componentes.", vbCritical, "Erro"
        Exit Sub
    End If

    ' Mapa reverso: tag AutoCAD -> indice da coluna Excel
    Set mapa = ObterMapa()
    Set mapaReverso = CreateObject("Scripting.Dictionary")
    Dim chave As Variant
    Dim colTmp As ListColumn
    For Each chave In mapa.Keys
        Set colTmp = Nothing
        On Error Resume Next
        Set colTmp = lo.ListColumns(CStr(chave))
        On Error GoTo 0
        If Not colTmp Is Nothing Then
            mapaReverso(CStr(mapa(chave))) = colTmp.Index
        End If
    Next chave

    ' Indice HANDLE -> linha na tabela Excel
    Dim dictHandle As Object
    Set dictHandle = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            Dim hVal As String
            hVal = UCase(StripCR(Trim(CStr(lo.DataBodyRange.Cells(i, colHandle).Value))))
            If hVal <> "" And Not dictHandle.Exists(hVal) Then
                dictHandle(hVal) = i
            End If
        Next i
    End If

    Application.ScreenUpdating = False
    atualizados = 0 : naoEncontrados = 0

    On Error GoTo ErroLeitura
    numArq = FreeFile
    Open caminho For Input As #numArq

    Line Input #numArq, cabecalho
    cabecalho = StripCR(cabecalho)
    listaCols = Split(cabecalho, ";")

    Do While Not EOF(numArq)
        Line Input #numArq, linha
        linha = StripCR(linha)
        If linha = "" Then GoTo ProximaLinhaImp

        Dim dados() As String
        dados = Split(linha, ";")
        If UBound(dados) < 1 Then GoTo ProximaLinhaImp

        Dim handle As String
        handle = UCase(StripCR(Trim(dados(0))))

        If Not dictHandle.Exists(handle) Then
            naoEncontrados = naoEncontrados + 1
            GoTo ProximaLinhaImp
        End If

        Dim rowIdx As Long
        rowIdx = dictHandle(handle)

        ' Atualiza NOME_DO_BLOCO se disponivel
        If colNome > 0 And UBound(dados) >= 1 Then
            lo.DataBodyRange.Cells(rowIdx, colNome).Value = StripCR(Trim(dados(1)))
        End If

        ' Atualiza colunas mapeadas
        For c = 2 To UBound(dados)
            If c <= UBound(listaCols) Then
                Dim tagName As String
                tagName = StripCR(listaCols(c))
                If mapaReverso.Exists(tagName) Then
                    Dim colIdx As Long
                    colIdx = mapaReverso(tagName)
                    Dim val As String
                    val = StripCR(Trim(dados(c)))
                    If val <> "" Then
                        Dim cel As Range
                        Set cel = lo.DataBodyRange.Cells(rowIdx, colIdx)
                        If StripCR(Trim(CStr(cel.Value))) <> val Then
                            cel.Value = val
                            cel.Font.Color = RGB(0, 128, 0)
                            atualizados = atualizados + 1
                        End If
                    End If
                End If
            End If
        Next c

ProximaLinhaImp:
    Loop

    Close #numArq
    Application.ScreenUpdating = True

    Dim msgImp As String
    msgImp = atualizados & " campo(s) atualizado(s) do AutoCAD." & vbNewLine & _
             "Celulas em verde = atualizadas nesta importacao."
    If naoEncontrados > 0 Then
        msgImp = msgImp & vbNewLine & vbNewLine & _
                 naoEncontrados & " handle(s) do CSV nao encontrado(s) na tabela." & vbNewLine & _
                 "(Blocos novos no CAD — rode Sincronizar para atualizar a lista.)"
    End If
    MsgBox msgImp, vbInformation, "Importar do CAD"
    Exit Sub

ErroLeitura:
    Close #numArq
    Application.ScreenUpdating = True
    MsgBox "Erro ao ler o arquivo: " & Err.Description, vbCritical, "Erro"
End Sub

' ----------------------------------------------------------------
' EXPORTAR EXCEL -> CAD  (atualiza o CSV existente em memoria)
' Le Desktop\todos_atributos.csv, atualiza os campos mapeados
' pelo HANDLE e grava de volta SEM alterar estrutura ou colunas.
' Em seguida rode STWImportar no AutoCAD para aplicar.
' ----------------------------------------------------------------
Sub ExportarParaCAD()
    Dim ws As Worksheet, lo As ListObject
    Dim mapa As Object
    Dim caminho As String, numArq As Integer
    Dim i As Long, j As Long
    Dim atualizados As Long, naoEncontrados As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Componentes")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Aba 'Componentes' nao encontrada.", vbCritical, "Erro": Exit Sub

    On Error Resume Next
    Set lo = ws.ListObjects("Tabela_Componentes")
    On Error GoTo 0
    If lo Is Nothing Then MsgBox "Tabela 'Tabela_Componentes' nao encontrada.", vbCritical, "Erro": Exit Sub

    caminho = Environ("USERPROFILE") & "\Desktop\todos_atributos.csv"
    If Dir(caminho) = "" Then
        MsgBox "Arquivo nao encontrado:" & vbNewLine & caminho & vbNewLine & vbNewLine & _
               "Execute primeiro o comando STWExportar no AutoCAD.", vbCritical, "Erro"
        Exit Sub
    End If

    Set mapa = ObterMapa()

    ' --- 1. Ler o CSV inteiro em memoria ---
    Dim buf() As String
    Dim totalLinhas As Long: totalLinhas = 0
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
        MsgBox "CSV vazio ou sem dados.", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' --- 2. Parsear cabecalho -> dicionario nome_coluna: indice ---
    Dim cabCols() As String
    cabCols = Split(buf(0), ";")
    Dim nCsvCols As Long: nCsvCols = UBound(cabCols) + 1

    Dim idxCsvCol As Object
    Set idxCsvCol = CreateObject("Scripting.Dictionary")
    idxCsvCol.CompareMode = vbTextCompare
    For j = 0 To UBound(cabCols)
        Dim nomeCol As String: nomeCol = StripCR(Trim(cabCols(j)))
        cabCols(j) = nomeCol
        If nomeCol <> "" And Not idxCsvCol.Exists(nomeCol) Then idxCsvCol(nomeCol) = j
    Next j

    If Not idxCsvCol.Exists("HANDLE_CAD") Then
        MsgBox "Coluna HANDLE_CAD nao encontrada no CSV." & vbNewLine & _
               "Verifique se o arquivo foi gerado pelo comando STWExportar do AutoCAD.", vbCritical, "Erro"
        Exit Sub
    End If
    Dim idxHandleCSV As Long: idxHandleCSV = idxCsvCol("HANDLE_CAD")

    ' --- 3. Dicionario HANDLE -> indice da linha no buffer ---
    Dim handleParaLinha As Object
    Set handleParaLinha = CreateObject("Scripting.Dictionary")
    handleParaLinha.CompareMode = vbTextCompare
    Dim celulas() As String
    For i = 1 To totalLinhas - 1
        celulas = Split(buf(i), ";")
        If UBound(celulas) >= idxHandleCSV Then
            Dim hCSV As String: hCSV = UCase(StripCR(Trim(celulas(idxHandleCSV))))
            If hCSV <> "" And Not handleParaLinha.Exists(hCSV) Then handleParaLinha(hCSV) = i
        End If
    Next i

    ' --- 4. Coluna HANDLE na tabela Excel ---
    Dim colHandle As Long: colHandle = 0
    On Error Resume Next
    colHandle = lo.ListColumns("HANDLE").Index
    On Error GoTo 0
    If colHandle = 0 Then
        MsgBox "Coluna HANDLE nao encontrada em Tabela_Componentes.", vbCritical, "Erro"
        Exit Sub
    End If

    ' --- 5. Pares validos (coluna Excel <-> coluna CSV) ---
    Dim chavesExcel() As String, idxExcelCol() As Long, idxCsvDest() As Long
    Dim nPares As Long: nPares = 0
    Dim chave As Variant, colTmp As ListColumn
    For Each chave In mapa.Keys
        Set colTmp = Nothing
        On Error Resume Next
        Set colTmp = lo.ListColumns(CStr(chave))
        On Error GoTo 0
        If Not colTmp Is Nothing Then
            Dim tagDest As String: tagDest = CStr(mapa(chave))
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
        MsgBox "Nenhum par de colunas encontrado." & vbNewLine & _
               "Verifique o mapeamento em ObterMapa() e as colunas do CSV.", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' --- 6. Percorrer Componentes e atualizar buffer ---
    Application.ScreenUpdating = False
    Dim handlesVistos As Object
    Set handlesVistos = CreateObject("Scripting.Dictionary")
    handlesVistos.CompareMode = vbTextCompare
    atualizados = 0 : naoEncontrados = 0

    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            Dim hExcel As String
            hExcel = UCase(StripCR(Trim(CStr(lo.DataBodyRange.Cells(i, colHandle).Value))))
            If hExcel = "" Then GoTo ProximaLinha
            If handlesVistos.Exists(hExcel) Then GoTo ProximaLinha
            handlesVistos(hExcel) = True

            If Not handleParaLinha.Exists(hExcel) Then
                naoEncontrados = naoEncontrados + 1
                GoTo ProximaLinha
            End If

            Dim idxRow As Long: idxRow = handleParaLinha(hExcel)
            celulas = Split(buf(idxRow), ";")
            If UBound(celulas) + 1 < nCsvCols Then ReDim Preserve celulas(nCsvCols - 1)

            Dim p As Long
            For p = 0 To nPares - 1
                Dim valExcel As String
                valExcel = StripCR(Trim(CStr(lo.DataBodyRange.Cells(i, idxExcelCol(p)).Value)))
                If valExcel <> "" Then
                    valExcel = Replace(valExcel, ";", ",")
                    Dim valAtual As String: valAtual = StripCR(celulas(idxCsvDest(p)))
                    If valExcel <> valAtual Then
                        celulas(idxCsvDest(p)) = valExcel
                        atualizados = atualizados + 1
                    End If
                End If
            Next p
            buf(idxRow) = Join(celulas, ";")

ProximaLinha:
        Next i
    End If

    ' --- 7. Gravar CSV de volta ---
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
              naoEncontrados & " handle(s) da planilha nao encontrado(s) no CSV."
    End If
    msg = msg & vbNewLine & vbNewLine & "Execute o comando STWImportar no AutoCAD."
    MsgBox msg, vbInformation, "Exportar para CAD"
    Exit Sub

ErroArquivo:
    On Error GoTo 0
    Close #numArq
    Application.ScreenUpdating = True
    MsgBox "Erro ao acessar o arquivo: " & Err.Description, vbCritical, "Erro"
End Sub
