' ================================================================
' MODULOS VBA — Planilha Excel (compoteste.xlsm)
'
' FLUXO UNIFICADO CAD <-> Excel via todos_atributos.csv:
'
'   CAD -> Excel:
'     1. AutoCAD: ExportarTodos   -> Desktop\todos_atributos.csv
'     2. Excel:   ImportarDoCAD   <- Desktop\todos_atributos.csv
'
'   Excel -> CAD:
'     1. Excel:   ExportarParaCAD -> Desktop\todos_atributos.csv
'     2. AutoCAD: ImportarTodos   <- Desktop\todos_atributos.csv
'
' Modulos:
'   Modulo1 — SalvarNaBaseServidor
'   Modulo2 — GravarAlteracaoNaBase + MapearColuna
'   Modulo3 — ImportarDoCAD + ExportarParaCAD + ObterMapa (NOVO)
'   Modulo4 — AtualizarE_Sincronizar_Total2 + MES3 + Backup
' ================================================================


' ================================================================
' MODULO 1 — SalvarNaBaseServidor
' Exporta a tabela "teste" da aba ativa para um arquivo .txt
' ================================================================
Option Explicit

Sub SalvarNaBaseServidor()
    Dim wsOrigem As Worksheet
    Dim tabela As ListObject
    Dim caminhoBase As String
    Dim numArquivo As Integer
    Dim r As Long, c As Long
    Dim linhaTexto As String

    caminhoBase = "C:\Users\oderlan.colcenti\Desktop\teste.txt"
    Set wsOrigem = ThisWorkbook.ActiveSheet

    On Error Resume Next
    Set tabela = wsOrigem.ListObjects("teste")
    On Error GoTo 0

    If tabela Is Nothing Then
        MsgBox "A tabela 'teste' nao foi encontrada nesta aba!", vbCritical, "Erro"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Salvando alteracoes no arquivo de texto..."

    On Error GoTo ErroGravacao
    numArquivo = FreeFile
    Open caminhoBase For Output As #numArquivo

    For r = 1 To tabela.Range.Rows.Count
        linhaTexto = ""
        For c = 1 To tabela.Range.Columns.Count
            If c = 1 Then
                linhaTexto = tabela.Range.Cells(r, c).Text
            Else
                linhaTexto = linhaTexto & vbTab & tabela.Range.Cells(r, c).Text
            End If
        Next c
        Print #numArquivo, linhaTexto
    Next r

    Close #numArquivo
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    MsgBox "Arquivo de texto atualizado com sucesso!", vbInformation, "Sucesso"
    Exit Sub

ErroGravacao:
    Close #numArquivo
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    MsgBox "Erro ao salvar: " & Err.Description, vbCritical, "Erro"
End Sub


' ================================================================
' MODULO 2 — GravarAlteracaoNaBase
' Atualiza linha correspondente no arquivo Base (teste.xlsx)
' com base na celula selecionada na tabela da aba ativa.
' ================================================================

Sub GravarAlteracaoNaBase()
    Dim wbBase As Workbook
    Dim wsBase As Worksheet
    Dim wsAtual As Worksheet
    Dim tabela As ListObject
    Dim indice As Long
    Dim categoria As String
    Dim novoValor As String
    Dim caminhoBase As String
    Dim linhaBase As Long

    caminhoBase = "C:\Users\oderlan.colcenti\Desktop\teste.xlsx"
    Set wsAtual = ThisWorkbook.ActiveSheet

    If wsAtual.ListObjects.Count > 0 Then
        Set tabela = wsAtual.ListObjects(1)
    Else
        Set tabela = Nothing
    End If

    If tabela Is Nothing Then
        MsgBox "Nenhuma tabela formatada encontrada nesta aba!", vbCritical, "Erro"
        Exit Sub
    End If

    If Intersect(ActiveCell, tabela.DataBodyRange) Is Nothing Then
        MsgBox "Clique em uma celula valida da tabela antes de salvar!", vbExclamation, "Aviso"
        Exit Sub
    End If

    On Error GoTo ErroDados
    indice    = Intersect(ActiveCell.EntireRow, tabela.ListColumns("INDICE").Range).Value
    categoria = Intersect(ActiveCell.EntireRow, tabela.ListColumns("CATEGORIA").Range).Value
    novoValor = Intersect(ActiveCell.EntireRow, tabela.ListColumns("TAG").Range).Value
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Abrindo a base e aplicando alteracoes..."

    On Error GoTo ErroGrava
    Set wbBase = Workbooks.Open(caminhoBase)
    Set wsBase = wbBase.Sheets("Summary")
    linhaBase = indice + 1

    Select Case UCase(Trim(categoria))
        Case "VALVULA ABRE":        wsBase.Cells(linhaBase, MapearColuna(wsBase, "VALVULA_ABRE")).Value = novoValor
        Case "VALVULA FECHA":       wsBase.Cells(linhaBase, MapearColuna(wsBase, "VALVULA_FECHA")).Value = novoValor
        Case "SENSOR ABERTO":       wsBase.Cells(linhaBase, MapearColuna(wsBase, "SENSOR_ABERTO")).Value = novoValor
        Case "SENSOR FECHADO":      wsBase.Cells(linhaBase, MapearColuna(wsBase, "SENSOR_FECHADO")).Value = novoValor
        Case "SENSOR NIVEL BAIXO":  wsBase.Cells(linhaBase, MapearColuna(wsBase, "SENSOR_NIVEL_BAIXO")).Value = novoValor
        Case "SENSOR NIVEL ALTO":   wsBase.Cells(linhaBase, MapearColuna(wsBase, "SENSOR_NIVEL_ALTO")).Value = novoValor
        Case Else
            wsBase.Cells(linhaBase, MapearColuna(wsBase, "0E_TAG")).Value = novoValor
    End Select

    Dim pot As String, corr As String, tens As String
    pot  = Intersect(ActiveCell.EntireRow, tabela.ListColumns("POTENCIA").Range).Value
    corr = Intersect(ActiveCell.EntireRow, tabela.ListColumns("CORRENTE").Range).Value
    tens = Intersect(ActiveCell.EntireRow, tabela.ListColumns("TENSAO").Range).Value

    If pot <> "" Or corr <> "" Or tens <> "" Then
        wsBase.Cells(linhaBase, MapearColuna(wsBase, "0I_DADOS_ENGENHARIA")).Value = pot & "/" & corr & "/" & tens
    End If

    wbBase.Close SaveChanges:=True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = ""
    MsgBox "Base atualizada! INDICE " & indice & " -> linha " & linhaBase, vbInformation, "Sucesso"
    Exit Sub

ErroDados:
    MsgBox "Erro ao ler colunas. Verifique: INDICE, CATEGORIA, TAG.", vbCritical, "Erro de Coluna"
    Exit Sub

ErroGrava:
    If Not wbBase Is Nothing Then wbBase.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = ""
    MsgBox "Erro ao salvar: " & Err.Description, vbCritical, "Erro"
End Sub

Function MapearColuna(aba As Worksheet, nomeColuna As String) As Long
    Dim celula As Range
    Set celula = aba.Rows(1).Find(What:=nomeColuna, LookIn:=xlValues, LookAt:=xlWhole)
    If Not celula Is Nothing Then
        MapearColuna = celula.Column
    Else
        MsgBox "Coluna '" & nomeColuna & "' nao encontrada na base!", vbCritical, "Erro"
        End
    End If
End Function


' ================================================================
' MODULO 3 — SincronizarCAD (NOVO - substitui ExportarParaCAD antigo)
'
' Duas macros bidirecionais usando UM UNICO arquivo:
'   Desktop\todos_atributos.csv
'
' ImportarDoCAD:  le CSV exportado pelo AutoCAD -> atualiza Excel
' ExportarParaCAD: le Excel -> grava CSV -> AutoCAD importa
'
' Mapeamento central (coluna Excel <-> tag AutoCAD):
'   Edite ObterMapa() para adicionar/remover colunas
' ================================================================

' --- Mapeamento central bidireccional ---
' Chave   = nome da coluna na Tabela_Componentes (Excel)
' Valor   = nome do atributo no bloco AutoCAD
Private Function ObterMapa() As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    m.Add "TAG",                "0E_TAG"
    m.Add "ACIONAMENTO",        "0H_ACIONAMENTO"
    m.Add "SENSOR",             "0J_SENSOR"
    m.Add "FAMILIA",            "0C_FAMILIA"
    m.Add "SETOR",              "0B_SETOR"
    m.Add "MODELO",             "0D_MODELO"
    m.Add "DESCRICAO GERAL",    "0G_DESCRICAO_GERAL"
    m.Add "EQUIPAMENTO AUX",    "0F_TAG_AUXILIAR"
    m.Add "CAIXA DE PASSAGEM",  "CAIXA_DE_PASSAGEM"
    m.Add "VALVULA ABRE",       "VALVULA_ABRE"
    Set ObterMapa = m
End Function

' Helper: remove \r do final da linha (CRLF Windows)
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
               "Execute primeiro o comando ExportarTodos no AutoCAD.", _
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

    ' Coluna HANDLE na tabela Excel
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
            hVal = Trim(CStr(lo.DataBodyRange.Cells(i, colHandle).Value))
            If hVal <> "" And Not dictHandle.Exists(UCase(hVal)) Then
                dictHandle(UCase(hVal)) = i
            End If
        Next i
    End If

    Application.ScreenUpdating = False
    atualizados = 0 : naoEncontrados = 0

    On Error GoTo ErroLeitura
    numArq = FreeFile
    Open caminho For Input As #numArq

    ' Le cabecalho: HANDLE_CAD;Nome_Bloco;TAG1;TAG2;...
    Line Input #numArq, cabecalho
    cabecalho = StripCR(cabecalho)
    listaCols = Split(cabecalho, ";")

    ' Le linhas de dados
    Do While Not EOF(numArq)
        Line Input #numArq, linha
        linha = StripCR(linha)
        If linha = "" Then GoTo ProximaLinha

        Dim dados() As String
        dados = Split(linha, ";")
        If UBound(dados) < 1 Then GoTo ProximaLinha

        Dim handle As String
        handle = UCase(Trim(dados(0)))

        If Not dictHandle.Exists(handle) Then
            naoEncontrados = naoEncontrados + 1
            GoTo ProximaLinha
        End If

        Dim rowIdx As Long
        rowIdx = dictHandle(handle)

        ' Atualiza NOME_DO_BLOCO se disponivel
        If colNome > 0 And UBound(dados) >= 1 Then
            lo.DataBodyRange.Cells(rowIdx, colNome).Value = Trim(dados(1))
        End If

        ' Atualiza colunas mapeadas
        For c = 2 To UBound(dados)
            If c <= UBound(listaCols) Then
                Dim tagName As String
                tagName = listaCols(c)
                If mapaReverso.Exists(tagName) Then
                    Dim colIdx As Long
                    colIdx = mapaReverso(tagName)
                    Dim val As String
                    val = Trim(dados(c))
                    If val <> "" Then
                        Dim cel As Range
                        Set cel = lo.DataBodyRange.Cells(rowIdx, colIdx)
                        If Trim(CStr(cel.Value)) <> val Then
                            cel.Value = val
                            cel.Font.Color = RGB(0, 128, 0) ' Verde = atualizado do CAD
                            atualizados = atualizados + 1
                        End If
                    End If
                End If
            End If
        Next c

ProximaLinha:
    Loop

    Close #numArq
    Application.ScreenUpdating = True

    Dim msg As String
    msg = atualizados & " campo(s) atualizado(s) do AutoCAD." & vbNewLine & _
          "Celulas em verde = atualizadas nesta importacao."
    If naoEncontrados > 0 Then
        msg = msg & vbNewLine & vbNewLine & _
              naoEncontrados & " handle(s) do CSV nao encontrado(s) na tabela." & vbNewLine & _
              "(Blocos novos no CAD — rode Sincronizar para atualizar a lista.)"
    End If
    MsgBox msg, vbInformation, "Importar do CAD"
    Exit Sub

ErroLeitura:
    Close #numArq
    Application.ScreenUpdating = True
    MsgBox "Erro ao ler o arquivo: " & Err.Description, vbCritical, "Erro"
End Sub

' ----------------------------------------------------------------
' EXPORTAR EXCEL -> CAD
' Le Tabela_Componentes e grava Desktop\todos_atributos.csv
' no mesmo formato do AutoCAD (HANDLE_CAD;Nome_Bloco;TAG1;...)
' Em seguida rode ImportarTodos no AutoCAD para aplicar.
' ----------------------------------------------------------------
Sub ExportarParaCAD()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim mapa As Object
    Dim caminho As String
    Dim numArq As Integer
    Dim i As Long, c As Long
    Dim header As String
    Dim linha As String
    Dim handle As String
    Dim nomeBloco As String
    Dim val As String
    Dim exportados As Long
    Dim semHandle As Long

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

    Set mapa = ObterMapa()

    ' Descobre quais colunas mapeadas existem na tabela
    Dim tagsCols()  As String
    Dim idxCols()   As Long
    Dim nCols As Long: nCols = 0

    Dim chave As Variant
    Dim colTmp As ListColumn
    For Each chave In mapa.Keys
        Set colTmp = Nothing
        On Error Resume Next
        Set colTmp = lo.ListColumns(CStr(chave))
        On Error GoTo 0
        If Not colTmp Is Nothing Then
            ReDim Preserve tagsCols(nCols)
            ReDim Preserve idxCols(nCols)
            tagsCols(nCols) = CStr(mapa(chave))  ' tag AutoCAD no cabecalho
            idxCols(nCols)  = colTmp.Index
            nCols = nCols + 1
        End If
    Next chave

    If nCols = 0 Then
        MsgBox "Nenhuma coluna mapeavel encontrada na tabela.", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' Colunas obrigatorias
    Dim colHandle As Long: colHandle = 0
    Dim colNome As Long:   colNome = 0
    On Error Resume Next
    colHandle = lo.ListColumns("HANDLE").Index
    colNome   = lo.ListColumns("NOME_DO_BLOCO").Index
    On Error GoTo 0
    If colHandle = 0 Then
        MsgBox "Coluna HANDLE nao encontrada em Tabela_Componentes.", vbCritical, "Erro"
        Exit Sub
    End If

    ' Grava CSV
    caminho = Environ("USERPROFILE") & "\Desktop\todos_atributos.csv"
    Application.ScreenUpdating = False

    On Error GoTo ErroArquivo
    numArq = FreeFile
    Open caminho For Output As #numArq

    ' Cabecalho identico ao gerado pelo AutoCAD
    header = "HANDLE_CAD;Nome_Bloco"
    For c = 0 To nCols - 1
        header = header & ";" & tagsCols(c)
    Next c
    Print #numArq, header

    ' Linhas de dados — uma por HANDLE unico
    Dim handlesVistos As Object
    Set handlesVistos = CreateObject("Scripting.Dictionary")
    exportados = 0 : semHandle = 0

    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            handle = Trim(CStr(lo.DataBodyRange.Cells(i, colHandle).Value))

            If handle = "" Then
                semHandle = semHandle + 1
            ElseIf Not handlesVistos.Exists(UCase(handle)) Then
                handlesVistos(UCase(handle)) = True

                If colNome > 0 Then
                    nomeBloco = Trim(CStr(lo.DataBodyRange.Cells(i, colNome).Value))
                Else
                    nomeBloco = ""
                End If

                linha = handle & ";" & nomeBloco
                For c = 0 To nCols - 1
                    val = Trim(CStr(lo.DataBodyRange.Cells(i, idxCols(c)).Value))
                    val = Replace(val, ";", ",") ' evita quebrar o CSV
                    linha = linha & ";" & val
                Next c
                Print #numArq, linha
                exportados = exportados + 1
            End If
        Next i
    End If

    Close #numArq
    Application.ScreenUpdating = True

    Dim msgExp As String
    msgExp = exportados & " bloco(s) exportado(s) para:" & vbNewLine & caminho & vbNewLine & vbNewLine & _
             "Proximo passo: no AutoCAD execute o comando ImportarTodos."
    If semHandle > 0 Then
        msgExp = msgExp & vbNewLine & "(" & semHandle & " linha(s) sem HANDLE ignoradas)"
    End If
    MsgBox msgExp, vbInformation, "Exportar para CAD"
    Exit Sub

ErroArquivo:
    Close #numArq
    Application.ScreenUpdating = True
    MsgBox "Erro ao gravar o arquivo: " & Err.Description, vbCritical, "Erro"
End Sub


' ================================================================
' MODULO 4 — AtualizarE_Sincronizar_Total2
' Sincroniza Tabela_Componentes e MES3 via Power Query,
' preservando valores manuais e detectando duplicatas.
' ================================================================

Sub AtualizarE_Sincronizar_Total2()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Componentes")
    Dim lo As ListObject: Set lo = ws.ListObjects("Tabela_Componentes")
    Dim dados As Object: Set dados = CreateObject("Scripting.Dictionary")
    Dim i As Long, j As Long, chave As String, tagAtual As String, indiceAtual As String
    Dim contDuplicados As Long: contDuplicados = 0

    Const COL_LIMITE As Long = 20

    RealizarBackupFisico

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1. Memoriza formulas e valores
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            tagAtual    = UCase(Trim(CStr(ws.Cells(lo.DataBodyRange.Row + i - 1, "H").Value)))
            indiceAtual = CStr(ws.Cells(lo.DataBodyRange.Row + i - 1, "B").Value)
            chave = tagAtual & "|" & indiceAtual
            If tagAtual <> "" Then
                Dim arr(2 To COL_LIMITE) As Variant
                For j = 2 To COL_LIMITE
                    If ws.Cells(lo.DataBodyRange.Row + i - 1, j).HasFormula Then
                        arr(j) = ws.Cells(lo.DataBodyRange.Row + i - 1, j).FormulaLocal
                    Else
                        arr(j) = ws.Cells(lo.DataBodyRange.Row + i - 1, j).Value
                    End If
                Next j
                dados(chave) = arr
            End If
        Next i
    End If

    ' 2. Atualiza Power Query
    lo.QueryTable.Refresh BackgroundQuery:=False
    DoEvents

    ' 3. Reintegracao e conferencia
    ws.Cells.Validation.Delete
    If Not lo.DataBodyRange Is Nothing Then
        Dim rng As Range: Set rng = lo.DataBodyRange
        rng.Font.ColorIndex = xlAutomatic

        For i = 1 To lo.ListRows.Count
            tagAtual    = UCase(Trim(CStr(ws.Cells(rng.Row + i - 1, "H").Value)))
            indiceAtual = CStr(ws.Cells(rng.Row + i - 1, "B").Value)
            chave = tagAtual & "|" & indiceAtual

            If tagAtual <> "" And WorksheetFunction.CountIf(lo.ListColumns("TAG").DataBodyRange, tagAtual) > 1 Then
                ws.Cells(rng.Row + i - 1, "H").Interior.Color = RGB(255, 0, 0)
                contDuplicados = contDuplicados + 1
            Else
                ws.Cells(rng.Row + i - 1, "H").Interior.ColorIndex = xlNone
            End If

            If dados.Exists(chave) Then
                Dim resgatados As Variant: resgatados = dados(chave)
                For j = 2 To COL_LIMITE
                    If j <> 8 And j <> 2 Then
                        Dim cel2 As Range: Set cel2 = ws.Cells(rng.Row + i - 1, j)
                        Dim vAntigo As Variant: vAntigo = resgatados(j)
                        If Left(CStr(vAntigo), 1) = "=" Then
                            cel2.FormulaLocal = vAntigo
                            cel2.Font.Color = RGB(218, 165, 32)
                        ElseIf (cel2.Value = "" Or cel2.Value = "//") And Trim(CStr(vAntigo)) <> "" Then
                            cel2.Value = vAntigo
                            cel2.Font.Color = RGB(218, 165, 32)
                        ElseIf Trim(CStr(cel2.Value)) <> Trim(CStr(vAntigo)) And Trim(CStr(vAntigo)) <> "" Then
                            cel2.Value = vAntigo
                            cel2.Font.Color = RGB(255, 0, 0)
                        End If
                    End If
                Next j
            Else
                rng.Rows(i).Font.Color = RGB(255, 140, 0)
            End If
        Next i
    End If

    AtualizarMES3_Sequencial

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If contDuplicados = 0 Then
        MsgBox "Sincronizacao OK!", vbInformation
    Else
        MsgBox "Sincronizacao concluida com " & contDuplicados & " duplicata(s).", vbExclamation
    End If
End Sub

Private Sub AtualizarMES3_Sequencial()
    Dim wsDestino As Worksheet: Set wsDestino = ThisWorkbook.Worksheets("MES3")
    Dim loDestino As ListObject: Set loDestino = wsDestino.ListObjects("Tabela3")
    Dim wsFonte As Worksheet: Set wsFonte = ThisWorkbook.Worksheets("DADOS_MES3")
    Dim loFonte As ListObject: Set loFonte = wsFonte.ListObjects("Tabela_Fonte_MES3")

    Dim dictQuery As Object: Set dictQuery = CreateObject("Scripting.Dictionary")
    Dim i As Long, ultimoID As Long, tagAtual As String

    loFonte.QueryTable.Refresh BackgroundQuery:=False

    If Not loFonte.DataBodyRange Is Nothing Then
        For i = 1 To loFonte.ListRows.Count
            tagAtual = UCase(Trim(CStr(loFonte.ListColumns("tag").DataBodyRange.Cells(i).Value)))
            If tagAtual <> "" Then
                dictQuery(tagAtual) = Array( _
                    loFonte.ListColumns("name").DataBodyRange.Cells(i).Value, _
                    loFonte.ListColumns("description").DataBodyRange.Cells(i).Value, _
                    loFonte.ListColumns("tag_size").DataBodyRange.Cells(i).Value, _
                    loFonte.ListColumns("tag_amount").DataBodyRange.Cells(i).Value)
            End If
        Next i
    End If

    If Not loDestino.DataBodyRange Is Nothing Then
        For i = loDestino.ListRows.Count To 1 Step -1
            tagAtual = UCase(Trim(loDestino.ListColumns("tag").DataBodyRange.Cells(i).Value))
            If Not dictQuery.Exists(tagAtual) Then loDestino.ListRows(i).Delete
        Next i
    End If

    ultimoID = 0
    If Not loDestino.DataBodyRange Is Nothing Then
        ultimoID = Application.WorksheetFunction.Max(loDestino.ListColumns("id").DataBodyRange)
    End If

    Dim tagChave As Variant
    For Each tagChave In dictQuery.Keys
        Dim existe As Boolean: existe = False
        If Not loDestino.DataBodyRange Is Nothing Then
            If Not IsError(Application.Match(tagChave, loDestino.ListColumns("tag").DataBodyRange, 0)) Then existe = True
        End If
        If Not existe Then
            ultimoID = ultimoID + 1
            With loDestino.ListRows.Add
                .Range(1, 1).Value = ultimoID
                .Range(1, 2).Value = dictQuery(tagChave)(0)
                .Range(1, 3).Value = dictQuery(tagChave)(1)
                .Range(1, 4).Value = tagChave
                .Range(1, 5).Value = dictQuery(tagChave)(2)
                .Range(1, 6).Value = dictQuery(tagChave)(3)
            End With
        End If
    Next tagChave
End Sub

Private Sub RealizarBackupFisico()
    Dim caminhoBase As String: caminhoBase = ThisWorkbook.Path
    If caminhoBase = "" Then Exit Sub
    Dim caminhoBackup As String: caminhoBackup = caminhoBase & "\Backup\"
    If Dir(caminhoBackup, vbDirectory) = "" Then MkDir caminhoBackup
    Dim nomeArquivo As String
    nomeArquivo = "Backup_" & Format(Now, "yyyy-mm-dd_hhmm") & ".xlsm"
    ThisWorkbook.SaveCopyAs caminhoBackup & nomeArquivo
End Sub
