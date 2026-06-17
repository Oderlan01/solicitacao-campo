' ================================================================
' MODULO 1 — SalvarNaBaseServidor
' Exporta a tabela "teste" da aba ativa para um arquivo .txt
' ================================================================
' Attribute VB_Name = "Modulo_SalvarBase"
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
' Atualiza a linha correspondente no arquivo Base (teste.xlsx)
' com base na celula selecionada na tabela da aba ativa.
'
' CORRECOES APLICADAS:
'   - On Error GoTo ErroDados isolado apenas para leitura de colunas
'   - On Error GoTo 0 redefine o handler antes das operacoes de arquivo
'   - ErroGrava: agora e efetivamente ativado
' ================================================================
' Attribute VB_Name = "Modulo_GravarBase"

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

    ' --- Captura dados da linha selecionada (handler isolado) ---
    On Error GoTo ErroDados
    indice    = Intersect(ActiveCell.EntireRow, tabela.ListColumns("INDICE").Range).Value
    categoria = Intersect(ActiveCell.EntireRow, tabela.ListColumns("CATEGORIA").Range).Value
    novoValor = Intersect(ActiveCell.EntireRow, tabela.ListColumns("TAG").Range).Value
    On Error GoTo 0  ' reseta antes das operacoes de arquivo

    ' --- Abre o arquivo Base ---
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Abrindo a base e aplicando alteracoes..."

    On Error GoTo ErroGrava
    Set wbBase = Workbooks.Open(caminhoBase)
    Set wsBase = wbBase.Sheets("Summary")
    linhaBase = indice + 1

    ' --- Mapeia categoria para coluna na Base ---
    Select Case UCase(Trim(categoria))
        Case "VALVULA ABRE":          wsBase.Cells(linhaBase, MapearColuna(wsBase, "VALVULA_ABRE")).Value = novoValor
        Case "VALVULA FECHA":         wsBase.Cells(linhaBase, MapearColuna(wsBase, "VALVULA_FECHA")).Value = novoValor
        Case "SENSOR ABERTO":         wsBase.Cells(linhaBase, MapearColuna(wsBase, "SENSOR_ABERTO")).Value = novoValor
        Case "SENSOR FECHADO":        wsBase.Cells(linhaBase, MapearColuna(wsBase, "SENSOR_FECHADO")).Value = novoValor
        Case "SENSOR NIVEL BAIXO":    wsBase.Cells(linhaBase, MapearColuna(wsBase, "SENSOR_NIVEL_BAIXO")).Value = novoValor
        Case "SENSOR NIVEL ALTO":     wsBase.Cells(linhaBase, MapearColuna(wsBase, "SENSOR_NIVEL_ALTO")).Value = novoValor
        Case Else
            wsBase.Cells(linhaBase, MapearColuna(wsBase, "0E_TAG")).Value = novoValor
    End Select

    ' --- Dados de engenharia (opcional) ---
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

    MsgBox "Base atualizada com sucesso!" & vbCrLf & _
           "INDICE " & indice & " -> linha " & linhaBase & " da base.", vbInformation, "Sucesso"
    Exit Sub

ErroDados:
    MsgBox "Erro ao ler colunas da tabela. Verifique: INDICE, CATEGORIA, TAG.", vbCritical, "Erro de Coluna"
    Exit Sub

ErroGrava:
    If Not wbBase Is Nothing Then wbBase.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = ""
    MsgBox "Erro ao salvar na base: " & Err.Description, vbCritical, "Erro"
End Sub

Function MapearColuna(aba As Worksheet, nomeColuna As String) As Long
    Dim celula As Range
    Set celula = aba.Rows(1).Find(What:=nomeColuna, LookIn:=xlValues, LookAt:=xlWhole)
    If Not celula Is Nothing Then
        MapearColuna = celula.Column
    Else
        MsgBox "Coluna '" & nomeColuna & "' nao encontrada na base!", vbCritical, "Erro de Estrutura"
        End
    End If
End Function


' ================================================================
' MODULO 3 — ExportarParaCAD
' Exporta os campos editaveis da Tabela_Componentes para
' Desktop\campos_manuais.csv, pronto para ser importado no
' AutoCAD com o comando ImportarCamposManuais (automacao.lsp).
' ================================================================
' Attribute VB_Name = "Modulo_ExportarParaCAD"

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

    ' Mapeamento: nome da coluna Excel -> tag do atributo no AutoCAD
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

    ' Descobrir quais colunas do mapa existem na tabela
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
        MsgBox "Nenhuma coluna mapeavel encontrada na tabela.", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' Coluna HANDLE (obrigatoria)
    Dim colHandle As Long
    colHandle = 0
    On Error Resume Next
    colHandle = lo.ListColumns("HANDLE").Index
    On Error GoTo 0
    If colHandle = 0 Then
        MsgBox "Coluna HANDLE nao encontrada em Tabela_Componentes.", vbCritical, "Erro"
        Exit Sub
    End If

    ' Gerar CSV
    caminho = Environ("USERPROFILE") & "\Desktop\campos_manuais.csv"
    Application.ScreenUpdating = False

    On Error GoTo ErroArquivo
    numArq = FreeFile
    Open caminho For Output As #numArq

    header = "HANDLE"
    For c = 0 To nCols - 1
        header = header & ";" & tagsCols(c)
    Next c
    Print #numArq, header

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
              "Atencao: " & semHandle & " linha(s) sem HANDLE ignoradas."
    End If
    MsgBox msg, vbInformation, "Exportar para CAD"
    Exit Sub

ErroArquivo:
    Close #numArq
    Application.ScreenUpdating = True
    MsgBox "Erro ao gravar o arquivo:" & vbNewLine & Err.Description, vbCritical, "Erro"
End Sub


' ================================================================
' MODULO 4 — AtualizarE_Sincronizar_Total2
' Sincroniza Tabela_Componentes e MES3 via Power Query,
' preservando valores manuais e detectando duplicatas.
' ================================================================
' Attribute VB_Name = "Modulo_Sincronizar"

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

    ' --- PARTE 1: COMPONENTES ---

    ' 1. Memoriza formulas e valores
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            tagAtual     = UCase(Trim(CStr(ws.Cells(lo.DataBodyRange.Row + i - 1, "H").Value)))
            indiceAtual  = CStr(ws.Cells(lo.DataBodyRange.Row + i - 1, "B").Value)
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
                        Dim cell As Range: Set cell = ws.Cells(rng.Row + i - 1, j)
                        Dim vAntigo As Variant: vAntigo = resgatados(j)

                        If Left(CStr(vAntigo), 1) = "=" Then
                            cell.FormulaLocal = vAntigo
                            cell.Font.Color = RGB(218, 165, 32)
                        ElseIf (cell.Value = "" Or cell.Value = "//") And Trim(CStr(vAntigo)) <> "" Then
                            cell.Value = vAntigo
                            cell.Font.Color = RGB(218, 165, 32)
                        ElseIf Trim(CStr(cell.Value)) <> Trim(CStr(vAntigo)) And Trim(CStr(vAntigo)) <> "" Then
                            cell.Value = vAntigo
                            cell.Font.Color = RGB(255, 0, 0)
                        End If
                    End If
                Next j
            Else
                rng.Rows(i).Font.Color = RGB(255, 140, 0)
            End If
        Next i
    End If

    ' --- PARTE 2: MES3 ---
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
