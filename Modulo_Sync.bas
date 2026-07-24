Attribute VB_Name = "Modulo_Sync"
Option Explicit

' ================================================================
' MODULO_SYNC — Sincronizacao Componentes + MES3 (v2, chave por ocorrencia)
'
' COMO USAR: cole este codigo POR CIMA do conteudo do Modulo1 existente
' (substituindo tudo). Os nomes das macros foram mantidos
' (AtualizarE_Sincronizar_Total2 etc.) para os botoes continuarem
' funcionando. NAO deixe duas copias das mesmas macros em modulos
' diferentes — gera "Nome ambiguo detectado" e trava o projeto.
'
' ATUALIZACOES INCORPORADAS (versao do usuario + correcoes):
'   1. CHAVE POR OCORRENCIA: memoriza/reintegra por TAG + ordem de aparicao
'      ("VALV-01|1", "VALV-01|2") — amarra corretamente TAGs repetidas.
'   2. COLUNAS DINAMICAS: TAG/HANDLE/INDICE resolvidos por ListColumns
'      (no layout v2 do Power Query a TAG e a coluna F, nao mais a H).
'      Colunas protegidas (nunca sobrescritas): INDICE, HANDLE, TAG.
'   3. BLINDAGEM: handler de erro que SEMPRE restaura EnableEvents/
'      Calculation/ScreenUpdating (sem isso, um erro no refresh deixa o
'      Excel "travado").
'   4. PERFORMANCE: memorizacao/reintegracao em ARRAYS (uma leitura, uma
'      escrita), dedup por Dictionary (sem CountIf O(N^2)), MES3 em bloco.
'
' Cores na reintegracao (regras preservadas):
'   OURO     = formula/valor manual restaurado apos o refresh
'   VERMELHO = divergencia (CAD difere do manual) / TAG duplicada (fundo)
'   LARANJA  = linha nova vinda do CAD
' ================================================================

' Gera copia .xlsm em \Backup\ antes de cada sync (False = so via BackupManual)
Private Const FAZER_BACKUP As Boolean = True

Private Const COR_OURO As Long = 2139610       ' RGB(218,165,32)
Private Const COR_VERMELHO As Long = 255        ' RGB(255,0,0)
Private Const COR_LARANJA As Long = 36095       ' RGB(255,140,0)

Public Sub AtualizarE_Sincronizar_Total2()
    Dim ws As Worksheet, lo As ListObject
    Dim mem As Object
    Dim contDuplicados As Long
    Dim prevEvents As Boolean, prevScreen As Boolean
    Dim prevCalc As XlCalculation

    On Error GoTo Limpeza

    Set ws = ThisWorkbook.Worksheets("Componentes")
    Set lo = ws.ListObjects("Tabela_Componentes")

    ' --- BACKUP (opcional via constante) ---
    If FAZER_BACKUP Then RealizarBackupFisico

    ' --- Configuracoes de performance (restauradas SEMPRE em Limpeza) ---
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- PARTE 1: COMPONENTES ---
    ' 1. Memorizacao por TAG + ocorrencia (formulas e valores)
    Set mem = MemorizarComponentes(lo)

    ' 2. Atualizacao do Power Query
    lo.QueryTable.Refresh BackgroundQuery:=False
    DoEvents

    ' 3. Reintegracao + conferencia de duplicidade
    ws.Cells.Validation.Delete
    contDuplicados = ReintegrarComponentes(lo, mem)

    ' --- PARTE 2: MES3 ---
    AtualizarMES3_Sequencial

    ' --- Finalizacao ---
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen

    If contDuplicados = 0 Then
        MsgBox "Sincronizacao OK!", vbInformation
    Else
        MsgBox "Sincronizacao concluida com " & contDuplicados & " duplicatas.", vbExclamation
    End If
    Exit Sub

Limpeza:
    ' Restaura SEMPRE o estado do Excel para nao travar a planilha
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    On Error GoTo 0
    MsgBox "Falha na sincronizacao: " & Err.Description, vbExclamation, "Sincronizar"
End Sub

' ----------------------------------------------------------------
' Memoriza cada linha por chave TAG|ocorrencia ("VALV-01|1", "VALV-01|2").
' Guarda a linha inteira em FormulaLocal (preserva formulas E valores).
' ----------------------------------------------------------------
Private Function MemorizarComponentes(lo As ListObject) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set MemorizarComponentes = d
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim colTag As Long: colTag = lo.ListColumns("TAG").Index

    Dim contador As Object: Set contador = CreateObject("Scripting.Dictionary")
    contador.CompareMode = vbTextCompare

    Dim f As Variant: f = lo.DataBodyRange.FormulaLocal   ' 2D linhas x colunas
    Dim nR As Long, nC As Long, r As Long, c As Long
    nR = UBound(f, 1): nC = UBound(f, 2)

    For r = 1 To nR
        Dim tagAtual As String
        tagAtual = UCase(Trim(CStr(f(r, colTag))))
        If tagAtual <> "" Then
            ' Incrementa a ocorrencia desta TAG antes do refresh
            If contador.Exists(tagAtual) Then
                contador(tagAtual) = contador(tagAtual) + 1
            Else
                contador(tagAtual) = 1
            End If
            Dim chave As String: chave = tagAtual & "|" & contador(tagAtual)

            Dim linha() As Variant: ReDim linha(1 To nC)
            For c = 1 To nC
                linha(c) = f(r, c)
            Next c
            d(chave) = linha
        End If
    Next r
End Function

' ----------------------------------------------------------------
' Reintegra o trabalho manual pela MESMA chave TAG|ocorrencia (recontada
' na tabela pos-refresh) e marca TAG duplicada. Retorna qtde de duplicatas.
' Colunas protegidas (nunca sobrescritas): INDICE, HANDLE, TAG.
' ----------------------------------------------------------------
Private Function ReintegrarComponentes(lo As ListObject, mem As Object) As Long
    ReintegrarComponentes = 0
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim colTag As Long: colTag = lo.ListColumns("TAG").Index
    Dim colHnd As Long: colHnd = ColunaOpcional(lo, "HANDLE")
    Dim colIdx As Long: colIdx = ColunaOpcional(lo, "ÍNDICE")

    Dim novoV As Variant: novoV = lo.DataBodyRange.Value        ' valores pos-refresh
    Dim saida As Variant: saida = lo.DataBodyRange.FormulaLocal ' base para escrita
    Dim nR As Long, nC As Long, r As Long, c As Long
    nR = UBound(novoV, 1): nC = UBound(novoV, 2)

    ' Reset de formatacao em lote
    lo.DataBodyRange.Font.ColorIndex = xlAutomatic
    lo.ListColumns("TAG").DataBodyRange.Interior.ColorIndex = xlNone

    ' Contagem total de cada TAG (para dedup) numa unica passada
    Dim contaTag As Object: Set contaTag = CreateObject("Scripting.Dictionary")
    contaTag.CompareMode = vbTextCompare
    For r = 1 To nR
        Dim t As String: t = UCase(Trim(CStr(novoV(r, colTag))))
        If t <> "" Then
            If contaTag.Exists(t) Then contaTag(t) = contaTag(t) + 1 Else contaTag(t) = 1
        End If
    Next r

    ' Reconta ocorrencias na tabela atualizada e reintegra
    Dim contador As Object: Set contador = CreateObject("Scripting.Dictionary")
    contador.CompareMode = vbTextCompare

    Dim rgOuro As Range, rgVerm As Range, rgLaranja As Range, rgDup As Range
    Dim contDup As Long: contDup = 0

    For r = 1 To nR
        Dim tagAtual As String, chave As String
        tagAtual = UCase(Trim(CStr(novoV(r, colTag))))

        If tagAtual <> "" Then
            If contador.Exists(tagAtual) Then
                contador(tagAtual) = contador(tagAtual) + 1
            Else
                contador(tagAtual) = 1
            End If
            chave = tagAtual & "|" & contador(tagAtual)

            ' Duplicidade: TAG aparece mais de uma vez -> fundo vermelho
            If contaTag(tagAtual) > 1 Then
                AddCelula rgDup, lo, r, colTag
                contDup = contDup + 1
            End If
        Else
            chave = ""
        End If

        If chave <> "" And mem.Exists(chave) Then
            Dim antigo As Variant: antigo = mem(chave)
            For c = 1 To nC
                ' Protege INDICE, HANDLE e TAG (chaves/indices novos do banco)
                If c <> colTag And c <> colHnd And c <> colIdx Then
                    Dim vAnt As Variant: vAnt = antigo(c)
                    Dim vNew As Variant: vNew = novoV(r, c)

                    ' Formula memorizada -> restaura (ouro)
                    If Left(CStr(vAnt), 1) = "=" Then
                        saida(r, c) = vAnt
                        AddCelula rgOuro, lo, r, c

                    ' Banco trouxe vazio/"//" e havia valor -> restaura (ouro)
                    ElseIf (CStr(vNew) = "" Or CStr(vNew) = "//") And Trim(CStr(vAnt)) <> "" Then
                        saida(r, c) = vAnt
                        AddCelula rgOuro, lo, r, c

                    ' Divergencia manual -> restaura memorizado (vermelho)
                    ElseIf Trim(CStr(vNew)) <> Trim(CStr(vAnt)) And Trim(CStr(vAnt)) <> "" Then
                        saida(r, c) = vAnt
                        AddCelula rgVerm, lo, r, c
                    End If
                End If
            Next c
        Else
            ' TAG/ocorrencia inedita -> linha nova vinda do CAD (laranja)
            If rgLaranja Is Nothing Then
                Set rgLaranja = lo.DataBodyRange.Rows(r)
            Else
                Set rgLaranja = Union(rgLaranja, lo.DataBodyRange.Rows(r))
            End If
        End If
    Next r

    ' Escreve a tabela de volta numa unica operacao
    lo.DataBodyRange.FormulaLocal = saida

    ' Cores em lote
    If Not rgOuro Is Nothing Then rgOuro.Font.Color = COR_OURO
    If Not rgVerm Is Nothing Then rgVerm.Font.Color = COR_VERMELHO
    If Not rgLaranja Is Nothing Then rgLaranja.Font.Color = COR_LARANJA
    If Not rgDup Is Nothing Then rgDup.Interior.Color = COR_VERMELHO

    ReintegrarComponentes = contDup
End Function

' Indice de uma coluna da tabela; 0 se nao existir (coluna opcional)
Private Function ColunaOpcional(lo As ListObject, nome As String) As Long
    ColunaOpcional = 0
    On Error Resume Next
    ColunaOpcional = lo.ListColumns(nome).Index
    On Error GoTo 0
End Function

' Acumula a celula (r,c) do corpo da tabela numa Range (para formatar em lote)
Private Sub AddCelula(ByRef acc As Range, lo As ListObject, r As Long, c As Long)
    Dim cel As Range: Set cel = lo.DataBodyRange.Cells(r, c)
    If acc Is Nothing Then
        Set acc = cel
    Else
        Set acc = Union(acc, cel)
    End If
End Sub

' ----------------------------------------------------------------
' MES3 — regras originais preservadas (remove o que sumiu da fonte,
' adiciona novos com id = max+1, ids existentes estaveis), mas montada
' em BLOCO unico (sem ListRows.Add/Delete em laco — muito mais rapido).
' ----------------------------------------------------------------
Private Sub AtualizarMES3_Sequencial()
    Dim loFonte As ListObject, loDest As ListObject
    Set loFonte = ThisWorkbook.Worksheets("DADOS_MES3").ListObjects("Tabela_Fonte_MES3")
    Set loDest = ThisWorkbook.Worksheets("MES3").ListObjects("Tabela3")

    ' 1. Atualiza a fonte da MES3
    loFonte.QueryTable.Refresh BackgroundQuery:=False

    ' 2. ids existentes: tag -> id (estaveis) e maior id atual
    Dim idsExist As Object: Set idsExist = CreateObject("Scripting.Dictionary")
    idsExist.CompareMode = vbTextCompare
    Dim maxId As Long: maxId = 0
    If Not loDest.DataBodyRange Is Nothing Then
        Dim dv As Variant: dv = loDest.DataBodyRange.Value
        Dim cTagD As Long, cIdD As Long, i As Long
        cTagD = loDest.ListColumns("tag").Index
        cIdD = loDest.ListColumns("id").Index
        For i = 1 To UBound(dv, 1)
            Dim tD As String: tD = UCase(Trim(CStr(dv(i, cTagD))))
            Dim idv As Long: idv = Val(CStr(dv(i, cIdD)))
            If tD <> "" And Not idsExist.Exists(tD) Then idsExist(tD) = idv
            If idv > maxId Then maxId = idv
        Next i
    End If

    ' 3. Monta a saida a partir da fonte (id, name, description, tag, tag_size, tag_amount)
    If loFonte.DataBodyRange Is Nothing Then Exit Sub
    Dim fv As Variant: fv = loFonte.DataBodyRange.Value
    Dim cTag As Long, cName As Long, cDesc As Long, cSize As Long, cAmt As Long
    cTag = loFonte.ListColumns("tag").Index
    cName = loFonte.ListColumns("name").Index
    cDesc = loFonte.ListColumns("description").Index
    cSize = loFonte.ListColumns("tag_size").Index
    cAmt = loFonte.ListColumns("tag_amount").Index

    Dim nF As Long: nF = UBound(fv, 1)
    Dim saida() As Variant: ReDim saida(1 To nF, 1 To 6)
    Dim vistos As Object: Set vistos = CreateObject("Scripting.Dictionary")
    vistos.CompareMode = vbTextCompare
    Dim outN As Long: outN = 0

    For i = 1 To nF
        Dim tagF As String: tagF = UCase(Trim(CStr(fv(i, cTag))))
        If tagF <> "" And Not vistos.Exists(tagF) Then
            vistos(tagF) = True
            Dim idFinal As Long
            If idsExist.Exists(tagF) Then
                idFinal = idsExist(tagF)      ' preserva o id existente
            Else
                maxId = maxId + 1             ' novo item ao final
                idFinal = maxId
            End If
            outN = outN + 1
            saida(outN, 1) = idFinal          ' id
            saida(outN, 2) = fv(i, cName)     ' name
            saida(outN, 3) = fv(i, cDesc)     ' description
            saida(outN, 4) = fv(i, cTag)      ' tag
            saida(outN, 5) = fv(i, cSize)     ' tag_size
            saida(outN, 6) = fv(i, cAmt)      ' tag_amount
        End If
    Next i

    ' 4. Reescreve a Tabela3 em bloco (o que sumiu da fonte sai automaticamente)
    If Not loDest.DataBodyRange Is Nothing Then loDest.DataBodyRange.Delete
    If outN = 0 Then Exit Sub
    Dim destino As Range
    Set destino = loDest.HeaderRowRange.Cells(1, 1).Offset(1, 0).Resize(outN, 6)
    loDest.Resize loDest.Range.Resize(outN + 1, loDest.Range.Columns.Count)
    destino.Value = saida
End Sub

' ----------------------------------------------------------------
' Backup fisico timestamped em \Backup\ ao lado do arquivo.
' Chamado no inicio do sync (se FAZER_BACKUP=True) ou manualmente.
' ----------------------------------------------------------------
Private Sub RealizarBackupFisico()
    Dim caminhoBase As String: caminhoBase = ThisWorkbook.Path
    If caminhoBase = "" Then Exit Sub

    Dim caminhoBackup As String: caminhoBackup = caminhoBase & "\Backup\"
    If Dir(caminhoBackup, vbDirectory) = "" Then MkDir caminhoBackup

    Dim nomeArquivo As String
    nomeArquivo = "Backup_" & Format(Now, "yyyy-mm-dd_hhmm") & ".xlsm"
    ThisWorkbook.SaveCopyAs caminhoBackup & nomeArquivo
End Sub

' Backup sob demanda (Alt+F8 -> BackupManual)
Public Sub BackupManual()
    RealizarBackupFisico
    MsgBox "Backup criado na pasta \Backup\.", vbInformation
End Sub
