Attribute VB_Name = "Modulo_Sync"
Option Explicit

' ================================================================
' MODULO_SYNC — Sincronizacao Componentes + MES3  (REFORMULADO / enxuto)
'
' IMPORTANTE ao importar no workbook: REMOVA antes as macros antigas do
' Modulo1 (AtualizarE_Sincronizar_Total2, AtualizarMES3_Sequencial,
' RealizarBackupFisico). Manter as duas versoes causa "Nome ambiguo
' detectado", que trava o projeto inteiro. Este modulo usa nomes novos
' (AtualizarTudo / BackupManual) para nao colidir.
'
' Objetivo da reformulacao: acabar com as travas. Mudancas-chave:
'   1. AtualizarTudo SEMPRE restaura EnableEvents/Calculation/ScreenUpdating
'      (handler "Limpeza"), mesmo em erro. Antes, um erro no meio deixava o
'      Excel "congelado" (eventos off, calculo manual).
'   2. Memorizacao/reintegracao das colunas manuais em ARRAYS (uma leitura e
'      uma escrita), nao mais celula a celula.
'   3. Deduplicacao por Dictionary numa unica passada (sem CountIf O(N^2)).
'   4. MES3 montada em BLOCO unico (sem ListRows.Add/Delete em laco).
'   5. Backup deixou de rodar a cada sync -> agora BackupManual sob demanda.
'
' Chave de negocio Componentes = TAG + HANDLE (colunas protegidas: nao sao
' sobrescritas na reintegracao). Ver FLUXO_DADOS.md (RN1-RN3).
'
' Cores na reintegracao:
'   OURO     = formula/valor manual restaurado apos o refresh
'   VERMELHO = divergencia (valor do CAD difere do manual) / TAG duplicada
'   LARANJA  = linha nova vinda do CAD
' ================================================================

Private Const COR_OURO As Long = 1671391      ' RGB(218,165,32)
Private Const COR_VERMELHO As Long = 255       ' RGB(255,0,0)
Private Const COR_LARANJA As Long = 36095      ' RGB(255,140,0)

Public Sub AtualizarTudo()
    Dim ws As Worksheet, lo As ListObject
    Dim mem As Object
    Dim prevEvents As Boolean, prevScreen As Boolean
    Dim prevCalc As XlCalculation

    On Error GoTo Limpeza

    Set ws = ThisWorkbook.Worksheets("Componentes")
    Set lo = ws.ListObjects("Tabela_Componentes")

    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1. Memoriza o estado manual (formulas e valores) antes do refresh
    Set mem = MemorizarComponentes(lo)

    ' 2. Atualiza o Power Query (Componentes)
    lo.QueryTable.Refresh BackgroundQuery:=False

    ' 3. Reintegra o trabalho manual + marca duplicatas
    ReintegrarComponentes lo, mem

    ' 4. MES3 (bloco unico)
    AtualizarMES3

    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox "Sincronizacao concluida.", vbInformation, "Atualizar"
    Exit Sub

Limpeza:
    ' Restaura SEMPRE o estado do Excel para nao travar
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    On Error GoTo 0
    MsgBox "Falha na sincronizacao: " & Err.Description, vbExclamation, "Atualizar"
End Sub

' ----------------------------------------------------------------
' Memoriza, por chave (TAG|HANDLE), a linha inteira em formato .Formula
' (preserva formulas e valores). Retorna Dictionary chave -> array 1D.
' ----------------------------------------------------------------
Private Function MemorizarComponentes(lo As ListObject) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set MemorizarComponentes = d
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim colTag As Long, colHnd As Long
    colTag = lo.ListColumns("TAG").Index
    colHnd = lo.ListColumns("HANDLE").Index

    Dim f As Variant: f = lo.DataBodyRange.Formula  ' 2D (linhas x colunas)
    Dim nR As Long, nC As Long, r As Long, c As Long
    nR = UBound(f, 1): nC = UBound(f, 2)

    For r = 1 To nR
        Dim chave As String
        chave = UCase(Trim(CStr(f(r, colTag)))) & "|" & UCase(Trim(CStr(f(r, colHnd))))
        If Trim(CStr(f(r, colTag))) <> "" Then
            Dim linha() As Variant: ReDim linha(1 To nC)
            For c = 1 To nC
                linha(c) = f(r, c)
            Next c
            d(chave) = linha
        End If
    Next r
End Function

' ----------------------------------------------------------------
' Reintegra o estado manual memorizado e marca duplicatas, tudo em array.
' ----------------------------------------------------------------
Private Sub ReintegrarComponentes(lo As ListObject, mem As Object)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim colTag As Long, colHnd As Long
    colTag = lo.ListColumns("TAG").Index
    colHnd = lo.ListColumns("HANDLE").Index

    Dim novoV As Variant: novoV = lo.DataBodyRange.Value     ' valores pos-refresh
    Dim saida As Variant: saida = lo.DataBodyRange.Formula   ' base para escrita
    Dim nR As Long, nC As Long, r As Long, c As Long
    nR = UBound(novoV, 1): nC = UBound(novoV, 2)

    ' Reset de formatacao (uma operacao)
    lo.DataBodyRange.Font.ColorIndex = xlAutomatic
    lo.ListColumns("TAG").DataBodyRange.Interior.ColorIndex = xlNone

    Dim contaTag As Object: Set contaTag = CreateObject("Scripting.Dictionary")
    contaTag.CompareMode = vbTextCompare

    Dim rgOuro As Range, rgVerm As Range, rgLaranja As Range, rgDup As Range

    For r = 1 To nR
        Dim tagAtual As String, hnd As String, chave As String
        tagAtual = UCase(Trim(CStr(novoV(r, colTag))))
        hnd = UCase(Trim(CStr(novoV(r, colHnd))))
        chave = tagAtual & "|" & hnd

        ' Contagem de TAG para deduplicacao (uma passada)
        If tagAtual <> "" Then
            If contaTag.Exists(tagAtual) Then
                contaTag(tagAtual) = contaTag(tagAtual) + 1
            Else
                contaTag(tagAtual) = 1
            End If
        End If

        If mem.Exists(chave) Then
            Dim antigo As Variant: antigo = mem(chave)
            For c = 1 To nC
                If c <> colTag And c <> colHnd Then
                    Dim vAnt As Variant: vAnt = antigo(c)
                    Dim vNew As Variant: vNew = novoV(r, c)
                    If Left(CStr(vAnt), 1) = "=" Then
                        saida(r, c) = vAnt
                        AddCelula rgOuro, lo, r, c
                    ElseIf (CStr(vNew) = "" Or CStr(vNew) = "//") And Trim(CStr(vAnt)) <> "" Then
                        saida(r, c) = vAnt
                        AddCelula rgOuro, lo, r, c
                    ElseIf Trim(CStr(vNew)) <> Trim(CStr(vAnt)) And Trim(CStr(vAnt)) <> "" Then
                        saida(r, c) = vAnt
                        AddCelula rgVerm, lo, r, c
                    End If
                End If
            Next c
        Else
            ' Linha nova vinda do CAD
            If rgLaranja Is Nothing Then
                Set rgLaranja = lo.DataBodyRange.Rows(r)
            Else
                Set rgLaranja = Union(rgLaranja, lo.DataBodyRange.Rows(r))
            End If
        End If
    Next r

    ' Escreve a tabela de volta numa unica operacao
    lo.DataBodyRange.Formula = saida

    ' Marca duplicatas de TAG (fundo vermelho na celula TAG)
    For r = 1 To nR
        tagAtual = UCase(Trim(CStr(novoV(r, colTag))))
        If tagAtual <> "" Then
            If contaTag(tagAtual) > 1 Then AddCelula rgDup, lo, r, colTag
        End If
    Next r

    ' Aplica cores em lote
    If Not rgOuro Is Nothing Then rgOuro.Font.Color = COR_OURO
    If Not rgVerm Is Nothing Then rgVerm.Font.Color = COR_VERMELHO
    If Not rgLaranja Is Nothing Then rgLaranja.Font.Color = COR_LARANJA
    If Not rgDup Is Nothing Then rgDup.Interior.Color = COR_VERMELHO
End Sub

' Adiciona a celula (r,c) relativa ao corpo da tabela a uma Range acumuladora.
Private Sub AddCelula(ByRef acc As Range, lo As ListObject, r As Long, c As Long)
    Dim cel As Range: Set cel = lo.DataBodyRange.Cells(r, c)
    If acc Is Nothing Then
        Set acc = cel
    Else
        Set acc = Union(acc, cel)
    End If
End Sub

' ----------------------------------------------------------------
' MES3 — monta a tabela destino em BLOCO unico (sem Add/Delete em laco).
' Preserva o id de tags que ja existem; novos recebem max(id)+1.
' ----------------------------------------------------------------
Private Sub AtualizarMES3()
    Dim loFonte As ListObject, loDest As ListObject
    Set loFonte = ThisWorkbook.Worksheets("DADOS_MES3").ListObjects("Tabela_Fonte_MES3")
    Set loDest = ThisWorkbook.Worksheets("MES3").ListObjects("Tabela3")

    ' 1. Atualiza a fonte
    loFonte.QueryTable.Refresh BackgroundQuery:=False

    ' 2. ids existentes: tag -> id (preserva) e descobre o maior id
    Dim idsExist As Object: Set idsExist = CreateObject("Scripting.Dictionary")
    idsExist.CompareMode = vbTextCompare
    Dim maxId As Long: maxId = 0
    If Not loDest.DataBodyRange Is Nothing Then
        Dim dv As Variant: dv = loDest.DataBodyRange.Value
        Dim cTagD As Long, cIdD As Long, i As Long
        cTagD = loDest.ListColumns("tag").Index
        cIdD = loDest.ListColumns("id").Index
        For i = 1 To UBound(dv, 1)
            Dim t As String: t = UCase(Trim(CStr(dv(i, cTagD))))
            Dim idv As Long: idv = Val(CStr(dv(i, cIdD)))
            If t <> "" And Not idsExist.Exists(t) Then idsExist(t) = idv
            If idv > maxId Then maxId = idv
        Next i
    End If

    ' 3. Le a fonte e monta a saida (id, name, description, tag, tag_size, tag_amount)
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
    Dim out As Long: out = 0

    For i = 1 To nF
        Dim tag As String: tag = UCase(Trim(CStr(fv(i, cTag))))
        If tag <> "" And Not vistos.Exists(tag) Then
            vistos(tag) = True
            Dim idFinal As Long
            If idsExist.Exists(tag) Then
                idFinal = idsExist(tag)
            Else
                maxId = maxId + 1
                idFinal = maxId
            End If
            out = out + 1
            saida(out, 1) = idFinal           ' id
            saida(out, 2) = fv(i, cName)       ' name
            saida(out, 3) = fv(i, cDesc)       ' description
            saida(out, 4) = fv(i, cTag)        ' tag
            saida(out, 5) = fv(i, cSize)       ' tag_size
            saida(out, 6) = fv(i, cAmt)        ' tag_amount
        End If
    Next i

    ' 4. Redimensiona a tabela destino e escreve em bloco
    EscreverTabela loDest, saida, out, 6
End Sub

' Redimensiona um ListObject para nLin linhas e escreve o array saida.
Private Sub EscreverTabela(lo As ListObject, saida As Variant, nLin As Long, nCol As Long)
    Dim hdr As Range: Set hdr = lo.HeaderRowRange
    ' Remove o corpo atual
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    If nLin = 0 Then Exit Sub
    ' Novo corpo logo abaixo do cabecalho
    Dim destino As Range
    Set destino = hdr.Cells(1, 1).Offset(1, 0).Resize(nLin, nCol)
    lo.Resize lo.Range.Resize(nLin + 1, lo.Range.Columns.Count)
    destino.Value = saida
End Sub

' ----------------------------------------------------------------
' Backup sob demanda (nao roda mais a cada sync).
' ----------------------------------------------------------------
Public Sub BackupManual()
    Dim base As String: base = ThisWorkbook.Path
    If base = "" Then
        MsgBox "Salve a pasta de trabalho antes de gerar backup.", vbExclamation
        Exit Sub
    End If
    Dim pasta As String: pasta = base & "\Backup\"
    If Dir(pasta, vbDirectory) = "" Then MkDir pasta
    Dim nome As String: nome = "Backup_" & Format(Now, "yyyy-mm-dd_hhmm") & ".xlsm"
    ThisWorkbook.SaveCopyAs pasta & nome
    MsgBox "Backup criado: " & pasta & nome, vbInformation
End Sub
