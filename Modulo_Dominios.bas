Attribute VB_Name = "Modulo_Dominios"
Option Explicit

' ================================================================
' MODULO Dominios — consome os dominios de validacao do LISP (RN4)
'
' Importa Desktop\dominios.csv (gerado por STWExportarDominios) para a aba
' oculta "Dominios" e fornece as opcoes de dropdown aos eventos da planilha.
'
' Formato do CSV: CAMPO;DEPENDE_DE;VALOR_PAI;OPCOES
'   - VALOR_PAI e OPCOES tem itens separados por "|"
'   - VALOR_PAI = "*" significa "qualquer outro valor do campo pai"
'
' Substitui as listas hard-coded do antigo Worksheet_SelectionChange:
' a regra agora vive no schema do LISP. Ver FLUXO_DADOS.md (Dominios).
' ================================================================

Public Const ABA_DOMINIOS As String = "Dominios"

' Le o CSV de dominios para a aba oculta "Dominios" (colunas A:D).
Sub ImportarDominios()
    Dim caminho As String
    caminho = Environ("USERPROFILE") & "\Desktop\dominios.csv"
    If Dir(caminho) = "" Then
        MsgBox "Arquivo nao encontrado:" & vbNewLine & caminho & vbNewLine & vbNewLine & _
               "Execute o comando STWExportarDominios no AutoCAD.", vbExclamation, "Dominios"
        Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_DOMINIOS)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = ABA_DOMINIOS
    End If
    ws.Visible = xlSheetVisible
    ws.Cells.Clear

    Dim numArq As Integer, linha As String, r As Long
    numArq = FreeFile
    Open caminho For Input As #numArq
    r = 1
    Do While Not EOF(numArq)
        Line Input #numArq, linha
        linha = StripCR(linha)
        If linha <> "" Then
            Dim partes() As String
            partes = Split(linha, ";")
            Dim c As Long
            For c = 0 To UBound(partes)
                ws.Cells(r, c + 1).Value = "'" & partes(c) ' texto literal
            Next c
            r = r + 1
        End If
    Loop
    Close #numArq

    ws.Visible = xlSheetHidden
End Sub

' Retorna a lista de opcoes (separada por ",") para um CAMPO, dado o valor
' do campo pai. Procura a regra cujo VALOR_PAI contem valorPai; se nenhuma
' regra especifica casar, usa a regra coringa "*". Vazio se nao houver.
Public Function ObterOpcoesDominio(ByVal campo As String, ByVal valorPai As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_DOMINIOS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim ultima As Long: ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long, opcoesCoringa As String
    campo = UCase(Trim(campo))
    valorPai = UCase(Trim(valorPai))

    For i = 2 To ultima ' linha 1 = cabecalho
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = campo Then
            Dim valoresPai As String: valoresPai = Trim(CStr(ws.Cells(i, 3).Value))
            Dim opcoes As String: opcoes = Replace(CStr(ws.Cells(i, 4).Value), "|", ",")
            If valoresPai = "*" Then
                opcoesCoringa = opcoes
            Else
                Dim v As Variant
                For Each v In Split(valoresPai, "|")
                    If UCase(Trim(CStr(v))) = valorPai Then
                        ObterOpcoesDominio = opcoes
                        Exit Function
                    End If
                Next v
            End If
        End If
    Next i

    ObterOpcoesDominio = opcoesCoringa
End Function

' Remove \r final (CRLF do Windows/Excel)
Private Function StripCR(s As String) As String
    If Len(s) > 0 And Right(s, 1) = Chr(13) Then
        StripCR = Left(s, Len(s) - 1)
    Else
        StripCR = s
    End If
End Function
