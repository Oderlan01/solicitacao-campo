VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha_Componentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ================================================================
' Modulo de codigo da planilha "Componentes"
'
' Worksheet_SelectionChange — dropdowns dependentes (RN4) montados a partir
' dos dominios exportados pelo LISP (aba oculta "Dominios", carregada por
' Modulo_Dominios.ImportarDominios). Substitui as listas hard-coded.
'
' Layout (ver power-query-csv.m, Selecao Final):
'   col 7  (G) = CATEGORIA
'   col 10 (J) = ACIONAMENTO  -> dropdown depende de CATEGORIA
'   col 11 (K) = INTERFACE    -> dropdown depende de ACIONAMENTO
' ================================================================

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Apenas uma celula por vez, linha 2+ e colunas 10 (J) ou 11 (K)
    If Target.CountLarge > 1 Or Target.Row < 2 Then Exit Sub
    If Target.Column <> 10 And Target.Column <> 11 Then Exit Sub

    Dim campo As String, valorPai As String, lista As String

    ' Limpa validacao anterior para evitar conflito de lista antiga
    On Error Resume Next
    Target.Validation.Delete
    On Error GoTo 0

    If Target.Column = 10 Then
        campo = "ACIONAMENTO"
        valorPai = CStr(Me.Cells(Target.Row, 7).Value)  ' CATEGORIA
    ElseIf Target.Column = 11 Then
        campo = "INTERFACE"
        valorPai = CStr(Me.Cells(Target.Row, 10).Value) ' ACIONAMENTO
    End If

    ' Opcoes vem do dominio exportado pelo LISP (Modulo_Dominios)
    lista = ObterOpcoesDominio(campo, valorPai)

    If lista <> "" Then
        With Target.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=lista
            .InCellDropdown = True
            .IgnoreBlank = True
        End With
    End If
End Sub
