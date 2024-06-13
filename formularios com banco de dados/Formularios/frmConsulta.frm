VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConsulta 
   Caption         =   "CONSULTAR"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmConsulta.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BUSCAR_Click()
    Application.ScreenUpdating = False

    Dim numeroRequisicao As String
    Dim encontrado As Boolean
    Dim wsBD As Worksheet
    Dim wsLancamentos As Worksheet
    Dim rng As Range
    Dim cel As Range

    numeroRequisicao = Me.TextBox1.Value
    encontrado = False

    Set wsBD = ThisWorkbook.Sheets("BANCO DE DADOS")
    Set wsLancamentos = ThisWorkbook.Sheets("GUIA EXAMES")
    Set rng = wsBD.Range("B1:B10000")

    wsBD.Unprotect Password:="2015"
    wsLancamentos.Unprotect Password:="2015"

    For Each cel In rng
        If cel.Value = numeroRequisicao Then
            encontrado = True
            Exit For
        End If
    Next cel

    If encontrado Then
        wsLancamentos.Range("AC6").Value = numeroRequisicao
        wsLancamentos.Range("B12").Value = numeroRequisicao
        wsLancamentos.Range("F15").Value = wsLancamentos.Range("AH6")
        wsLancamentos.Range("M15").Value = wsLancamentos.Range("AK6")
        wsLancamentos.Range("I30").Value = wsLancamentos.Range("AP6")
        If wsLancamentos.Range("AM6") = wsLancamentos.Range("F18") Then
            wsLancamentos.Range("E18") = "X"
        ElseIf wsLancamentos.Range("AM6") = wsLancamentos.Range("F20") Then
            wsLancamentos.Range("E20") = "X"
        ElseIf wsLancamentos.Range("AM6") = wsLancamentos.Range("F22") Then
            wsLancamentos.Range("E22") = "X"
        ElseIf wsLancamentos.Range("AM6") = wsLancamentos.Range("F24") Then
            wsLancamentos.Range("E24") = "X"
        ElseIf wsLancamentos.Range("AM6") = wsLancamentos.Range("I18") Then
            wsLancamentos.Range("H18") = "X"
        ElseIf wsLancamentos.Range("AM6") = wsLancamentos.Range("I22") Then
            wsLancamentos.Range("H22") = "X"
        Else
            wsLancamentos.Range("K24") = "X"
        End If
    Else
        MsgBox "Número de Requisição não encontrado.", vbExclamation
    End If

    wsBD.Protect Password:="2015", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    wsLancamentos.Protect Password:="2015"

    Application.ScreenUpdating = True
    Unload Me
End Sub


Private Sub CANCELAR_Click()
Unload Me
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()
End Sub

