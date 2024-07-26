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

    Set wsBD = ThisWorkbook.Sheets("ENTRADA_BD")
    Set wsLancamentos = ThisWorkbook.Sheets("LANÇAMENTOS")
    Set rng = wsBD.Range("A1:A60000")

    wsBD.Unprotect Password:="2015"
    wsLancamentos.Unprotect Password:="2015"

    For Each cel In rng
        If cel.Value = numeroRequisicao Then
            encontrado = True
            Exit For
        End If
    Next cel

    If encontrado Then
        LIMPAR
        wsLancamentos.Range("R7").Value = numeroRequisicao
        
        wsLancamentos.Range("I6").Value = wsLancamentos.Range("R6")
        wsLancamentos.Range("F7").Value = wsLancamentos.Range("S6")
        wsLancamentos.Range("F9").Value = wsLancamentos.Range("T6")
        wsLancamentos.Range("F11").Value = wsLancamentos.Range("X6")
        wsLancamentos.Range("F13").Value = wsLancamentos.Range("Y6")
        wsLancamentos.Range("F16").Value = wsLancamentos.Range("Z6")
        wsLancamentos.Range("F17").Value = wsLancamentos.Range("AA6")
        wsLancamentos.Range("H16").Value = wsLancamentos.Range("AC6")
        wsLancamentos.Range("H17").Value = wsLancamentos.Range("AD6")
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

