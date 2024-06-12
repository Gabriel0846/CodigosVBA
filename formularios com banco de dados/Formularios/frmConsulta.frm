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
        wsLancamentos.Range("H1").Value = numeroRequisicao
        
        wsLancamentos.Range("C5:D5").Value = wsLancamentos.Range("N4").Value
        wsLancamentos.Range("C6:F6").Value = wsLancamentos.Range("O4").Value
        wsLancamentos.Range("C7:F7").Value = wsLancamentos.Range("P4").Value
        wsLancamentos.Range("B12:D12").Value = wsLancamentos.Range("Q4").Value
        wsLancamentos.Range("E12").Value = wsLancamentos.Range("R4").Value
        wsLancamentos.Range("F12").Value = wsLancamentos.Range("S4").Value
        wsLancamentos.Range("G12").Value = wsLancamentos.Range("T4").Value
        wsLancamentos.Range("B13:D13").Value = wsLancamentos.Range("U4").Value
        wsLancamentos.Range("E13").Value = wsLancamentos.Range("V4").Value
        wsLancamentos.Range("F13").Value = wsLancamentos.Range("W4").Value
        wsLancamentos.Range("G13").Value = wsLancamentos.Range("X4").Value
        wsLancamentos.Range("B14:D14").Value = wsLancamentos.Range("Y4").Value
        wsLancamentos.Range("E14").Value = wsLancamentos.Range("Z4").Value
        wsLancamentos.Range("F14").Value = wsLancamentos.Range("AA4").Value
        wsLancamentos.Range("G14").Value = wsLancamentos.Range("AB4").Value
        wsLancamentos.Range("B15:D15").Value = wsLancamentos.Range("AC4").Value
        wsLancamentos.Range("E15").Value = wsLancamentos.Range("AD4").Value
        wsLancamentos.Range("F15").Value = wsLancamentos.Range("AE4").Value
        wsLancamentos.Range("G15").Value = wsLancamentos.Range("AF4").Value
        wsLancamentos.Range("B16:D16").Value = wsLancamentos.Range("AG4").Value
        wsLancamentos.Range("E16").Value = wsLancamentos.Range("AH4").Value
        wsLancamentos.Range("F16").Value = wsLancamentos.Range("AI4").Value
        wsLancamentos.Range("G16").Value = wsLancamentos.Range("AJ4").Value
        wsLancamentos.Range("B17:D17").Value = wsLancamentos.Range("AK4").Value
        wsLancamentos.Range("E17").Value = wsLancamentos.Range("AL4").Value
        wsLancamentos.Range("F17").Value = wsLancamentos.Range("AM4").Value
        wsLancamentos.Range("G17").Value = wsLancamentos.Range("AN4").Value
        wsLancamentos.Range("B19:H19").Value = wsLancamentos.Range("AO4").Value
        wsLancamentos.Range("C25").Value = wsLancamentos.Range("AP4").Value
        wsLancamentos.Range("D25").Value = wsLancamentos.Range("AQ4").Value
        wsLancamentos.Range("E25").Value = wsLancamentos.Range("AR4").Value
        wsLancamentos.Range("F25").Value = wsLancamentos.Range("AS4").Value
        wsLancamentos.Range("L21").Value = wsLancamentos.Range("AT4").Value
        wsLancamentos.Range("L22").Value = wsLancamentos.Range("AU4").Value
        wsLancamentos.Range("L23").Value = wsLancamentos.Range("AV4").Value
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

