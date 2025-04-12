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

    Dim numeroProduto As String
    Dim encontrado As Boolean
    Dim wsBD As Worksheet
    Dim wSEntrada As Worksheet
    Dim rng As Range
    Dim cel As Range

    numeroProduto = Me.TextBox1.Value
    encontrado = False

    Set wsBD = ThisWorkbook.Sheets("BANCO_DE_DADOS")
    Set wSEntrada = ThisWorkbook.Sheets("ENTRADA")
    Set rng = wsBD.Range("A1:A100000")

    wsBD.Unprotect Password:="3141"
    wSEntrada.Unprotect Password:="3141"

    For Each cel In rng
        If cel.Value = numeroProduto Then
            encontrado = True
            Exit For
        End If
    Next cel

    If encontrado Then
        wSEntrada.Range("B2").Value = "CONSULTA"
        wSEntrada.Range("D6").Value = numeroProduto
        wSEntrada.Range("U4").Value = numeroProduto
        wSEntrada.Range("D7").Value = wSEntrada.Range("V4")
        wSEntrada.Range("M7").Value = wSEntrada.Range("W4")
        wSEntrada.Range("M6").Value = wSEntrada.Range("X4")
        wSEntrada.Range("D9").Value = wSEntrada.Range("Y4")
        wSEntrada.Range("J9").Value = wSEntrada.Range("Z4")
        wSEntrada.Range("H6").Value = wSEntrada.Range("AA4")
        wSEntrada.Range("D12").Value = wSEntrada.Range("AB4")
    Else
        MsgBox "Produto não encontrado.", vbExclamation
    End If

    wsBD.Protect Password:="3141", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    wSEntrada.Protect Password:="3141"

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

