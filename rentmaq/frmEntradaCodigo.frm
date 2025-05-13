VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntradaCodigo 
   Caption         =   "BUSCAR"
   ClientHeight    =   1350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   OleObjectBlob   =   "frmEntradaCodigo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEntradaCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoDigitado As String

Private Sub btnBuscar_Click()
    If Trim(txtCodigo.Value) = "" Then
        MsgBox "Digite um código!", vbExclamation
    Else
        CodigoDigitado = txtCodigo.Value
        Me.Hide
    End If
End Sub

Private Sub btnCancelar_Click()
    CodigoDigitado = ""
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    ' Define o foco no campo txtCodigo ao abrir o formulário
    Me.txtCodigo.SetFocus
End Sub
Private Sub UserForm_Activate()
    Me.txtCodigo.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnBuscar_Click
    End If
End Sub
