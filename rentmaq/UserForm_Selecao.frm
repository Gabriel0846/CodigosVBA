VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Selecao 
   Caption         =   "UserForm1"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   OleObjectBlob   =   "UserForm_Selecao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Selecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public linhasEncontradas As Collection
Public linhaSelecionada As Long

Private Sub btnSelecionar_Click()
    If lstOpcoes.ListIndex = -1 Then
        MsgBox "Selecione uma das opções.", vbExclamation
        Exit Sub
    End If
    
    linhaSelecionada = linhasEncontradas(lstOpcoes.ListIndex + 1)
    Me.Hide
End Sub

