VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_lstEntrada 
   Caption         =   "UserForm1"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6735
   OleObjectBlob   =   "UserForm_lstEntrada.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_lstEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim linhaSelecionada As Long

Private Sub cmdSelecionar_Click()
    If lstResultados.ListIndex >= 0 Then
        linhaSelecionada = lstResultados.List(lstResultados.ListIndex, 1) ' Coluna oculta com número da linha
        Me.Hide
    Else
        MsgBox "Selecione um item da lista.", vbExclamation
    End If
End Sub


Public Function MostrarLista(codigos As Collection) As Long
    Dim item As Variant
    lstResultados.Clear

    For Each item In codigos
        lstResultados.AddItem item(0) ' Coluna I (descrição)
        lstResultados.List(lstResultados.ListCount - 1, 1) = item(1) ' Linha da planilha
    Next item

    linhaSelecionada = 0
    Me.Show
    MostrarLista = linhaSelecionada
End Function

