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
Private Sub btnBuscar_Click()
    Dim codigo As String
    Dim wsBD As Worksheet
    Dim wsConsulta As Worksheet
    Dim rng As Range
    Dim celula As Range
    Dim encontrados As Collection
    Dim descricoes As Collection
    Dim linhaSelecionada As Long
    Dim i As Long

    codigo = Trim(Me.txtCodigo.Text)
    
    If codigo = "" Then
        MsgBox "Por favor, digite um código.", vbExclamation
        Exit Sub
    End If

    Set wsBD = ThisWorkbook.Sheets("BD")
    Set wsConsulta = ThisWorkbook.Sheets("consulta")
    Set rng = wsBD.Range("B2:B" & wsBD.Cells(wsBD.Rows.Count, "B").End(xlUp).Row)
    Set encontrados = New Collection
    Set descricoes = New Collection

    ' Buscar todas as ocorrências
    For Each celula In rng
        If Trim(CStr(celula.Value)) = codigo Then
            encontrados.Add celula.Row
            descricoes.Add wsBD.Cells(celula.Row, 9).Value ' Coluna I (descrição)
        End If
    Next celula

    If encontrados.Count = 0 Then
        MsgBox "Código não encontrado na aba 'BD'.", vbExclamation
        Exit Sub
    ElseIf encontrados.Count = 1 Then
        linhaSelecionada = encontrados(1)
    Else
        With UserForm_Selecao
            Set .linhasEncontradas = New Collection ' <-- Adiciona isso
            
            .lstOpcoes.Clear
            For i = 1 To descricoes.Count
                .lstOpcoes.AddItem descricoes(i)
                .linhasEncontradas.Add encontrados(i) ' <-- Copia item por item
            Next i
            
            .Show
            linhaSelecionada = .linhaSelecionada
        End With
        If linhaSelecionada = 0 Then
            MsgBox "Nenhuma opção foi selecionada.", vbInformation
            Exit Sub
        End If
    End If

    ' Copiar da coluna B até O da linha escolhida
    wsBD.Range(wsBD.Cells(linhaSelecionada, 2), wsBD.Cells(linhaSelecionada, 16)).Copy
    wsConsulta.Range("U2").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Define o foco no campo txtCodigo ao abrir o formulário
    Me.txtCodigo.SetFocus
End Sub
Private Sub UserForm_Activate()
    Me.txtCodigo.SetFocus
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub txtCodigo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnBuscar_Click
    End If
End Sub

