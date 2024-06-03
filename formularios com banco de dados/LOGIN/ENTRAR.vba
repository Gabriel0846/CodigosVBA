Private Sub btnLogin_Click()
    Dim wsDados As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim usuario As String
    Dim senha As String
    
    usuario = UCase(txtUsuario.Text)
    senha = txtSenha.Text
    Set wsLancamentos = ThisWorkbook.Sheets("LANÇAMENTOS")
    
    wsLancamentos.Unprotect Password:="2015"
    
    If ValidaLogin(usuario, senha) Then
        With ThisWorkbook.Sheets("LANÇAMENTOS")
            .Range("M8").Value = usuario
            .Range("N8").Value = GetSigla(usuario)
        End With
        Me.Hide
    Else
        MsgBox "Usuário ou senha inválidos. Tente novamente.", vbCritical
    End If
    
    wsLancamentos.Protect Password:="2015"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

Function ValidaLogin(usuario As String, senha As String) As Boolean
    Dim wsDados As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim loginCorreto As Boolean
    
    Set wsDados = ThisWorkbook.Sheets("DADOS")
    
    ultimaLinha = wsDados.Cells(wsDados.Rows.Count, "A").End(xlUp).Row
    loginCorreto = False
    
    For i = 2 To ultimaLinha
        If wsDados.Cells(i, 1).Value = usuario And wsDados.Cells(i, 4).Value = senha Then
            loginCorreto = True
            Exit For
        End If
    Next i
    
    ValidaLogin = loginCorreto
End Function

Function GetSigla(usuario As String) As String
    Dim wsDados As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim sigla As String
    
    Set wsDados = ThisWorkbook.Sheets("DADOS")
    
    ultimaLinha = wsDados.Cells(wsDados.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To ultimaLinha
        If wsDados.Cells(i, 1).Value = usuario Then
            sigla = wsDados.Cells(i, 2).Value
            Exit For
        End If
    Next i
    
    GetSigla = sigla
End Function
