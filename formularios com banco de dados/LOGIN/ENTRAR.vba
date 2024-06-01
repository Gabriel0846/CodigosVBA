Private Sub btnLogin_Click()
    Dim wsDados As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim usuario As String
    Dim senha As String
    
    usuario = txtUsuario.Text
    senha = txtSenha.Text
    
    If ValidaLogin(usuario, senha) Then
        With ThisWorkbook.Sheets("LANÇAMENTOS")
            .Range("M8").Value = usuario
            .Range("N8").Value = GetSigla(usuario)
        End With
        Me.Hide
    Else
        MsgBox "Usuário ou senha inválidos. Tente novamente.", vbCritical
    End If
End Sub