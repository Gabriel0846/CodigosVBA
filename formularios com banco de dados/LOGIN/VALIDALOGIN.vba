Function ValidaLogin(usuario As String, senha As String) As Boolean
    Dim wsDados As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim loginCorreto As Boolean
    
    Set wsDados = ThisWorkbook.Sheets("DADOS")
    
    ultimaLinha = wsDados.Cells(wsDados.Rows.Count, "B").End(xlUp).Row
    loginCorreto = False
    
    For i = 2 To ultimaLinha ' Assumindo que os dados começam na linha 2
        If wsDados.Cells(i, 1).Value = usuario And wsDados.Cells(i, 4).Value = senha Then
            loginCorreto = True
            Exit For
        End If
    Next i
    
    ValidaLogin = loginCorreto
End Function