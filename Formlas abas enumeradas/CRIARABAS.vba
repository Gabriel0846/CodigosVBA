Sub CriarAbasCopiando()
    Dim quantidade As Integer
    Dim i As Integer
    Dim abaBase As Worksheet
    Dim novaAba As Worksheet
    Dim novoNumero As Integer
    Dim senha As String

    senha = "2015"

    quantidade = Application.InputBox("Digite a quantidade de abas que deseja criar:", Type:=1)
    If quantidade <= 0 Then
        MsgBox "Por favor, insira um número válido maior que 0."
        Exit Sub
    End If

    Set abaBase = ThisWorkbook.Sheets("BRANCO")

    ' numero inial da criação das abas
    novoNumero = 5301

    For i = 1 To quantidade
        abaBase.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set novaAba = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        novaAba.Unprotect Password:=senha
        novaAba.Name = CStr(novoNumero)
        novaAba.Range("N3").Value = novoNumero
        novaAba.Protect Password:=senha
        novoNumero = novoNumero + 1
    Next i
End Sub

