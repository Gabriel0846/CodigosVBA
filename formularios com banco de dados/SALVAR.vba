Sub SALVAR()
    Dim wsBD As Worksheet
    Dim wsLancamentos As Worksheet
    Dim valorH1 As Variant
    Dim encontrado As Boolean
    Dim ultimaLinha As Long
    Dim i As Long
    Dim resposta As VbMsgBoxResult
    Dim linhaEncontrada As Long
    
    Set wsBD = ThisWorkbook.Sheets("BD")
    Set wsLancamentos = ThisWorkbook.Sheets("LANÇAMENTOS")
    
    wsBD.Unprotect Password:="2015"
    wsLancamentos.Unprotect Password:="2015"
    
    valorH1 = wsLancamentos.Range("H1").Value
    
    encontrado = False
    
    ultimaLinha = wsBD.Cells(wsBD.Rows.Count, "A").End(xlUp).Row
    
    For i = 1 To ultimaLinha
        If wsBD.Cells(i, 1).Value = valorH1 Then
            encontrado = True
            linhaEncontrada = i
            Exit For
        End If
    Next i
    
    If encontrado Then
        resposta = MsgBox("O número da requisição já existe no banco de dados. Deseja atualiza-lo pelos valores atuais?", vbYesNo + vbQuestion, "Confirmação")
        
        If resposta = vbYes Then
            wsLancamentos.Range("M2:AS2").Copy
            wsBD.Rows(linhaEncontrada).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            MsgBox "Requisição " & valorH1 & " atualizada no banco de dados."
        Else
            MsgBox "Operação cancelada pelo usuário.", vbInformation
        End If
    Else
        wsLancamentos.Range("M2:AS2").Copy
        wsBD.Rows(2).Insert Shift:=xlDown
        wsBD.Rows(2).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        wsLancamentos.Range("H1").Value = wsLancamentos.Range("H1").Value + 1
        MsgBox "Requisição " & valorH1 & " registrada com sucesso."
    End If
    
    wsBD.Protect Password:="2015", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    wsLancamentos.Protect Password:="2015"

    LIMPAR
    
End Sub