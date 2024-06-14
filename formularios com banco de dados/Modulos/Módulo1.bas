Attribute VB_Name = "M�dulo1"
Sub SALVAR()
        
    Dim wsBancoDeDados As Worksheet
    Dim wsGuiaExames As Worksheet
    Dim nextRow As Long
    Dim wb As Workbook
    Dim encontrado As Boolean
    Dim ultimaLinha As Long
    Dim i As Long
    Dim numExame As Variant
    
    Set wb = ThisWorkbook
    Set wsBancoDeDados = Sheets("BANCO DE DADOS")
    Set wsGuiaExames = Sheets("GUIA EXAMES")
    Set numExame = wsGuiaExames.Range("B12").Value
    
    wsBancoDeDados.Unprotect Password:="2015"
    wsGuiaExames.Unprotect Password:="2015"

    encontrado = False
    ultimaLinha = wsBancoDeDados.Cells(wsBancoDeDados.Rows.Count, "A").End(xlUp).Row

    For i = l To ultimaLinha
        If wsBancoDeDados.Cells(i,1).Value = valorH1 Then
            encontrado = True
            linhaEncontrada = i
            Exit For
        End If
    Next i

    If encontrado Then
        resposta = MsgBox("O número do exame já existe no banco de dados. Deseja atualiza-lo pelos valores atuais?", vbYesNo + vbQuestion, "Confirmação")

        If resposta = vbYes Then
            wsGuiaExames.Range("AC5:AO5").Copy
            wsBancoDeDados.Rows(linhaEncontrada).PasteSpecial Paste:xlPasteValues
            Application.CutCopyMode = False
            MsgBox "Exame " & NumExame & " registrado com sucesso"
        Else
            MsgBox "Operação cancelada pelo usuário.", vbInformation
        End If
    Else
        wsBancoDeDados.Rows(5).Insert Shift:=xlDown
        wsGuiaExames.Range("AC5:AO5").Copy
        wsBancoDeDados.Range("B5:N5").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        LIMPAR
        MsgBox "Exame " & numExame & " registrada com sucesso."
    End If
    
    wsBancoDeDados.Protect Password:="2015", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    wsGuiaExames.Protect Password:="2015"
    
    wb.Save
    
End Sub

Sub LIMPAR()
    Range( _
        "P2,F15:I15,M15:N15,E18,E20,E22,E24,E36,H18,H22,K24,I30:N30,I32:N32,I34:N34,I36:N36" _
        ).Select
    Selection.ClearContents
    Range("P2").Select
End Sub

Sub IMPRIMIR()
    If ActiveSheet.PageSetup.PrintArea <> "" Then
        ActiveSheet.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    Else
        MsgBox "N�o h� uma �rea de impress�o definida na aba ativa.", vbExclamation
    End If
End Sub

Sub NOVO()

    Set wsGuiaExames = Sheets("GUIA EXAMES")
    
    wsGuiaExames.Unprotect Password:="2015"

    Range("B12").Value = Range("AD2").Value
    LIMPAR
    Range("AC6").ClearContents
    Range("F15:I15").ClearContents
    Range("M15:N15").ClearContents
    Range("P2").Select
    
    
    wsGuiaExames.Protect Password:="2015"
End Sub

Sub LIMPAR_btn()
    Dim resposta As VbMsgBoxResult
    Dim rng1 As Range
    Dim rng2 As Range
    Dim rng3 As Range
    Dim rng4 As Range
    Dim rng5 As Range

    resposta = MsgBox("Voc� tem certeza que deseja limpar o conte�do?", vbYesNo + vbQuestion, "Confirma��o")
    
    ActiveSheet.Unprotect Password:="2015"
    If resposta = vbYes Then
        Range( _
        "P2,F15:I15,M15:N15,E18,E20,E22,E24,E36,H18,H22,K24,I30:N30,I32:N32,I34:N34,I36:N36" _
        ).Select
    Selection.ClearContents
    End If
    Range("P2").Select
    ActiveSheet.Protect Password:="2015"
End Sub

Sub CONSULTA()
    frmConsulta.Show
End Sub
