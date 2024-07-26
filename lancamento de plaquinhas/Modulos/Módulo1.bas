Attribute VB_Name = "Módulo1"
Sub SALVAR()

    Application.ScreenUpdating = False

    Dim wsBancoDeDados As Worksheet
    Dim wsLancamentos As Worksheet
    Dim wb As Workbook
    Dim encontrado As Boolean
    Dim ultimaLinha As Long
    Dim i As Long
    Dim numPlaquinha As Variant
    Dim linhaEncontrada As Long
    Dim resposta As VbMsgBoxResult
    Dim quantidadePlaquinhas As Variant
    Dim j As Integer
    Dim obs As String

    Set wb = ThisWorkbook
    Set wsBancoDeDados = Sheets("ENTRADA_BD")
    Set wsLancamentos = Sheets("LANÇAMENTOS")
    numPlaquinha = wsLancamentos.Range("B10").Value
    obs = wsLancamentos.Range("F13").Value

    ' Verificar se numPlaquinha é válido
    If IsEmpty(numPlaquinha) Or Not IsNumeric(numPlaquinha) Or numPlaquinha <= 0 Then
        MsgBox "O número da plaquinha é inválido. Por favor, insira um número válido maior que zero.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    wsBancoDeDados.Unprotect Password:="2015"
    wsLancamentos.Unprotect Password:="2015"

    ' Limpar filtros temporariamente na aba de banco de dados, se houver filtros
    On Error Resume Next
    If wsBancoDeDados.FilterMode Then
        wsBancoDeDados.ShowAllData
    End If
    On Error GoTo 0

    encontrado = False
    ultimaLinha = wsBancoDeDados.Cells(wsBancoDeDados.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultimaLinha
        If wsBancoDeDados.Cells(i, "A").Value = numPlaquinha Then
            encontrado = True
            linhaEncontrada = i
            Exit For
        End If
    Next i

    If encontrado Then
        resposta = MsgBox("O número da Plaquinha já existe no banco de dados. Deseja atualizá-lo pelos valores atuais?", vbYesNo + vbQuestion, "Confirmação")

        If resposta = vbYes Then
            wsLancamentos.Range("R5:AE5").Copy
            wsBancoDeDados.Range("A" & linhaEncontrada & ":N" & linhaEncontrada).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            MsgBox "Plaquinha " & numPlaquinha & " atualizado com sucesso."
        Else
            MsgBox "Operação cancelada pelo usuário.", vbInformation
        End If
    Else
        quantidadePlaquinhas = InputBox("Quantas plaquinhas você deseja adicionar?", "Quantidade de Plaquinhas", 1)
        
        If quantidadePlaquinhas = "" Or Not IsNumeric(quantidadePlaquinhas) Or quantidadePlaquinhas <= 0 Then
            MsgBox "Operação cancelada pelo usuário.", vbInformation
        Else
            For j = 1 To quantidadePlaquinhas
                wsBancoDeDados.Rows(3).Insert Shift:=xlDown
                wsLancamentos.Cells(6, "I").Value = numPlaquinha + j - 1
                wsLancamentos.Cells(13, "F").Value = obs & " - PLAQUINHA " & j
                wsLancamentos.Range("R5:AE5").Copy
                wsBancoDeDados.Range("A3:N3").PasteSpecial Paste:=xlPasteValues
                Call Esperar(1)
            Next j
                
            Application.CutCopyMode = False
            wsLancamentos.Activate ' Ativa a planilha LANÇAMENTOS
            LIMPAR
            MsgBox "Plaquinhas registradas com sucesso."
        End If
    End If

    wsBancoDeDados.Protect Password:="2015", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    wsLancamentos.Protect Password:="2015"

    wb.Save
    
    Application.ScreenUpdating = True

End Sub

Sub Esperar(segundos As Double)
    Dim fim As Double
    fim = Timer + segundos
    Do While Timer < fim
        DoEvents ' Permite que o Excel processe outros eventos
    Loop
End Sub
Sub LIMPAR()
    Range( _
        "F7:G7,I6,F9,F11:I11,F13:I13,F16,F17,H16:I16,H17:I19,R7" _
        ).Select
    Selection.ClearContents
    Range("I6:I8").Select
End Sub

Sub BAIXA()
    Dim wsLc As Worksheet

    Set wsLc = Sheets("LANÇAMENTOS")
    wsLc.Unprotect Password:="2015"
    
    Range("F16").Value = Range("R2").Value
    Range("F17").Value = Range("U3").Value
    
    wsLc.Protect Password:="2015"
End Sub

Sub NOVO()

    Set wsGuiaExames = Sheets("LANÇAMENTOS")
    
    wsGuiaExames.Unprotect Password:="2015"

    LIMPAR
    Range("R7").ClearContents
    
    wsGuiaExames.Protect Password:="2015"
End Sub

Sub CONSULTA()
    frmConsulta.Show
End Sub


