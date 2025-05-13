Attribute VB_Name = "Módulo1"
Sub AbrirFormularioDeBusca()
    frmConsulta.txtCodigo.Text = ""
    frmConsulta.Show
End Sub

Sub LIMPAR_CONSULTA()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CONSULTA")

    ws.Unprotect

    ws.Range("U2:AI2").ClearContents

    ws.Protect

End Sub

Sub LIMPAR_ENTRADA()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ENTRADA")
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Tem certeza que deseja limpar os campos?", vbYesNo + vbQuestion, "Confirmação")

    If resposta = vbNo Then Exit Sub
    
    ws.Unprotect
    
    ws.Range( _
        "E6:G6,K6:M6,D7:E7,M7:N7,D8:J8,M8:N8,E9:F9,H9:I9,K9:L9,E10:F10,I10:M10,F14:G14,I14,K14,M14" _
        ).ClearContents
    
    ws.Protect
End Sub

Sub NOVO()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ENTRADA")

    ws.Unprotect
    
    ws.Range( _
        "E6:G6,K6:M6,D7:E7,M7:N7,D8:J8,M8:N8,E9:F9,H9:I9,K9:L9,E10:F10,I10:M10,F14:G14,I14,K14,M14" _
        ).ClearContents

    ws.Range("T4:AI4").ClearContents

    ws.Range("B2:O3").Value = "NOVO"

    Dim valorConcatenado As String
    ws.Range("H7:I7").Value = ws.Range("ac6").Value

    ws.Protect

End Sub
Sub SalvarProduto()
    Dim wsBD As Worksheet, wsEntrada As Worksheet
    Dim modo As String
    Dim ultimaLinha As Long, i As Long, linhaEncontrada As Long
    Dim dados As Variant

    Set wsBD = ThisWorkbook.Sheets("BD")
    If wsBD.AutoFilterMode Then wsBD.AutoFilterMode = False

    Set wsEntrada = ThisWorkbook.Sheets("ENTRADA")
    
    wsBD.Unprotect
    wsEntrada.Unprotect

    ' Detecta se é NOVO ou BUSCA
    modo = Trim(wsEntrada.Range("B2").Value)

    If UCase(modo) = "NOVO" Then
        ' Copia T2:AI2 -> A:P (valores) na próxima linha vazia
        dados = wsEntrada.Range("T2:AI2").Value
        ultimaLinha = wsBD.Cells(wsBD.Rows.Count, "A").End(xlUp).Row + 1
        wsBD.Range("A" & ultimaLinha & ":P" & ultimaLinha).Value = dados
        MsgBox "Novo item cadastrado no Banco de Dados.", vbInformation
        wsEntrada.Activate

    ElseIf UCase(modo) = "BUSCA" Then
        ' Pega o valor de T4 e busca na coluna A do BD
        Dim chaveBusca As String
        chaveBusca = Trim(wsEntrada.Range("T4").Value)
        
        ultimaLinha = wsBD.Cells(wsBD.Rows.Count, "A").End(xlUp).Row
        linhaEncontrada = 0

        For i = 2 To ultimaLinha
            If Trim(wsBD.Cells(i, "A").Value) = chaveBusca Then
                linhaEncontrada = i
                Exit For
            End If
        Next i

        If linhaEncontrada > 0 Then
            Dim resposta As VbMsgBoxResult
            resposta = MsgBox("Tem certeza que deseja atualizar este item?", vbYesNo + vbQuestion, "Confirmar Atualização")

            If resposta = vbYes Then
                dados = wsEntrada.Range("T2:AI2").Value
                wsBD.Range("A" & linhaEncontrada & ":P" & linhaEncontrada).Value = dados
                MsgBox "Item atualizado no Banco de Dados.", vbInformation
                wsEntrada.Activate
            Else
                MsgBox "Atualização cancelada.", vbInformation
            End If
        End If
    End If

    wsBD.Protect
    wsEntrada.Protect
End Sub


Sub BuscarProdutoComForm()
    Dim wsBD As Worksheet, wsEntrada As Worksheet
    Dim codBusca As String
    Dim ultimaLinha As Long, i As Long
    Dim achados As Collection
    Dim linhaDestino As Long
    Dim dados(1 To 1, 1 To 16) As Variant

    Set wsBD = ThisWorkbook.Sheets("BD")
    If wsBD.AutoFilterMode Then wsBD.AutoFilterMode = False

    Set wsEntrada = ThisWorkbook.Sheets("ENTRADA")
    Set achados = New Collection
    
    wsEntrada.Unprotect

    ' Abre o formulário de entrada de código
    frmEntradaCodigo.Show
    codBusca = frmEntradaCodigo.CodigoDigitado

    If codBusca = "" Then Exit Sub

    ' Procura o código na aba BD
    ultimaLinha = wsBD.Cells(wsBD.Rows.Count, "B").End(xlUp).Row
    For i = 2 To ultimaLinha
        If Trim(wsBD.Cells(i, "B").Value) = codBusca Then
            achados.Add Array(wsBD.Cells(i, "I").Value, i) ' descrição e linha
        End If
    Next i

    If achados.Count = 0 Then
    MsgBox "Código '" & codBusca & "' não foi encontrado na base de dados.", vbExclamation, "Produto não encontrado"
    Exit Sub
    End If

    If achados.Count = 1 Then
        linhaDestino = achados(1)(1)
    Else
        linhaDestino = UserForm_lstEntrada.MostrarLista(achados)
        If linhaDestino = 0 Then Exit Sub
    End If

    ' Copia os dados da linha A:P para T2:AI2
    For i = 1 To 16
        dados(1, i) = wsBD.Cells(linhaDestino, i).Value
    Next i
    wsEntrada.Range("T4:AI4").Value = dados
    wsEntrada.Range("H7").Value = wsEntrada.Range("T4").Value
    wsEntrada.Range("D7").Value = wsEntrada.Range("U4").Value
    wsEntrada.Range("E6:G6").Value = wsEntrada.Range("V4").Value
    wsEntrada.Range("M7:N7").Value = wsEntrada.Range("W4").Value
    wsEntrada.Range("M8:N8").Value = wsEntrada.Range("X4").Value
    wsEntrada.Range("K6:M6").Value = wsEntrada.Range("Y4").Value
    wsEntrada.Range("I10:M10").Value = wsEntrada.Range("Z4").Value
    wsEntrada.Range("E10:F10").Value = wsEntrada.Range("AA4").Value
    wsEntrada.Range("D8:J8").Value = wsEntrada.Range("AB4").Value
    wsEntrada.Range("E9:F9").Value = wsEntrada.Range("AC4").Value
    wsEntrada.Range("H9:I9").Value = wsEntrada.Range("AD4").Value
    wsEntrada.Range("K9:L9").Value = wsEntrada.Range("AE4").Value
    wsEntrada.Range("F14:G14").Value = wsEntrada.Range("AF4").Value
    wsEntrada.Range("I14").Value = wsEntrada.Range("AG4").Value
    wsEntrada.Range("K14").Value = wsEntrada.Range("AH4").Value
    wsEntrada.Range("M14").Value = wsEntrada.Range("AI4").Value
    wsEntrada.Range("B2:O3").Value = "BUSCA"
    
    wsEntrada.Protect
End Sub


