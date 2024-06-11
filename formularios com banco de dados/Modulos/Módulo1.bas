Attribute VB_Name = "M�dulo1"
Sub SALVAR()
    Dim wsBD As Worksheet
    Dim wsLancamentos As Worksheet
    Dim valorH1 As Variant
    Dim encontrado As Boolean
    Dim ultimaLinha As Long
    Dim i As Long
    Dim resposta As VbMsgBoxResult
    Dim linhaEncontrada As Long
    Dim usuario As String
    Dim sigla As String
    
    Set wsBD = ThisWorkbook.Sheets("BD")
    Set wsLancamentos = ThisWorkbook.Sheets("LAN�AMENTOS")
    
    ' Obter o usu�rio e a sigla atual
    usuario = wsLancamentos.Range("M8").Value
    sigla = wsLancamentos.Range("N8").Value
    
    ' Verificar permiss�es
    If Not TemPermissaoLancar(usuario, sigla) Then
        MsgBox "Voc� n�o tem permiss�o para lan�ar dados.", vbCritical
        Exit Sub
    End If
    
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
        resposta = MsgBox("O n�mero da requisi��o j� existe no banco de dados. Deseja atualiza-lo pelos valores atuais?", vbYesNo + vbQuestion, "Confirma��o")
        
        If resposta = vbYes Then
            wsLancamentos.Range("M2:AV2").Copy
            wsBD.Rows(linhaEncontrada).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            MsgBox "Requisi��o " & valorH1 & " atualizada no banco de dados."
        Else
            MsgBox "Opera��o cancelada pelo usu�rio.", vbInformation
        End If
    Else
        wsLancamentos.Range("M2:AV2").Copy
        wsBD.Rows(2).Insert Shift:=xlDown
        wsBD.Rows(2).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        wsLancamentos.Range("H1").Value = wsLancamentos.Range("H1").Value + 1
        MsgBox "Requisi��o " & valorH1 & " registrada com sucesso."
    End If
    
    wsBD.Protect Password:="2015", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    wsLancamentos.Protect Password:="2015"

    LIMPAR
    
End Sub

Function TemPermissaoLancar(usuario As String, sigla As String) As Boolean
    Dim wsDados As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim permissao As Boolean
    
    Set wsDados = ThisWorkbook.Sheets("DADOS")
    
    ultimaLinha = wsDados.Cells(wsDados.Rows.Count, "A").End(xlUp).Row
    permissao = False
    
    For i = 2 To ultimaLinha
        If wsDados.Cells(i, 1).Value = usuario And wsDados.Cells(i, 2).Value = sigla Then
            If wsDados.Cells(i, 5).Value = 1 Then
                permissao = True
            End If
            Exit For
        End If
    Next i
    
    TemPermissaoLancar = permissao
End Function

Sub NOVO()
    ActiveSheet.Unprotect Password:="2015"

    Range("H1").Value = Range("M6").Value
    LIMPAR
    Range("L21").Value = Range("M8").Value
    Range("L22, L23").ClearContents
    Range("C5:D5").Select
    
    
    ActiveSheet.Protect Password:="2015"
End Sub


Sub CONSULTA()
    frmConsulta.Show
End Sub
