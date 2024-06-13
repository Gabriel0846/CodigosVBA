Attribute VB_Name = "Módulo1"
Sub SALVAR()
        
    Dim wsBancoDeDados As Worksheet
    Dim wsGuiaExames As Worksheet
    Dim nextRow As Long
    
    Set wsBancoDeDados = Sheets("BANCO DE DADOS")
    Set wsGuiaExames = Sheets("GUIA EXAMES")
    
    wsBancoDeDados.Unprotect Password:="2015"
    wsGuiaExames.Unprotect Password:="2015"
    
    wsBancoDeDados.Rows(5).Insert Shift:=xlDown
    wsGuiaExames.Range("AC5:AO5").Copy
    wsBancoDeDados.Range("B5:N5").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    LIMPAR
    
    wsBancoDeDados.Protect Password:="2015", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    wsGuiaExames.Protect Password:="2015"
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
        MsgBox "Não há uma área de impressão definida na aba ativa.", vbExclamation
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

    resposta = MsgBox("Vocâ tem certeza que deseja limpar o conteúdo?", vbYesNo + vbQuestion, "Confirmação")
    
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
