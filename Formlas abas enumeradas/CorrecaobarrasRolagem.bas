Attribute VB_Name = "M�dulo1"
Sub LimparFormata��oDesnecess�ria()
Attribute LimparFormata��oDesnecess�ria.VB_ProcData.VB_Invoke_Func = "d\n14"

    Dim ws As Worksheet
    Dim ultimaColuna As Long
    Dim ultimaLinha As Long
    
    Set ws = ActiveSheet

    ' Encontrar a �ltima coluna usada
    ultimaColuna = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' Encontrar a �ltima linha usada
    ultimaLinha = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    ' Limpar a formata��o das colunas al�m da �ltima usada
    If ultimaColuna < ws.Columns.Count Then
        ws.Range(ws.Cells(1, ultimaColuna + 1), ws.Cells(ws.Rows.Count, ws.Columns.Count)).Clear
    End If

    ' Limpar a formata��o das linhas al�m da �ltima usada
    If ultimaLinha < ws.Rows.Count Then
        ws.Range(ws.Cells(ultimaLinha + 1, 1), ws.Cells(ws.Rows.Count, ws.Columns.Count)).Clear
    End If

    ' Redefinir o intervalo utilizado
    Dim dummy As Range
    Set dummy = ws.Cells(ws.Rows.Count, ws.Columns.Count)
    dummy.Select
    ws.Cells(1, 1).Select ' Reposicionar o cursor

    ' Salvar e fechar a planilha para aplicar as mudan�as
    ThisWorkbook.Save
    Application.CutCopyMode = False

    MsgBox "Formata��o desnecess�ria removida e intervalo redefinido."

End Sub

