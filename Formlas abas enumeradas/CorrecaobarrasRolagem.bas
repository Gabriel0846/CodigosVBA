Attribute VB_Name = "Módulo1"
Sub LimparFormataçãoDesnecessária()
Attribute LimparFormataçãoDesnecessária.VB_ProcData.VB_Invoke_Func = "d\n14"

    Dim ws As Worksheet
    Dim ultimaColuna As Long
    Dim ultimaLinha As Long
    
    Set ws = ActiveSheet

    ' Encontrar a última coluna usada
    ultimaColuna = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' Encontrar a última linha usada
    ultimaLinha = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    ' Limpar a formatação das colunas além da última usada
    If ultimaColuna < ws.Columns.Count Then
        ws.Range(ws.Cells(1, ultimaColuna + 1), ws.Cells(ws.Rows.Count, ws.Columns.Count)).Clear
    End If

    ' Limpar a formatação das linhas além da última usada
    If ultimaLinha < ws.Rows.Count Then
        ws.Range(ws.Cells(ultimaLinha + 1, 1), ws.Cells(ws.Rows.Count, ws.Columns.Count)).Clear
    End If

    ' Redefinir o intervalo utilizado
    Dim dummy As Range
    Set dummy = ws.Cells(ws.Rows.Count, ws.Columns.Count)
    dummy.Select
    ws.Cells(1, 1).Select ' Reposicionar o cursor

    ' Salvar e fechar a planilha para aplicar as mudanças
    ThisWorkbook.Save
    Application.CutCopyMode = False

    MsgBox "Formatação desnecessária removida e intervalo redefinido."

End Sub

