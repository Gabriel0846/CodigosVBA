Private Sub BUSCAR_Click()
    Application.ScreenUpdating = False

    Dim numeroRequisicao As String
    Dim encontrado As Boolean
    Dim ws As Worksheet
    Dim rng As Range
    Dim cel As Range
    
    numeroRequisicao = Me.TextBox1.Value
    encontrado = False
    
    Set ws = ThisWorkbook.Sheets("BD")
    Set rng = ws.Range("A1:A10000")
    
    For Each cel In rng
        If cel.Value = numeroRequisicao Then
            encontrado = True
            Exit For
        End If
    Next cel
    
    If encontrado Then
        ThisWorkbook.Sheets("LANÇAMENTOS").Range("H1").Value = numeroRequisicao
        Range("N4").Select
        Selection.Copy
        Range("C5:D5").Select
        ActiveSheet.Paste
        Range("O4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("C6:F6").Select
        ActiveSheet.Paste
        Range("P4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("C7:F7").Select
        ActiveSheet.Paste
        Range("Q4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("B12:D12").Select
        ActiveSheet.Paste
        Range("R4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("E12").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("S4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("F12").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("T4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("G12").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("U4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("B13:D13").Select
        Range("V4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("U4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("B13:D13").Select
        ActiveSheet.Paste
        Range("V4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("E13").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("W4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("F13").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("X4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("G13").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("Y4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("B14:D14").Select
        ActiveSheet.Paste
        Range("Z4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("E14").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AA4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("F14").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AB4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("G14").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AC4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("B15:D15").Select
        ActiveSheet.Paste
        Range("AD4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("E15").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AE4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("F15").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AF4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("G15").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AG4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("B16:D16").Select
        ActiveSheet.Paste
        Range("AH4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("E16").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AI4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("F16").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AJ4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("G16").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AK4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("B17:D17").Select
        ActiveSheet.Paste
        Range("AL4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("E17").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AM4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("F17").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AN4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("G17").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AO4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("B19:H19").Select
        ActiveSheet.Paste
        Range("AP4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("C25").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AQ4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("D25").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AR4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("E25").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("AS4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Range("F25").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("C5:D5").Select
    Else
        MsgBox "Número de Requisição não encontrado.", vbExclamation
    End If
    
    Unload Me
End Sub