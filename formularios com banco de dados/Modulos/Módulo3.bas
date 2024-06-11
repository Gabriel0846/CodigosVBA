Attribute VB_Name = "Módulo3"
Sub LIMPAR()
    Range( _
        "C5:D5,C6:F6,C7:F7,B12:D12,B13:D13,B14:D14,B15:D15,B16:D16,B17:D17," & _
        "E12,E13,E14,E15,E16,E17,F12,F13,F14,F15,F16,F17,G12,G13,G14,G15,G16,G17," & _
        "B19:H19,B20:H20,B21:H21,C25,D25,E25,F25" _
        ).Select
    Selection.ClearContents
End Sub

Sub LIMPAR_btn()
    Dim resposta As VbMsgBoxResult
    Dim rng1 As Range
    Dim rng2 As Range
    Dim rng3 As Range
    Dim rng4 As Range
    Dim rng5 As Range

    resposta = MsgBox("Você tem certeza que deseja limpar o conteúdo?", vbYesNo + vbQuestion, "Confirmação")
    
    If resposta = vbYes Then
        Set rng1 = Range("C5:D5,C6:F6,C7:F7,B12:D12,B13:D13,B14:D14,B15:D15,B16:D16,B17:D17")
        Set rng2 = Range("E12,E13,E14,E15,E16,E17,F12,F13,F14,F15,F16,F17,G12,G13,G14,G15,G16,G17")
        Set rng3 = Range("B19:H19,B20:H20,B21:H21")
        Set rng4 = Range("C25,D25,E25,F25")

        Union(rng1, rng2, rng3, rng4).ClearContents
    Else
        MsgBox "A operação foi cancelada.", vbInformation, "Operação Cancelada"
    End If
End Sub


