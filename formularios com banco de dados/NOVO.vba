Sub NOVO()
    ActiveSheet.Unprotect Password:="2015"

    Range("H1").Value = Range("M6").Value
    LIMPAR
    Range("L21").Value = Range("M8").Value
    Range("L22, L23").ClearContents
    Range("C5:D5").Select
    
    
    ActiveSheet.Protect Password:="2015"
End Sub