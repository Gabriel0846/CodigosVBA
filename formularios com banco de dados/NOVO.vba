Sub NOVO()
    ActiveSheet.Unprotect Password:="2015"

    Range("H1:H2").Value = Range("M6").Value
    Range( _
        "C5:D5,C6:F6,C7:F7,B12:D12,B13:D13,B14:D14,B15:D15,B16:D16,B17:D17," & _
        "E12,E13,E14,E15,E16,E17,F12,F13,F14,F15,F16,F17,G12,G13,G14,G15,G16,G17," & _
        "B19:H21,C25,D25,E25,F25" _
        ).ClearContents
    Range("C5:D5").Select
    
    ActiveSheet.Protect Password:="2015"
End Sub