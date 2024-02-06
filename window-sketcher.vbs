Sub RysownikOkien()

' Okno 1
' Okno 1
' Okno 1

    Okno1Podzial = Worksheets("Dane").Range("F2").Value
    Okno1Strona = Worksheets("Dane").Range("G2").Value
    Okno1Uchyl = Worksheets("Dane").Range("H2").Value


' Podział dwustronny
    If Okno1Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.25, 0, 54.25, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 0, 109, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 109, 0, 0, 129). _
        Select

' Stałe prawe
    ElseIf Okno1Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.25, 0, 54.25, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 0, 54.5, 65). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.5, 63.5, 0, 127). _
        Select

' Stałe lewe
    ElseIf Okno1Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.25, 0, 54.25, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 109, 0, 54.5, 65). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.5, 65, 109, 130). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno1Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 0, 109, 65). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 109, 65, 0, 130). _
        Select
    
' Prawa
    ElseIf Okno1Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 109, 0, 0, 65). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 65, 109, 130). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno1Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 130, 55, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 55, 0, 108, 130). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno1Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 130, 27.5, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 27.5, 0, 55, 130). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno1Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 80.5, 0, 55, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 80.5, 0, 108, 130). _
        Select
    
' Uchył obustronny
    ElseIf Okno1Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 130, 27.5, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 27.5, 0, 55, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 80.5, 0, 55, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 80.5, 0, 108, 130). _
        Select
    End If

' Okno 2
' Okno 2
' Okno 2

    Okno2Podzial = Worksheets("Dane").Range("F3").Value
    Okno2Strona = Worksheets("Dane").Range("G3").Value
    Okno2Uchyl = Worksheets("Dane").Range("H3").Value


' Podział dwustronny
    If Okno2Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 0, 208.25, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 0, 263, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 263, 0, 154, 129). _
        Select

' Stałe prawe
    ElseIf Okno2Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 0, 208.25, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 0, 208.25, 65). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 63.5, 154, 127). _
        Select

' Stałe lewe
    ElseIf Okno2Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 0, 208.25, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 263, 0, 208.25, 65). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 65, 263, 130). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno2Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 0, 263, 65). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 263, 65, 154, 130). _
        Select
    
' Prawa
    ElseIf Okno2Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 263, 0, 154, 65). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 65, 263, 130). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno2Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 130, 209, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 209, 0, 262, 130). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno2Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 130, 181.5, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 181.5, 0, 209, 130). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno2Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 234.5, 0, 209, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 234.5, 0, 262, 130). _
        Select

' Uchył obustronny
    ElseIf Okno2Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 130, 181.5, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 181.5, 0, 209, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 234.5, 0, 209, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 234.5, 0, 262, 130). _
        Select
    End If
    
' Okno 3
' Okno 3
' Okno 3

    Okno3Podzial = Worksheets("Dane").Range("F4").Value
    Okno3Strona = Worksheets("Dane").Range("G4").Value
    Okno3Uchyl = Worksheets("Dane").Range("H4").Value


' Podział dwustronny
    If Okno3Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 362.25, 0, 362.25, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 0, 417, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 417, 0, 308, 129). _
        Select

' Stałe prawe
    ElseIf Okno3Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 362.25, 0, 362.25, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 0, 362.25, 65). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 362.25, 63.5, 308, 127). _
        Select

' Stałe lewe
    ElseIf Okno3Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 362.25, 0, 362.25, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 417, 0, 362.25, 65). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 362.25, 65, 417, 130). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno3Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 0, 417, 65). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 417, 65, 308, 130). _
        Select
    
' Prawa
    ElseIf Okno3Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 417, 0, 308, 65). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 65, 417, 130). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno3Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 130, 363, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 363, 0, 416, 130). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno3Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 130, 335.5, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 335.5, 0, 363, 130). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno3Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 388.5, 0, 363, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 388.5, 0, 416, 130). _
        Select

' Uchył obustronny
    ElseIf Okno3Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 130, 335.5, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 335.5, 0, 363, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 388.5, 0, 363, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 388.5, 0, 416, 130). _
        Select
    End If
    
' Okno 4
' Okno 4
' Okno 4

    Okno4Podzial = Worksheets("Dane").Range("F5").Value
    Okno4Strona = Worksheets("Dane").Range("G5").Value
    Okno4Uchyl = Worksheets("Dane").Range("H5").Value


' Podział dwustronny
    If Okno4Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.25, 0, 516.25, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 0, 571, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 571, 0, 462, 129). _
        Select

' Stałe prawe
    ElseIf Okno4Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.5, 0, 516.5, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 0, 516.5, 65). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.5, 63.5, 462, 127). _
        Select

' Stałe lewe
    ElseIf Okno4Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.5, 0, 516.5, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 571, 0, 516.5, 65). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.5, 65, 571, 130). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno4Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 0, 571, 65). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 571, 65, 462, 130). _
        Select
    
' Prawa
    ElseIf Okno4Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 571, 0, 462, 65). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 65, 571, 130). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno4Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 130, 517, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 517, 0, 570, 130). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno4Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 130, 489.5, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 489.5, 0, 517, 130). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno4Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 542.5, 0, 517, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 542.5, 0, 570, 130). _
        Select

' Uchył obustronny
    ElseIf Okno4Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 130, 489.5, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 489.5, 0, 517, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 542.5, 0, 517, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 542.5, 0, 570, 130). _
        Select
    End If
    
' Okno 5
' Okno 5
' Okno 5

    Okno5Podzial = Worksheets("Dane").Range("F6").Value
    Okno5Strona = Worksheets("Dane").Range("G6").Value
    Okno5Uchyl = Worksheets("Dane").Range("H6").Value


' Podział dwustronny
    If Okno5Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.25, 0, 668.25, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 0, 723, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 723, 0, 614, 129). _
        Select

' Stałe prawe
    ElseIf Okno5Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.5, 0, 668.5, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 0, 668.5, 65). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.5, 63.5, 614, 127). _
        Select

' Stałe lewe
    ElseIf Okno5Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.5, 0, 668.5, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 723, 0, 668.5, 65). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.5, 65, 723, 130). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno5Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 0, 723, 65). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 723, 65, 614, 130). _
        Select
    
' Prawa
    ElseIf Okno5Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 723, 0, 614, 65). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 65, 723, 130). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno5Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 130, 669, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 669, 0, 722, 130). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno5Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 130, 641.5, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 641.5, 0, 669, 130). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno5Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 694.5, 0, 669, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 694.5, 0, 722, 130). _
        Select

' Uchył obustronny
    ElseIf Okno5Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 130, 641.5, 0). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 641.5, 0, 669, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 694.5, 0, 669, 130). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 694.5, 0, 722, 130). _
        Select
    End If
' Okno 6
' Okno 6
' Okno 6

    Okno6Podzial = Worksheets("Dane").Range("F7").Value
    Okno6Strona = Worksheets("Dane").Range("G7").Value
    Okno6Uchyl = Worksheets("Dane").Range("H7").Value


' Podział dwustronny
    If Okno6Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.25, 163, 54.25, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 163, 109, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 109, 163, 0, 129 + 163). _
        Select

' Stałe prawe
    ElseIf Okno6Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.25, 163, 54.25, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 163, 54.5, 65 + 163). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.5, 63.5 + 163, 0, 127 + 163). _
        Select

' Stałe lewe
    ElseIf Okno6Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.25, 163, 54.25, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 109, 163, 54.5, 65 + 163). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.5, 65 + 163, 109, 130 + 163). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno6Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 163, 109, 65 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 109, 65 + 163, 0, 130 + 163). _
        Select
    
' Prawa
    ElseIf Okno6Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 109, 163, 0, 65 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 65 + 163, 109, 130 + 163). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno6Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 130 + 163, 55, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 55, 163, 108, 130 + 163). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno6Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 130 + 163, 27.5, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 27.5, 163, 55, 130 + 163). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno6Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 80.5, 163, 55, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 80.5, 163, 108, 130 + 163). _
        Select
    
' Uchył obustronny
    ElseIf Okno6Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 130 + 163, 27.5, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 27.5, 163, 55, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 80.5, 163, 55, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 80.5, 163, 108, 130 + 163). _
        Select
    End If

' Okno 7
' Okno 7
' Okno 7

    Okno7Podzial = Worksheets("Dane").Range("F8").Value
    Okno7Strona = Worksheets("Dane").Range("G8").Value
    Okno7Uchyl = Worksheets("Dane").Range("H8").Value


' Podział dwustronny
    If Okno7Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 163, 208.25, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 163, 263, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 263, 163, 154, 129 + 163). _
        Select

' Stałe prawe
    ElseIf Okno7Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 163, 208.25, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 163, 208.25, 65 + 163). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 63.5 + 163, 154, 127 + 163). _
        Select

' Stałe lewe
    ElseIf Okno7Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 163, 208.25, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 263, 163, 208.25, 65 + 163). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 65 + 163, 263, 130 + 163). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno7Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 163, 263, 65 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 263, 65 + 163, 154, 130 + 163). _
        Select
    
' Prawa
    ElseIf Okno7Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 263, 163, 154, 65 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 65 + 163, 263, 130 + 163). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno7Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 130 + 163, 209, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 209, 163, 262, 130 + 163). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno7Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 130 + 163, 181.5, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 181.5, 163, 209, 130 + 163). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno7Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 234.5, 163, 209, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 234.5, 163, 262, 130 + 163). _
        Select

' Uchył obustronny
    ElseIf Okno7Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 130 + 163, 181.5, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 181.5, 163, 209, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 234.5, 163, 209, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 234.5, 163, 262, 130 + 163). _
        Select
    End If
    
' Okno 8
' Okno 8
' Okno 8

    Okno8Podzial = Worksheets("Dane").Range("F9").Value
    Okno8Strona = Worksheets("Dane").Range("G9").Value
    Okno8Uchyl = Worksheets("Dane").Range("H9").Value


' Podział dwustronny
    If Okno8Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 362.25, 163, 362.25, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 163, 417, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 417, 163, 308, 129 + 163). _
        Select

' Stałe prawe
    ElseIf Okno8Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 362.25, 163, 362.25, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 163, 362.25, 65 + 163). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 362.25, 63.5 + 163, 308, 127 + 163). _
        Select

' Stałe lewe
    ElseIf Okno8Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 362.25, 163, 362.25, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 417, 163, 362.25, 65 + 163). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 362.25, 65 + 163, 417, 130 + 163). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno8Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 163, 417, 65 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 417, 65 + 163, 308, 130 + 163). _
        Select
    
' Prawa
    ElseIf Okno8Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 417, 163, 308, 65 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 65 + 163, 417, 130 + 163). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno8Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 130 + 163, 363, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 363, 163, 416, 130 + 163). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno8Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 130 + 163, 335.5, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 335.5, 163, 363, 130 + 163). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno8Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 388.5, 163, 363, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 388.5, 163, 416, 130 + 163). _
        Select

' Uchył obustronny
    ElseIf Okno8Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 308, 130 + 163, 335.5, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 335.5, 163, 363, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 388.5, 163, 363, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 388.5, 163, 416, 130 + 163). _
        Select
    End If
    
' Okno 9
' Okno 9
' Okno 9

    Okno9Podzial = Worksheets("Dane").Range("F10").Value
    Okno9Strona = Worksheets("Dane").Range("G10").Value
    Okno9Uchyl = Worksheets("Dane").Range("H10").Value


' Podział dwustronny
    If Okno9Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.25, 163, 516.25, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 163, 571, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 571, 163, 462, 129 + 163). _
        Select

' Stałe prawe
    ElseIf Okno9Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.5, 163, 516.5, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 163, 516.5, 65 + 163). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.5, 63.5 + 163, 462, 127 + 163). _
        Select

' Stałe lewe
    ElseIf Okno9Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.5, 163, 516.5, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 571, 163, 516.5, 65 + 163). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.5, 65 + 163, 571, 130 + 163). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno9Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 163, 571, 65 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 571, 65 + 163, 462, 130 + 163). _
        Select
    
' Prawa
    ElseIf Okno9Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 571, 163, 462, 65 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 65 + 163, 571, 130 + 163). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno9Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 130 + 163, 517, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 517, 163, 570, 130 + 163). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno9Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 130 + 163, 489.5, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 489.5, 163, 517, 130 + 163). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno9Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 542.5, 163, 517, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 542.5, 163, 570, 130 + 163). _
        Select

' Uchył obustronny
    ElseIf Okno9Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 130 + 163, 489.5, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 489.5, 163, 517, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 542.5, 163, 517, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 542.5, 163, 570, 130 + 163). _
        Select
    End If
    
' Okno 10
' Okno 10
' Okno 10

    Okno10Podzial = Worksheets("Dane").Range("F11").Value
    Okno10Strona = Worksheets("Dane").Range("G11").Value
    Okno10Uchyl = Worksheets("Dane").Range("H11").Value


' Podział dwustronny
    If Okno10Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.25, 163, 668.25, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 163, 723, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 723, 163, 614, 129 + 163). _
        Select

' Stałe prawe
    ElseIf Okno10Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.5, 163, 668.5, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 163, 668.5, 65 + 163). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.5, 63.5 + 163, 614, 127 + 163). _
        Select

' Stałe lewe
    ElseIf Okno10Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.5, 163, 668.5, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 723, 163, 668.5, 65 + 163). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.5, 65 + 163, 723, 130 + 163). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno10Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 163, 723, 65 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 723, 65 + 163, 614, 130 + 163). _
        Select
    
' Prawa
    ElseIf Okno10Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 723, 163, 614, 65 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 65 + 163, 723, 130 + 163). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno10Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 130 + 163, 669, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 669, 163, 722, 130 + 163). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno10Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 130 + 163, 641.5, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 641.5, 163, 669, 130 + 163). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno10Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 694.5, 163, 669, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 694.5, 163, 722, 130 + 163). _
        Select

' Uchył obustronny
    ElseIf Okno10Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 130 + 163, 641.5, 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 641.5, 163, 669, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 694.5, 163, 669, 130 + 163). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 694.5, 163, 722, 130 + 163). _
        Select
    End If
    
' Okno 11
' Okno 11
' Okno 11

    Okno11Podzial = Worksheets("Dane").Range("F12").Value
    Okno11Strona = Worksheets("Dane").Range("G12").Value
    Okno11Uchyl = Worksheets("Dane").Range("H12").Value


' Podział dwustronny
    If Okno11Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.25, 326, 54.25, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 326, 109, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 109, 326, 0, 129 + 326). _
        Select

' Stałe prawe
    ElseIf Okno11Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.25, 326, 54.25, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 326, 54.5, 65 + 326). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.5, 63.5 + 326, 0, 127 + 326). _
        Select

' Stałe lewe
    ElseIf Okno11Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.25, 326, 54.25, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 109, 326, 54.5, 65 + 326). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 54.5, 65 + 326, 109, 130 + 326). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno11Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 326, 109, 65 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 109, 65 + 326, 0, 130 + 326). _
        Select
    
' Prawa
    ElseIf Okno11Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 109, 326, 0, 65 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 65 + 326, 109, 130 + 326). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno11Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 130 + 326, 55, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 55, 326, 108, 130 + 326). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno11Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 130 + 326, 27.5, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 27.5, 326, 55, 130 + 326). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno11Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 80.5, 326, 55, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 80.5, 326, 108, 130 + 326). _
        Select
    
' Uchył obustronny
    ElseIf Okno11Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 130 + 326, 27.5, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 27.5, 326, 55, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 80.5, 326, 55, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 80.5, 326, 108, 130 + 326). _
        Select
    End If

' Okno 12
' Okno 12
' Okno 12

    Okno12Podzial = Worksheets("Dane").Range("F13").Value
    Okno12Strona = Worksheets("Dane").Range("G13").Value
    Okno12Uchyl = Worksheets("Dane").Range("H13").Value


' Podział dwustronny
    If Okno12Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 326, 208.25, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 326, 263, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 263, 326, 154, 129 + 326). _
        Select

' Stałe prawe
    ElseIf Okno12Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 326, 208.25, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 326, 208.25, 65 + 326). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 63.5 + 326, 154, 127 + 326). _
        Select

' Stałe lewe
    ElseIf Okno12Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 326, 208.25, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 263, 326, 208.25, 65 + 326). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 208.25, 65 + 326, 263, 130 + 326). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno12Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 326, 263, 65 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 263, 65 + 326, 154, 130 + 326). _
        Select
    
' Prawa
    ElseIf Okno12Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 263, 326, 154, 65 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 65 + 326, 263, 130 + 326). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno12Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 130 + 326, 209, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 209, 326, 262, 130 + 326). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno12Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 130 + 326, 181.5, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 181.5, 326, 209, 130 + 326). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno12Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 234.5, 326, 209, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 234.5, 326, 262, 130 + 326). _
        Select

' Uchył obustronny
    ElseIf Okno12Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 154, 130 + 326, 181.5, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 181.5, 326, 209, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 234.5, 326, 209, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 234.5, 326, 262, 130 + 326). _
        Select
    End If
    
' Okno 13
' Okno 13
' Okno 13

    Okno13Podzial = Worksheets("Dane").Range("F14").Value
    Okno13Strona = Worksheets("Dane").Range("G14").Value
    Okno13Uchyl = Worksheets("Dane").Range("H14").Value


' Podział dwustronny
    If Okno13Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 360.25, 326, 360.25, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 306, 326, 415, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 415, 326, 306, 129 + 326). _
        Select

' Stałe prawe
    ElseIf Okno13Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 360.25, 326, 360.25, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 306, 326, 360.25, 65 + 326). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 360.25, 63.5 + 326, 306, 127 + 326). _
        Select

' Stałe lewe
    ElseIf Okno13Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 360.25, 326, 360.25, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 415, 326, 360.25, 65 + 326). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 360.25, 65 + 326, 415, 130 + 326). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno13Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 306, 326, 415, 65 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 415, 65 + 326, 306, 130 + 326). _
        Select
    
' Prawa
    ElseIf Okno13Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 415, 326, 306, 65 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 306, 65 + 326, 415, 130 + 326). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno13Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 306, 130 + 326, 361, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 361, 326, 414, 130 + 326). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno13Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 306, 130 + 326, 333.5, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 333.5, 326, 361, 130 + 326). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno13Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 386.5, 326, 361, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 386.5, 326, 414, 130 + 326). _
        Select

' Uchył obustronny
    ElseIf Okno13Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 306, 130 + 326, 333.5, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 333.5, 326, 361, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 386.5, 326, 361, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 386.5, 326, 414, 130 + 326). _
        Select
    End If
    
' Okno 14
' Okno 14
' Okno 14

    Okno14Podzial = Worksheets("Dane").Range("F15").Value
    Okno14Strona = Worksheets("Dane").Range("G15").Value
    Okno14Uchyl = Worksheets("Dane").Range("H15").Value


' Podział dwustronny
    If Okno14Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.25, 326, 516.25, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 326, 571, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 571, 326, 462, 129 + 326). _
        Select

' Stałe prawe
    ElseIf Okno14Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.5, 326, 516.5, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 326, 516.5, 65 + 326). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.5, 63.5 + 326, 462, 127 + 326). _
        Select

' Stałe lewe
    ElseIf Okno14Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.5, 326, 516.5, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 571, 326, 516.5, 65 + 326). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 516.5, 65 + 326, 571, 130 + 326). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno14Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 326, 571, 65 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 571, 65 + 326, 462, 130 + 326). _
        Select
    
' Prawa
    ElseIf Okno14Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 571, 326, 462, 65 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 65 + 326, 571, 130 + 326). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno14Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 130 + 326, 517, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 517, 326, 570, 130 + 326). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno14Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 130 + 326, 489.5, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 489.5, 326, 517, 130 + 326). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno14Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 542.5, 326, 517, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 542.5, 326, 570, 130 + 326). _
        Select

' Uchył obustronny
    ElseIf Okno14Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 462, 130 + 326, 489.5, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 489.5, 326, 517, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 542.5, 326, 517, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 542.5, 326, 570, 130 + 326). _
        Select
    End If
    
' Okno 15
' Okno 15
' Okno 15

    Okno15Podzial = Worksheets("Dane").Range("F16").Value
    Okno15Strona = Worksheets("Dane").Range("G16").Value
    Okno15Uchyl = Worksheets("Dane").Range("H16").Value


' Podział dwustronny
    If Okno15Podzial = "Dwuskrzydłowe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.25, 326, 668.25, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 326, 723, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 723, 326, 614, 129 + 326). _
        Select

' Stałe prawe
    ElseIf Okno15Podzial = "Stałe prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.5, 326, 668.5, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 326, 668.5, 65 + 326). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.5, 63.5 + 326, 614, 127 + 326). _
        Select

' Stałe lewe
    ElseIf Okno15Podzial = "Stałe lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.5, 326, 668.5, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 723, 326, 668.5, 65 + 326). _
    Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 668.5, 65 + 326, 723, 130 + 326). _
        Select
    End If
    
' STRONA W PRZYPADKU BRAKU PODZIAŁU

' Lewa
    If Okno15Strona = "Lewe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 326, 723, 65 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 723, 65 + 326, 614, 130 + 326). _
        Select
    
' Prawa
    ElseIf Okno15Strona = "Prawe" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 723, 326, 614, 65 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 65 + 326, 723, 130 + 326). _
        Select
    End If
    
' Uchył bez podziałki
    If Okno15Uchyl = "Tak" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 130 + 326, 669, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 669, 326, 722, 130 + 326). _
        Select
    
' Uchył z podziałką z lewej
    ElseIf Okno15Uchyl = "Lewa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 130 + 326, 641.5, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 641.5, 326, 669, 130 + 326). _
        Select
    
' Uchył z podziałką z prawej
    ElseIf Okno15Uchyl = "Prawa" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 694.5, 326, 669, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 694.5, 326, 722, 130 + 326). _
        Select

' Uchył obustronny
    ElseIf Okno15Uchyl = "Obustronny" Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 614, 130 + 326, 641.5, 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 641.5, 326, 669, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 694.5, 326, 669, 130 + 326). _
        Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 694.5, 326, 722, 130 + 326). _
        Select
    End If

' Formatowanie linii

    ActiveSheet.DrawingObjects.Select
    Selection.ShapeRange.ShapeStyle = msoLineStylePreset1
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 1.5
    End With

End Sub

