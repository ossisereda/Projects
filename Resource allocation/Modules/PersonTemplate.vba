Attribute VB_Name = "PersonTemplate"

Private Sub Starting_Point(ws As Worksheet, ByRef firstEmptyRow As Long, ByRef separatorRow As Long)
' Tutkii, mist‰ template aloitetaan, ja tekee aloituksen mukaiset toimenpiteet,
' kuten esim. tyhj‰n rivin lis‰ys ja t‰h‰n reunaviivojen lis‰‰minen
    Dim usedRange As Range
    Dim row As Long
    Dim cell As Range
    Dim lastColoredRow As Long

    Set usedRange = ws.usedRange

    ' Etsi viimeinen rivi, jossa on v‰ritetty tausta (sarakkeessa B)
    lastColoredRow = 0
    For row = usedRange.Rows.Count To 1 Step -1 ' K‰y l‰pi k‰ytetyt rivit alhaalta ylˆsp‰in
        Set cell = ws.Cells(row, 2) ' Tarkistetaan solut sarakkeessa B
        If cell.Interior.Color <> RGB(255, 255, 255) Then ' Tarkistetaan, onko solu v‰rillinen (ei-valkoinen)
            lastColoredRow = row
            Exit For
        End If
    Next row

    If lastColoredRow = 0 Then ' Jos ei lˆydy v‰ritetty‰ solua, asetetaan ensimm‰iseksi tyhj‰ rivi
        firstEmptyRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    Else ' Jos lˆydet‰‰n v‰ritetty solu, template aloitetaan sen seuraavasta rivist‰
        firstEmptyRow = lastColoredRow + 1
    End If

    ' Ennen templaten luomista lis‰‰ tyhj‰ rivi, jossa on alareunaviiva
    ws.Rows(firstEmptyRow).Borders(xlEdgeBottom).LineStyle = xlContinuous

End Sub


Private Sub Template(ws As Worksheet, firstEmptyRow As Long, numRows As Long)
' Luotavan templaten visualisointi
    Dim i As Long
    Dim startRow As Long
    Dim templateRange As Range
    Dim column As Long

    ' Otsikkorivin muotoilu
    With ws.Rows(firstEmptyRow + 1)
        .RowHeight = 25
        .Interior.Color = RGB(221, 235, 247) ' Korostusv‰ri
        With .Borders(xlEdgeBottom) ' Alaviivan muotoilu
            .LineStyle = xlDouble
            .Color = RGB(255, 0, 0)
        End With
    End With

    ' Henkilˆn nimi -solu (sarake A, yhdistet‰‰n 16 rivi‰)
    With ws.Range(ws.Cells(firstEmptyRow + 1, 1), ws.Cells(firstEmptyRow + numRows, 1))
        .Merge
        .Value = "<Kirjoita henkilˆn nimi t‰h‰n>"
        .WrapText = True
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' Punainen kaksoisreunaviiva ymp‰rille
        With .Borders(xlEdgeTop)
            .LineStyle = xlDouble
            .Color = RGB(255, 0, 0)
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .Color = RGB(255, 0, 0)
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlDouble
            .Color = RGB(255, 0, 0)
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlDouble
            .Color = RGB(255, 0, 0)
        End With
    End With

    ' Projekti-otsikko
    With ws.Cells(firstEmptyRow + 1, 2)
        .Value = "Projektit"
        .WrapText = True
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Kopioi otsikkoriville viikkonumerot rivilt‰ 1
    column = 5 ' Sarake E
    Do While ws.Cells(1, column).Value <> ""
        ws.Cells(firstEmptyRow + 1, column).Value = ws.Cells(1, column).Value
        column = column + 1
    Loop

    ' Apufunktio (piilotettuun) sarakkeeseen D, johon kopioidaan tarkasteltavan henkilˆn nimi
    startRow = firstEmptyRow + 1
    For i = 1 To numRows - 1
        ws.Cells(startRow + i, 4).Formula = "=$A$" & startRow
    Next i

    ' V‰ritet‰‰n joka toinen rivi
    For i = 0 To numRows - 1
        If i Mod 2 = 1 Then
            ws.Rows(startRow + i).Interior.Color = RGB(242, 220, 219)
        End If
    
        ' Yhdistet‰‰n B- ja C-sarakkeiden solut templatessa, ja lis‰t‰‰n pystysuuntainen reunaviiva
        With ws.Range(ws.Cells(startRow + i, 2), ws.Cells(startRow + i, 3))
            .Merge
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(200, 200, 200)
        End With
    Next i

    ' M‰‰ritell‰‰n alue, jolle lis‰t‰‰n kaksoisreunaviivat (B ja C sarakkeet)
    Set templateRange = ws.Range(ws.Cells(firstEmptyRow + 2, 2), ws.Cells(firstEmptyRow + numRows, 3))
    
    ' Lis‰t‰‰n punainen kaksoisreunaviiva projektialueen ymp‰rille
    With templateRange
        ' Yl‰reuna (luodaan nyt otsikkorivin kohdalla)
        'With .Borders(xlEdgeTop)
            '.LineStyle = xlDouble
            '.Color = RGB(255, 0, 0)
        'End With
        ' Alareuna (luodaan nyt viimeiselle riville kokonaan)
        'With .Borders(xlEdgeBottom)
            '.LineStyle = xlDouble
            '.Color = RGB(255, 0, 0)
        'End With
        ' Vasemmanpuoleinen reuna
        With .Borders(xlEdgeLeft)
            .LineStyle = xlDouble
            .Color = RGB(255, 0, 0)
        End With
        ' Oikeanpuoleinen reuna
        With .Borders(xlEdgeRight)
            .LineStyle = xlDouble
            .Color = RGB(255, 0, 0)
        End With
    End With
    
    ' Punainen kaksoisreunaviiva koko templaten alareunaan
    With ws.Rows(firstEmptyRow + numRows)
        .Borders(xlEdgeBottom).LineStyle = xlDouble ' viivan tyyli
        .Borders(xlEdgeBottom).Color = RGB(255, 0, 0) ' viivan v‰ri
    End With
    
    ' Poissaolorivin muotoilu
    With ws.Cells(firstEmptyRow + 16, 2)
        .Value = "POISSAOLOT"
        .Font.Bold = True
        .Locked = True
    End With

End Sub


Sub New_Person()
    Dim ws As Worksheet
    Dim separatorRow As Long
    Dim firstEmptyRow As Long
    Dim numRows As Long

    Set ws = ActiveSheet

    ' Virheenk‰sittely sek‰ automaattisen laskennan ja n‰kym‰p‰ivityksen tilap‰inen keskeytt‰minen: suorituskykytekninen asia
    On Error GoTo Cleanup
    IsMacroRunning = True
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Etsit‰‰n aloitusrivi
    Call Starting_Point(ws, firstEmptyRow, separatorRow) ' ks. Private Subit

    ' Luodaan template
    ' numRows avulla m‰‰ritet‰‰n montako rivi‰ templaten korkeus on (tarkasteltavien projektien m‰‰r‰ +1).
    ' TƒTƒ VOI TARVITTAESSA MUUTTAA
    numRows = 16
    Call Template(ws, firstEmptyRow, numRows) ' ks. Private Subit
    
' Varmistaa, ett‰ automaattinen laskenta k‰ynnistyy makron suorittamisen j‰lkeen
Cleanup:

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    IsMacroRunning = False
    
End Sub

