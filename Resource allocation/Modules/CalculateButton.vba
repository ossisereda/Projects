Attribute VB_Name = "CalculateButton"

Public Sub CalculateButton()
' K‰ynnist‰‰ laskennan nappia painettaessa
    Dim startTime As Double, endTime As Double, totalTime As Double
    Dim cell As Range
    Dim formulaCells As Range

    startTime = Timer ' Debug

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Alustetaan cache; ks. Caching-moduuli
    Call ClearCache
    Call InitCache

    ' Etsit‰‰n kaikki solut, joissa on kaavoja
    On Error Resume Next
    Set formulaCells = ActiveSheet.usedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    ' Lasketaan ja muotoillaan solut, joissa esiintyy "AggregateData" (eli myˆs AggregateDataFull)
    If Not formulaCells Is Nothing Then
        For Each cell In formulaCells
            If InStr(1, cell.Formula, "AggregateData", vbTextCompare) > 0 Then
                cell.Formula = cell.Formula ' Pakottaa UDF:n uudelleenlaskennan. P‰ivitt‰‰ v‰limuistin Calculation-moduulin mukaisesti
                Call Format(cell) ' Muotoilu cachen tietojen pohjalta
            End If
        Next cell
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Debug
    endTime = Timer
    totalTime = endTime - startTime
    Debug.Print "Laskenta-aika: " & Round(totalTime, 2) & " s."
    
End Sub

