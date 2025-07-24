Attribute VB_Name = "Calculation"

Function AggregateData(cell As Range) As String
' Viikko- ja henkilökohtaisen tuntiresursoinnin laskenta
    Dim ws As Worksheet
    Dim hours As Double, totalHours As Double, weekHours As Double, absences As Double
    Dim name As String, key As String
    Dim col As Long, lastRow As Long, i As Long

    Call InitCache

    Set ws = cell.Worksheet
    name = Trim(ws.Cells(cell.row, "D").Value)
    col = cell.column

    If name = "" Or name = "0" Then
        AggregateData = ""
        Exit Function
    End If

    totalHours = 0
    absences = 0

    lastRow = ws.usedRange.row + ws.usedRange.Rows.Count - 1
    For i = 35 To lastRow
        If Trim(ws.Cells(i, 4).Value) = name Then
            If IsNumeric(ws.Cells(i, col).Value) Then
                hours = ws.Cells(i, col).Value
                
                If ws.Cells(i, 2).Value = "POISSAOLOT" Then
                    absences = absences + hours
                Else
                    totalHours = totalHours + hours
                End If
                
            Else
                Debug.Print "Non-numeric value skipped at row " & i & ": """ & ws.Cells(i, col).Value & """"
            End If
        End If
    Next i

    weekHours = ws.Cells(cell.row, "C").Value - absences ' henkilölle merkatut viikkotunnit, joista vähennetty tarkasteltavan viikon merkatut poissaolotunnit

    ' tietojen tallennus Cacheen
    key = name & "|" & col
    Cache(key) = Array(totalHours, absences)

    AggregateData = totalHours & " / " & weekHours
    
End Function


Function AggregateDataFull(cell As Range) As String
' Yhteenvetovälilehdelle optimoitu laskenta; huomioi kaikki laskentavälilehdet
    Dim ws As Worksheet
    Dim hours As Double, totalHours As Double, weekHours As Double, absences As Double
    Dim name As String, key As String
    Dim col As Long, lastRow As Long, i As Long

    Call InitCache

    name = Trim(cell.Worksheet.Cells(cell.row, "D").Value)
    col = cell.column

    If name = "" Or name = "0" Then
        AggregateDataFull = ""
        Exit Function
    End If

    totalHours = 0
    absences = 0

    For Each ws In ThisWorkbook.Sheets
        If ws.name <> "Back-end" And ws.name <> "YHTEENVETO" And ws.Visible Then
            lastRow = ws.usedRange.row + ws.usedRange.Rows.Count - 1
            For i = 35 To lastRow
                If Trim(ws.Cells(i, 4).Value) = name Then
                    If IsNumeric(ws.Cells(i, col).Value) Then
                        hours = ws.Cells(i, col).Value
                        
                        If ws.Cells(i, 2).Value = "POISSAOLOT" Then
                            absences = absences + hours
                        Else
                            totalHours = totalHours + hours
                        End If
                        
                    Else
                        Debug.Print "Non-numeric value skipped at row " & i & ": """ & ws.Cells(i, col).Value & """"
                    End If
                End If
            Next i
        End If
    Next ws

    weekHours = cell.Worksheet.Cells(cell.row, "C").Value - absences

    key = name & "|" & col
    Cache(key) = Array(totalHours, absences, name)

    AggregateDataFull = totalHours & " / " & weekHours
    
End Function



