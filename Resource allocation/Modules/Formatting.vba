Attribute VB_Name = "Formatting"

Private Sub Conditions(totalHours As Double, weekHours As Double, cell As Range)
' Värimuotoilun ehdot (prosentuaaliset rajat tavoitetunneista)

    Dim deviation As Double
    
    ' Erikoistapaus: nollalla jakamisen estäminen
    If weekHours = 0 Then
        cell.Interior.Color = RGB(0, 255, 0)
        Exit Sub
    End If
    
    deviation = Abs((totalHours - weekHours) / weekHours)
    
    If deviation > 0.3 Then
        cell.Interior.Color = RGB(255, 0, 0)
    ElseIf deviation > 0.15 And deviation <= 0.3 Then
        cell.Interior.Color = RGB(255, 165, 0)
    Else
        cell.Interior.Color = RGB(0, 255, 0)
    End If

End Sub


Public Sub Format(cell As Range)
    ' Hakee tuntitiedot välimuistista ja muotoilee solun ehtojen mukaan
    Dim name As String, key As String
    Dim totalHours As Double, absences As Double, weekHours As Double
    Dim values As Variant

    If Cache Is Nothing Then Exit Sub

    name = Trim(cell.Worksheet.Cells(cell.row, "D").Value)
    If name = "" Or name = "0" Then
        cell.Interior.ColorIndex = xlNone
        Exit Sub
    End If

    key = name & "|" & cell.column
    If Not Cache.Exists(key) Then Exit Sub

    values = Cache(key)
    totalHours = values(0)
    absences = values(1)
    weekHours = cell.Worksheet.Cells(cell.row, "C").Value - absences

    Call Conditions(totalHours, weekHours, cell)
End Sub



