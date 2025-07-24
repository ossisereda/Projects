Attribute VB_Name = "Caching"

Public Cache As Object

Public Sub InitCache()
' Varmistaa, ett‰ tietojen v‰livarasto on olemassa. Jos ei ole, luo varaston.
    If Cache Is Nothing Then
        Set Cache = CreateObject("Scripting.Dictionary")
    End If
End Sub

Public Sub ClearCache()
' Tyhjent‰‰ InitCachen
    Set Cache = Nothing
End Sub


