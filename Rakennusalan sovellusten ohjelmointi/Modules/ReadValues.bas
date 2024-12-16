Attribute VB_Name = "ReadValues"
'
' In this module, functions associated with reading the values given by the user in the UI are defined.
' These functions have the purpose of returning the values given in the UI in such a manner,
' that it is easier to handle the values in further functions (module ValuesToRows)
'


Public Function GetHousingType()
' Checks which housing types are checked by the user
' Returns the values as a string, used as a parameter in the GetRowsForHousingType function
    Dim boxNames As Variant ' Array of check boxes
    Dim boxName As Variant ' Represents a single box in boxNames
    Dim housingBoxes As String ' String to store all checked housing type boxes
    
    boxNames = Array("KT", "RT", "PARIT", "OKT")
    housingBoxes = ""
    
    ' Check each box and build the checkedBoxes string by
    ' adding the name of each checked box into the string
    For Each boxName In boxNames
        If ActiveSheet.Shapes(boxName).ControlFormat.Value = 1 Then
            housingBoxes = housingBoxes & " " & boxName
        End If
    Next boxName
    
    GetHousingType = housingBoxes
    
End Function


Public Function GetRooms()
' Checks which room boxes are checked, and returns the value as a string
' Used as a parameter in the GetRowsForRooms function
    Dim boxNames As Variant ' Array of check boxes
    Dim boxName As Variant ' Represents a single box in boxNames
    Dim roomBoxes As String ' String to store all checked room boxes
    
    boxNames = Array("1", "2", "3", "4", "5", "6")
    roomBoxes = ""
    
    ' Check each box and build the checkedBoxes string by
    ' adding the name of each checked box into the string
    For Each boxName In boxNames
        If ActiveSheet.Shapes(boxName).ControlFormat.Value = 1 Then
            roomBoxes = roomBoxes & " " & boxName
        End If
    Next boxName
    
    GetRooms = roomBoxes
    
End Function


Public Function GetPrice(ByRef priceMin As Variant, ByRef priceMax As Variant)
' Checks the price range determined in the UI
' Used in the GetRowsForPrice function
    
    priceMin = Range("A6").Value
    priceMax = Range("B6").Value
    
    ' Check if priceMin and priceMax are numeric. If not, treat them to have, respectively, values
    ' of either 1 (= no minimum price condition) or 9 999 999 (= practically no maximum price condition,
    ' at least in the case of our available data)
    If Not IsNumeric(priceMin) Or IsEmpty(priceMin) Then
        priceMin = 1
    End If
    
    If Not IsNumeric(priceMax) Or IsEmpty(priceMax) Then
        priceMax = 9999999
    End If
    
End Function


Public Function GetSquares(ByRef squareMin As Variant, ByRef squareMax As Variant)
' Checks the square meter range determined in the UI
' Used in the GetRowsForSquares function
    
    squareMin = Range("A9").Value
    squareMax = Range("B9").Value
    
    ' Check if squareMin and squareMax are numeric. If not, treat them to have, respectively, values
    ' of either 1 (= no minimum m2 condition) or 999 (= practically no maximum m2 condition,
    ' at least in the case of our available data)
    If Not IsNumeric(squareMin) Or IsEmpty(squareMin) Then
        squareMin = 1
    End If
    
    If Not IsNumeric(squareMax) Or IsEmpty(squareMax) Then
        squareMax = 999
    End If
    
End Function
