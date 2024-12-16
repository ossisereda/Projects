Attribute VB_Name = "Main_Functions"
'
' In this module, immediate functions associated with the subroutines defined in module Main are defined
'


Public Function ProcessData(ByVal source As Worksheet)
' With the help of functions defined in modules ReadValues and ValuesToRows, return a collection of row numbers
' to be printed from the source sheet. Used in the PrintChoices subroutine
    Dim housingCollection As Collection, roomCollection As Collection, _
        priceCollection As Collection, squareCollection As Collection
                                                                      
    Set housingCollection = GetRowsForHousingType(source, GetHousingType)
    Set roomCollection = GetRowsForRooms(source, GetRooms)
    GetPrice priceMin, priceMax
    Set priceCollection = GetRowsForPrice(source, priceMin, priceMax)
    GetSquares squareMin, squareMax
    Set squareCollection = GetRowsForSquares(source, squareMin, squareMax)
    
    'Return all of the row numbers that match all of the search criteria given by the user,
    'but are also stated as available in the source sheet
    Set ProcessData = GetMatchingRows(source, housingCollection, roomCollection, priceCollection, squareCollection)
    
End Function


Public Function FindMissingValue(source As Worksheet, currentRow As Long, columnIdx As Long, Optional ByRef fullHouseName As String) As Long
' Depending on the situation, finds the closest rows in the data sheet with either value or non-value, and based on these, returns a value on missing parts
' Function is necessary for informative outputs due to the way some of the data is organized in the data sheet. Used in the PrintChoices subroutine
    Dim rowNum As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim closestNonEmptyRow As Long
    Dim closestEmptyRow As Long
    
    startRow = 1
    endRow = currentRow
    fullHouseName = ""
    
    ' Search for the next non-empty row starting from the current row and moving upwards, and store it to closestNonEmptyRow
    For rowNum = currentRow To startRow Step -1
        If source.Cells(rowNum, columnIdx).Value <> "" Then
            closestNonEmptyRow = rowNum
            Exit For
        End If
    Next rowNum
    
    ' Search for the closest empty row before closestNonEmptyRow, and store it to closestEmptyRow
        For rowNum = closestNonEmptyRow - 1 To startRow Step -1
            If source.Cells(rowNum, columnIdx).Value = "" Then
                closestEmptyRow = rowNum
                
                'Special case, where the program tries to print "Kohde ja osoite" to be part of the house name
                'due to the way data is organized in the source sheet
                If closestEmptyRow < 4 Then
                    closestEmptyRow = 4
                End If
                
                Exit For
            End If
        Next rowNum
    
    ' Concatenate values between closestNonEmptyRow and closestEmptyRow to get the full house name string
        For rowNum = closestNonEmptyRow To closestEmptyRow Step -1
            fullHouseName = source.Cells(rowNum, columnIdx).Value & " " & fullHouseName
        Next rowNum
        
        ' Trim leading and trailing spaces from concatenated values
        fullHouseName = Trim(fullHouseName)
        
        ' Return the closest non-empty row
        FindMissingValue = closestNonEmptyRow
    
End Function
