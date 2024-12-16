Attribute VB_Name = "ValuesToRows"
'
' In this module, functions associated with finding the relevant rows from the data sheet according to the users choices are defined.
' The functions return the values of rows based on each separate search criteria, so it is later easier to compare
' which of the rows match with the other criteria as well (defined in the last function, GetMatchingRows)
' The collection of rows with all criteria intact are then used in printing the relevant choices in the main program
'


Public Function GetRowsForSquares(ByVal source As Worksheet, ByVal squareMin As Double, ByVal squareMax As Double) As Collection
    Dim lastRow As Long
    Dim row As Long
    Dim matchFound As Boolean
    Dim matchedRows As New Collection
    
    ' Find the last row with data in column G (m2 column in the source sheet)
    lastRow = source.Cells(source.Rows.Count, "G").End(xlUp).row

    ' Loop through each row in the source sheet, data starts from row 4. Add the row number to the matchedRows collection,
    ' if the value of the cell in column G is not empty and is between squareMin and Squaremax
    For row = 4 To lastRow
        ' Reset matchFound flag for each row
        matchFound = False
        
        If Not IsEmpty(source.Cells(row, "G").Value) Then
            If source.Cells(row, "G").Value >= squareMin And source.Cells(row, "G").Value <= squareMax Then
                matchedRows.Add row
            End If
        End If
    Next row

    ' Return the collection of matching rows
    Set GetRowsForSquares = matchedRows
    
End Function


Public Function GetRowsForPrice(ByVal source As Worksheet, ByVal priceMin As Double, ByVal priceMax As Double) As Collection
    Dim lastRow As Long
    Dim row As Long
    Dim matchFound As Boolean
    Dim matchedRows As New Collection
    
    ' Find the last row with data in column K (price column in the source sheet)
    lastRow = source.Cells(source.Rows.Count, "K").End(xlUp).row

    ' Loop through each row in the source sheet, data starts from row 4. Add the row number to the matchedRows collection,
    ' if the value of the cell in column K is not empty and is between priceMin and pricemax
    For row = 4 To lastRow
        ' Reset matchFound flag for each row
        matchFound = False
        
        If Not IsEmpty(source.Cells(row, "K").Value) Then
            If source.Cells(row, "K").Value >= priceMin And source.Cells(row, "K").Value <= priceMax Then
                matchedRows.Add row
            End If
        End If
    Next row

    ' Return the collection of matching rows
    Set GetRowsForPrice = matchedRows
    
End Function


Public Function GetRowsForRooms(ByVal source As Worksheet, ByVal rooms As String) As Collection
    Dim lastRow As Long
    Dim roomDigits As Variant
    Dim row As Long
    Dim matchFound As Boolean
    Dim matchedRows As New Collection
    
    ' Find the last row with data in column F (number of rooms column in the source sheet)
    lastRow = source.Cells(source.Rows.Count, "F").End(xlUp).row

    ' If GetRooms return value is empty, regard as number of rooms being indifferent (add every row)
    If rooms = "" Then
        For row = 4 To lastRow
            matchedRows.Add row
        Next row
        
    Else
        ' Split the room digits based on space character
        roomDigits = Split(rooms, " ")

        ' Loop through each row in the source sheet, data starts from row 4
        For row = 4 To lastRow
            ' Reset matchFound flag for each row
            matchFound = False
            
            If Not IsEmpty(source.Cells(row, "F").Value) Then
                ' Loop through each digit extracted from GetRooms, after which loop through each digit extracted from the content of column F
                For Each roomDigit In roomDigits
                    For i = 1 To Len(source.Cells(row, "F").Value)
                        ' Check if the digit from GetRooms is found in the content of column F.
                        ' If a match is found, set matchFound flag to true and exit the loop
                        If Mid(source.Cells(row, "F").Value, i, 1) = roomDigit Then
                            matchFound = True
                            Exit For
                        End If
                    Next i
                    If matchFound Then Exit For
                Next roomDigit
                
                If matchFound Then
                    matchedRows.Add row
                End If
            End If
        Next row
    End If

    ' Return the collection of matching rows
    Set GetRowsForRooms = matchedRows
    
End Function


Public Function GetRowsForHousingType(ByVal source As Worksheet, ByVal housingType As String) As Collection
    Dim lastRow As Long
    Dim housingTypes As Variant
    Dim row As Long
    Dim nextRow As Long
    Dim matchFound As Boolean
    Dim matchedRows As New Collection
    
    ' Find the last row with data in column D (housing type column in the source sheet)
    lastRow = source.Cells(source.Rows.Count, "D").End(xlUp).row

    ' If GetHousingType return value is empty, regard as housing type being indifferent (add every row)
    If housingType = "" Then
        For row = 4 To lastRow
            matchedRows.Add row
        Next row
    Else
        ' Split the housing types based on space character
        housingTypes = Split(housingType, " ")

            ' Loop through each row in the source sheet, data starts from row 4
            For row = 4 To lastRow
                ' Reset matchFound flag for each row
                matchFound = False
                
                ' Check if the cell in column D is not empty
                If Not IsEmpty(source.Cells(row, "D").Value) Then
                    ' Loop through each housing type extracted from GetHousingType, and check if the cell in column D contains the housing type
                    ' If an exact match is found, set matchFound flag to true and exit the loop
                    For Each housingTypeItem In housingTypes
                        If source.Cells(row, "D").Value = housingTypeItem Then
                            matchFound = True
                            Exit For
                        End If
                    Next housingTypeItem
                    
                    If matchFound Then
                        matchedRows.Add row
                        ' Add details of subsequent rows until there is a row without any value in column L
                        nextRow = row + 1
                        Do Until IsEmpty(source.Cells(nextRow, "L").Value)
                            matchedRows.Add nextRow
                            nextRow = nextRow + 1
                        Loop
                    End If
                    
                End If
            Next row
    End If
    
    ' Return the collection of matching rows
    Set GetRowsForHousingType = matchedRows
    
End Function


Public Function GetMatchingRows(ByVal source As Worksheet, ByVal collection1 As Collection, ByVal collection2 As Collection, ByVal collection3 As Collection, ByVal collection4 As Collection) As Collection
    Dim resultCollection As New Collection
    Dim item As Variant
    Dim i As Integer
    Dim dict1 As Object, dict2 As Object, dict3 As Object, dict4 As Object  'Dictionaries for each collection
        
    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    Set dict3 = CreateObject("Scripting.Dictionary")
    Set dict4 = CreateObject("Scripting.Dictionary")
    
    ' Populate dictionaries with elements from each collection
    For Each item In collection1
        dict1(item) = True
    Next item
    For Each item In collection2
        dict2(item) = True
    Next item
    For Each item In collection3
        dict3(item) = True
    Next item
    For Each item In collection4
        dict4(item) = True
    Next item
    
    ' Check elements of the first dictionary against other dictionaries to get the rows in the data sheet
    ' which match with all of the search criteria (stored in resultCollection)
    For Each item In dict1.keys
        If dict2.exists(item) And dict3.exists(item) And dict4.exists(item) Then
            resultCollection.Add item
        End If
    Next item
    
    'Do one more iteration to the rows, where all non-available options (stated in column I of source sheet) are deleted
    For i = resultCollection.Count To 1 Step -1 ' Loop through the collection backwards to avoid issues when deleting items
        If source.Cells(resultCollection(i), "I").Value = 0 Then
            resultCollection.Remove i
        End If
    Next i
    
    ' Print a message, if there are no items in the resultCollection
    If resultCollection.Count = 0 Then
        MsgBox "There are no results that match with current search criteria."
    End If
    
    Set GetMatchingRows = resultCollection
    
End Function


