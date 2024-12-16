Attribute VB_Name = "Main"
'
' In this module, subroutines associated with immediate reactions for button clicks in UI are defined
'


Private Sub DiscardButton_Click()
' Discards all of the users search choices and resets the text and check boxes to default
    
    ' Uncheck checkboxes
    ActiveSheet.Shapes("KT").OLEFormat.Object.Value = False
    ActiveSheet.Shapes("RT").OLEFormat.Object.Value = False
    ActiveSheet.Shapes("PARIT").OLEFormat.Object.Value = False
    ActiveSheet.Shapes("OKT").OLEFormat.Object.Value = False
    ActiveSheet.Shapes("1").OLEFormat.Object.Value = False
    ActiveSheet.Shapes("2").OLEFormat.Object.Value = False
    ActiveSheet.Shapes("3").OLEFormat.Object.Value = False
    ActiveSheet.Shapes("4").OLEFormat.Object.Value = False
    ActiveSheet.Shapes("5").OLEFormat.Object.Value = False
    ActiveSheet.Shapes("6").OLEFormat.Object.Value = False
    
    ' Reset text in text cells
    Range("A6,A9,B6,B9").Value = ""

End Sub


Private Sub SearchButton_Click()
' Performs a search based on the users choices. Prints all of the relevant options from the data sheet to the UI
    Dim source As Worksheet
    Set source = Workbooks("TUUMA.xlsx").Sheets("Ennakkomarkkinointi Tampere")
    
    EmptyRange
    PrintChoices source
    
End Sub


Private Sub EmptyRange()
' Clears the content and formatting of all previous search results.
' Used in the SearchButton_Click subroutine
    Dim lastRow As Long
    Dim cell As Range
    
    ' Find the last row with data in column M (price column)
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "M").End(xlUp).row
    
    ' Check if there is data in column M (in this case, meaning if there is data at all to be removed)
    ' If is, delete the row on specified range (print area in UI)
    If lastRow > 1 Then
        For Each cell In ActiveSheet.Range("I2:M" & lastRow)
            cell.ClearFormats
            cell.Value = ""
        Next cell
    End If
    
End Sub


Private Sub PrintChoices(ByVal source As Worksheet)
' Print all the relevant information from the objects in the data sheet, which match all of the search criteria and are stated as available
' Used in the SearchButton_Click subroutine
    Dim finalCollection As Collection   'Collection of all the row numbers to be printed from source sheet
    Dim rowNum As Variant
    Dim targetRow As Long
    Dim currentRow As Long
    Dim i As Integer
    Dim housingType As Long
    Dim houseName As Long
    Dim fullName As String

    Set finalCollection = ProcessData(source)
    
    ' Start copying rows from source to the current sheet (starting from row 2, column I)
    targetRow = 2
    For Each rowNum In finalCollection
        
        currentRow = rowNum
        
        ' Housing type is not stated on every row in the data sheet, therefore it cant be copied as is
        ' Find the corresponding housing type, and copy it to column J (housing type column) on the UI sheet
        housingType = FindMissingValue(source, currentRow, 4)
        ActiveSheet.Cells(targetRow, 10).Value = source.Cells(housingType, 4).Value
        
        ' Find the closest non-empty value in column A on smaller rows and concatenate all values until the next empty cell
        houseName = FindMissingValue(source, currentRow, 1, fullName)
        
        ' Copy the concatenated string value in column A of the source sheet to column I in active sheet
        ActiveSheet.Cells(targetRow, 9).Value = fullName
        
        ' Copy values from other specified columns of source to corresponding columns of ActiveSheet
        ' Column F to Column K (number of rooms)
        ActiveSheet.Cells(targetRow, 11).Value = source.Cells(currentRow, 6).Value
        ' Column G to Column L (number of squares)
        ActiveSheet.Cells(targetRow, 12).Value = source.Cells(currentRow, 7).Value
        ' Column K to Column M (price)
        ActiveSheet.Cells(targetRow, 13).Value = source.Cells(currentRow, 11).Value
        
        targetRow = targetRow + 1
    Next rowNum
    
End Sub
