'1. Simple Data Summarization: This code will calculate the sum of a range of cells and display the result in a message box.

Sub Calculate_Sum()
    Dim myRange As Range
    Set myRange = Selection
    MsgBox "The sum is: " & Application.WorksheetFunction.Sum(myRange)
End Sub


'2. Sorting Data: This code will sort a range of cells in ascending order.

Sub Sort_Data()
    ActiveSheet.Sort.SortFields.Add Key:=Selection, _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Selection
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


'3. Filtering Data: This code will filter a range of cells based on a selected criteria. It will only display rows where the value in column A is greater than 10.

Sub Filter_Data()
    Dim myRange As Range
    Set myRange = Selection
    myRange.AutoFilter Field:=1, Criteria1:=">10"
End Sub


'4. Pivot Tables: This code will create a Pivot Table based on a range of cells.

Sub Create_Pivot_Table()
    Dim myRange As Range
    Dim myPivotTable As PivotTable
    Dim myPivotCache As PivotCache
    
    Set myRange = Selection
    Set myPivotCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=myRange)
    Set myPivotTable = ActiveSheet.PivotTables.Add(PivotCache:=myPivotCache, TableDestination:=Range("K1"), TableName:="MyPivotTable")
    
    With myPivotTable
        .PivotFields("Column1").Orientation = xlRowField
        .AddDataField .PivotFields("Column2"), "Sum of Column2", xlSum
    End With
End Sub

''These are just a few examples of what VBA scripts can do for data analytics in Excel. There are many other possibilities depending on the specific needs of the user.