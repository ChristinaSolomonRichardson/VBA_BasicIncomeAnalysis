Attribute VB_Name = "Module1"
'Sub to delete rows matching income - interest
Sub DeleteRows()

    Dim xLastRow As Integer
    
    xLastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    'For look for finding rows to delete
    For x = 1 To xLastRow
        If Range("A" & x).Value = "Income - Interest" Then
            Range("A" & x).EntireRow.Delete
            x = x + 1
        End If
    Next
    
End Sub

'Function to allocate dates to quarters
Function GetQuarter(dateValue As Date) As Integer
    
    'Select case to find month value and assign quarter value
    Select Case Month(dateValue)
        Case 1 To 3
            GetQuarter = 1
        Case 4 To 6
            GetQuarter = 2
        Case 7 To 9
            GetQuarter = 3
        Case Else
            GetQuarter = 4
    End Select
    
End Function
'Sub to add quarter and year columns
Sub GetYearQtr()
    Dim xLastRow As Integer
    Dim xYear As Integer
    Dim xQtr As Integer
    
    xLastRow = Range("A" & Rows.Count).End(xlUp).Row

    For x = 2 To xLastRow
        xYear = Year(Range("B" & x).Value)
        xQtr = GetQuarter(Range("B" & x).Value)
        
        Range("D" & x) = xYear
        Range("E" & x) = xQtr
    Next x

End Sub
'Sub to create pivot table with new data with Account Name,Year, Amount Fields.
Sub CreatePivotTable()
    
    ' Define variables
    Dim pvtCache As PivotCache
    Dim pvtTable As PivotTable
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Set worksheet variable
    Set ws = ActiveSheet
    
    ' Determine last row of data
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Define pivot cache range
    Set pvtCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=ws.Range("A1:E" & lastRow))
        
    ' Create pivot table
    Set pvtTable = pvtCache.CreatePivotTable( _
        TableDestination:=ws.Range("G2"))
        
    ' Add fields to pivot table
    With pvtTable
        With .PivotFields("Account Name")
            .Orientation = xlRowField
            .Position = 1
        End With
        With .PivotFields("Year")
            .Orientation = xlColumnField
            .Position = 1
        End With
        With .PivotFields("Amount")
            .Orientation = xlDataField
            .Function = xlSum
            .Position = 1
            .NumberFormat = "$#,##0"
        End With
    End With
    
End Sub

'Sub to execute all macros in sequence.
Sub ConsolidatedMacro()
    DeleteRows
    GetYearQtr
    CreatePivotTable
End Sub

