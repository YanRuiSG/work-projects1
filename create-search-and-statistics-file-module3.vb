Option Explicit



Sub create_all_statistical_summaries()
    'The main sub-procedure for the creation of all the pivot tables
    'and the formatting of each pivot table
    
    Sheets("Data").Activate
    
    'create the pivot table for the block statistics
    Call create_pivot_table("Blk", "Nationality", "Block statistics")
    
    Call UpdateProgress(43 / 100) 'Updating the Progress Indicator bar in the userform
    'apply the formatting for the pivot table for the block statistics
    Call format_block_statistics
    
    Call UpdateProgress(50 / 100)
    
    Sheets("Data").Activate
    
    'create the pivot table for the cluster statistics
    Call create_pivot_table("Cluster number", "Nationality", "Cluster statistics")
    
    Call UpdateProgress(55 / 100)
    'apply the formatting for the pivot table for the cluster statistics
    Call format_cluster_statistics
    
    Call UpdateProgress(64 / 100)
    
    Sheets("Data").Activate
    
    'create the pivot table for the company statistics
    Call create_pivot_table("Company", "Nationality", "Company statistics")
    
    Call UpdateProgress(67 / 100)
    'apply the formatting for the pivot table for the company statistics
    Call format_company_statistics
    
    Call UpdateProgress(75 / 100)
    
    Sheets("Data").Activate
    
    'create the pivot table for the room statistics
    Call create_pivot_table("Room No. Occupied", "Nationality", "Room statistics")
    
    Call UpdateProgress(80 / 100)
    'apply the formatting for the pivot table for the room statistics
    Call format_room_statistics
    
    Call UpdateProgress(95 / 100)

End Sub


Sub create_pivot_table(row As String, column As String, sheet_name As String)
    'A function to create a pivot table based on the row and column defined in the argument
    'It will directly reference the database which is hardcoded in this function
    'The pivot table will be created on a new worksheet with the name will be defined in the function argument
    
    Dim PTCache As PivotCache
    Dim PT As PivotTable
    
    ' Create the cache
    Set PTCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=Range("A2").CurrentRegion)

    ' Add a new worksheet for the pivot table
    Worksheets.Add
    
    ' Create the Pivot Table
    Set PT = ActiveSheet.PivotTables.Add( _
            PivotCache:=PTCache, _
            TableDestination:=Range("A3"))
    
    'Specify the fields
    With PT
        .PivotFields(row).Orientation = xlRowField
        .PivotFields(column).Orientation = xlColumnField
        .PivotFields("Count").Orientation = xlDataField
    End With
    
    'renaming the worksheet
    ActiveSheet.Name = sheet_name
    
End Sub


Sub format_block_statistics()

    Sheets("Block statistics").Activate
    
    'copy and paste values to remove the pivot table and keep only the values as a table
    'for easier formatting and editing
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A17").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("A1:A15").EntireRow.Delete
    Range("A2").value = "Block number"
    Range("A11").value = "Total per nationality"
    
    
    
    'find the last column
    Dim last_col As Integer
    Range("A2").Select
    Selection.End(xlToRight).Select
    last_col = ActiveCell.column
    
    Call Module1.UpdateProgress(44 / 100)
    
    'Adding the title
    Range("A1").value = "Occupancy summary per block for nationalities"
    Range(Cells(1, 1), Cells(1, last_col)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 20
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    
    'Highlighting the row labels
    Range("A2:A11").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = False
        .Font.Bold = True
        .Font.Size = 14
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    Call Module1.UpdateProgress(47 / 100)
    
    'setting the column widths
    Columns("A:A").ColumnWidth = 33
    Columns("B:B").ColumnWidth = 15
    Columns("C:C").ColumnWidth = 15
    Columns("D:D").ColumnWidth = 15
    Columns("E:E").ColumnWidth = 15
    Columns("F:F").ColumnWidth = 15
    Columns("G:G").ColumnWidth = 15
    Columns("H:H").ColumnWidth = 15
    Columns("I:I").ColumnWidth = 15
    Columns("J:J").ColumnWidth = 15
    Columns("K:K").ColumnWidth = 15
    
    'Aligning the values as center postion in each cell
    'Making the column names and some values bold
    Range(Cells(2, 2), Cells(11, last_col)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 14
    End With
    Range(Cells(11, 2), Cells(11, last_col)).Font.Bold = True
    Range(Cells(2, 2), Cells(2, last_col)).Font.Bold = True
    Range(Cells(3, last_col), Cells(10, last_col)).Font.Bold = True
    
    
    ' Adding the data bars for the data field summary
    Dim i As Integer
    For i = 2 To last_col
        Range(Cells(3, i), Cells(11, i)).Select
        Selection.FormatConditions.AddDatabar
        Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
        End With
        With Selection.FormatConditions(1).BarColor
            .Color = 13012579
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
        Selection.FormatConditions(1).Direction = xlContext
        Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
        Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
            xlDataBarColor
        With Selection.FormatConditions(1).BarBorder.Color
            .Color = 13012579
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
        With Selection.FormatConditions(1).AxisColor
            .Color = 0
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.Color
            .Color = 255
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.BorderColor
            .Color = 255
            .TintAndShade = 0
        End With
    Next i
    
    Call Module1.UpdateProgress(49 / 100)
    
    'adding the borders for the table
    Range(Cells(1, 1), Cells(11, last_col)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
 
End Sub


Sub format_cluster_statistics()

    Sheets("Cluster statistics").Activate
    
    'copy and paste values to remove the pivot table and keep only the values as a table
    'for easier formatting and editing
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A17").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("A1:A15").EntireRow.Delete
    
    
    
    'find the last column
    Dim last_col As Integer
    Range("A2").Select
    Selection.End(xlToRight).Select
    last_col = ActiveCell.column
    
    Range("A2").value = "Cluster number"
    Range("A6").value = "Total per nationality"
    
    'Adding the title
    Range("A1").value = "Occupancy summary per cluster for nationalities"
    Range(Cells(1, 1), Cells(1, last_col)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 20
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Call UpdateProgress(56 / 100)
    
    'Highlighting the row labels
    Range("A2:A6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = False
        .Font.Bold = True
        .Font.Size = 14
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    
    'setting the column widths
    Columns("A:A").ColumnWidth = 33
    Columns("B:B").ColumnWidth = 15
    Columns("C:C").ColumnWidth = 15
    Columns("D:D").ColumnWidth = 15
    Columns("E:E").ColumnWidth = 15
    Columns("F:F").ColumnWidth = 15
    Columns("G:G").ColumnWidth = 15
    Columns("H:H").ColumnWidth = 15
    Columns("I:I").ColumnWidth = 15
    Columns("J:J").ColumnWidth = 15
    Columns("K:K").ColumnWidth = 15
    
    Call UpdateProgress(59 / 100)
    
    'apply center alignment for all cell values
    'Making the column names and some values bold
    Range(Cells(2, 1), Cells(6, last_col)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 14
    End With
    Range("B6:K6").Font.Bold = True
    Range("B2:K2").Font.Bold = True
    Range("K3:K6").Font.Bold = True
    
    
    ' Adding the data bars for the data field summary
    Dim i As Integer
    For i = 2 To last_col
        Range(Cells(3, i), Cells(6, i)).Select
        Selection.FormatConditions.AddDatabar
        Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
        End With
        With Selection.FormatConditions(1).BarColor
            .Color = 13012579
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
        Selection.FormatConditions(1).Direction = xlContext
        Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
        Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
            xlDataBarColor
        With Selection.FormatConditions(1).BarBorder.Color
            .Color = 13012579
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
        With Selection.FormatConditions(1).AxisColor
            .Color = 0
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.Color
            .Color = 255
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.BorderColor
            .Color = 255
            .TintAndShade = 0
        End With
    Next i
    
    Call UpdateProgress(61 / 100)
    
    'adding the borders for the table
    Range("A1").CurrentRegion.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   


End Sub


Sub format_company_statistics()

    Sheets("Company statistics").Activate
    
    
    'copy and paste values to remove the pivot table and keep only the values as a table
    'for easier formatting and editing
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("M4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("A1:L1").EntireColumn.Delete
    
    
    
    'Delete the last row of the dataset
    Dim last_row As Integer
    Range("A4").Select
    Selection.End(xlDown).Select
    ActiveCell.EntireRow.Delete
    
    
    'Find the new last row of the dataset
    Range("A4").Select
    Selection.End(xlDown).Select
    last_row = ActiveCell.row
    
    'find the last column of the dataset
    Dim last_col As Integer
    Range("A4").Select
    Selection.End(xlToRight).Select
    last_col = ActiveCell.column
    
    Range("A4").value = "Name of company"
    
    Call Module1.UpdateProgress(68 / 100)
    
    'Adding the title
    Range("A3").value = "Breakdown of nationalities per company"
    Range(Cells(3, 1), Cells(3, last_col)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 20
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    'Highlighting the row labels
    Range("A4").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = False
        .Font.Bold = True
        .Font.Size = 11
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    Call Module1.UpdateProgress(69 / 100)
    
    'setting the column widths
    Columns("A:A").ColumnWidth = 40
    Columns("B:B").ColumnWidth = 15
    Columns("C:C").ColumnWidth = 15
    Columns("D:D").ColumnWidth = 15
    Columns("E:E").ColumnWidth = 15
    Columns("F:F").ColumnWidth = 15
    Columns("G:G").ColumnWidth = 15
    Columns("H:H").ColumnWidth = 15
    Columns("I:I").ColumnWidth = 15
    Columns("J:J").ColumnWidth = 15
    Columns("K:K").ColumnWidth = 15
    
    
    'Center alignmemt for all values
    'Making the column names and some values bold
    Range("A4").CurrentRegion.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 11
    End With
    Range(Cells(1, 1), Cells(last_row, 1)).Font.Bold = True
    Range(Cells(4, 2), Cells(4, last_col)).Font.Bold = True
    Range(Cells(4, last_col), Cells(last_row, last_col)).Font.Bold = True
    
    
    ' Adding the data bars for the data field summary
    Dim i As Integer
    For i = 5 To last_row
        Range(Cells(i, 2), Cells(i, last_col)).Select
        Selection.FormatConditions.AddDatabar
        Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
        End With
        With Selection.FormatConditions(1).BarColor
            .Color = 13012579
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
        Selection.FormatConditions(1).Direction = xlContext
        Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
        Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
            xlDataBarColor
        With Selection.FormatConditions(1).BarBorder.Color
            .Color = 13012579
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
        With Selection.FormatConditions(1).AxisColor
            .Color = 0
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.Color
            .Color = 255
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.BorderColor
            .Color = 255
            .TintAndShade = 0
        End With
    Next i
    
    Call Module1.UpdateProgress(70 / 100)
    
    'adding the borders for the table
    Range("A4").CurrentRegion.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    'apply thick borders around every row
    Dim row_num As Integer
    For row_num = 5 To last_row
        Range(Cells(row_num, 1), Cells(row_num, last_col)).Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Next row_num
    
    Call Module1.UpdateProgress(72 / 100)
    
    
    Range("A3").Font.Size = 20
    Range("A4").Select
    Selection.AutoFilter
    
    'set the filters for each column
    'sort by the company with the most of number of workers to be at the top
    ActiveWorkbook.Worksheets("Company statistics").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Company statistics").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range(Cells(4, last_col), Cells(last_row, last_col)), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Company statistics").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub


Sub format_room_statistics()

    Sheets("Room statistics").Activate
    

    'save the activebook name for reference later
    Dim name_ref As String
    name_ref = ActiveWorkbook.Name
    
    
    'Define the nationality count and total count for dynamic populating of values later
    Dim nationality_count As Integer, total_count As Integer
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    total_count = Selection.Count
    nationality_count = Selection.Count - 1
    
    

    'copy and paste values to remove the pivot table and keep only the values as a table
    'for easier formatting and editing
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("AG4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("A1:L1").EntireColumn.Delete
    
    Call Module1.UpdateProgress(81 / 100)
    
    'retrieve the list of all room numbers for reference
    Windows("Create DMS search and statistics file (version 1.0).xlsm").Activate
    Worksheets("Workings").Activate
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Workbooks(name_ref).Worksheets("Room statistics").Activate
    Range("A4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'shifting and renaming columns
    Sheets("Data").Activate
    Columns("M:M").Copy
    Sheets("Room statistics").Activate
    Columns("AI:AI").Select
    Selection.Insert Shift:=xlToRight
    Sheets("Data").Activate
    Columns("I:I").Copy
    Sheets("Room statistics").Activate
    Columns("AJ:AJ").Select
    Selection.Insert Shift:=xlToRight

    Range("E4").value = "Tenant/sub-tenant (if room physically occupied)"
    Range("F4").value = "Total number of workers"
    Range("G4").value = "Total space left"
    Range("V4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy Range("H4")
    
    Range("A4").Select
    Selection.End(xlToRight).Select
    ActiveCell.Clear
    
    Call Module1.UpdateProgress(83 / 100)
    
    'Define the room number-nationality table range for look up reference later
    Dim table1 As Range
    Range("U4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set table1 = Selection
    
    'Define the room number-company table range for look up reference later
    Dim table2 As Range
    Range("AI1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set table2 = Selection
    
    'Populating the values for the room statistics using a dynamic method which will allow for a variable number of different nationalities
    'that can exist in the dormitory
    'use the worksheet function vlookup to populate each cell with the right values
    'loop through across each column for each row
    Dim i As Integer, j As Integer
    For i = 5 To 804
        On Error Resume Next
        Range(Cells(i, 5), Cells(i, 5)).value = WorksheetFunction.VLookup(Range(Cells(i, 1), Cells(i, 1)), table2, 2, 0)
        Range(Cells(i, 6), Cells(i, 6)).value = WorksheetFunction.VLookup(Range(Cells(i, 1), Cells(i, 1)), table1, total_count, 0)
        Range(Cells(i, 7), Cells(i, 7)).value = 16 - Range(Cells(i, 6), Cells(i, 6)).value
        
    Next i

    For j = 2 To nationality_count
        For i = 5 To 804
            Range(Cells(i, j + 6), Cells(i, j + 6)).value = WorksheetFunction.VLookup(Range(Cells(i, 1), Cells(i, 1)), table1, j, 0)
    
        Next i
    Next j




    ' Delete the excess columns
    Range("U1:AJ1").EntireColumn.Delete
    
    
    'setting the column widths to make column values visible
    Columns("A:A").ColumnWidth = 15
    Columns("B:B").ColumnWidth = 8
    Columns("C:C").ColumnWidth = 8
    Columns("D:D").ColumnWidth = 8
    Columns("E:E").ColumnWidth = 60
    Columns("F:F").ColumnWidth = 20
    Columns("G:G").ColumnWidth = 20
    Columns("H:H").ColumnWidth = 10
    Columns("I:I").ColumnWidth = 10
    Columns("J:J").ColumnWidth = 10
    Columns("K:K").ColumnWidth = 10
    Columns("L:L").ColumnWidth = 10
    Columns("M:M").ColumnWidth = 10
    Columns("N:N").ColumnWidth = 10
    Columns("O:O").ColumnWidth = 10
    
    Call Module1.UpdateProgress(85 / 100)
    
    'Find the new last row of the dataset
    Dim last_row As Integer
    Range("A4").Select
    Selection.End(xlDown).Select
    last_row = ActiveCell.row
    
    
    
    'find the last column of the dataset
    Dim last_col As Integer
    Range("A4").Select
    Selection.End(xlToRight).Select
    last_col = ActiveCell.column
    
    
    
    'Center alignmemt for all values
    'Making the column names and some values bold
    Range("A4").CurrentRegion.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 11
    End With
    Range(Cells(4, 1), Cells(last_row, 1)).Font.Bold = True
    Range(Cells(4, 2), Cells(4, last_col)).Font.Bold = True
    Range(Cells(4, last_col), Cells(last_row, last_col)).Font.Bold = True
    
    Call Module1.UpdateProgress(87 / 100)
    
    ' Adding the data bars for the total workers in each in room
    For i = 5 To last_row
        Range(Cells(i, 6), Cells(i, 6)).Select
        Selection.FormatConditions.AddDatabar
        Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
        End With
        With Selection.FormatConditions(1).BarColor
            .Color = 13012579
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
        Selection.FormatConditions(1).Direction = xlContext
        Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
        Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
            xlDataBarColor
        With Selection.FormatConditions(1).BarBorder.Color
            .Color = 13012579
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
        With Selection.FormatConditions(1).AxisColor
            .Color = 0
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.Color
            .Color = 255
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.BorderColor
            .Color = 255
            .TintAndShade = 0
        End With
    Next i
    
    'Adding the conditional formatting for the column for total workers in each room
    Range("F5:F804").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=LEN(TRIM(F5))=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399945066682943
    End With
    
    Call Module1.UpdateProgress(90 / 100)
    
    'Adding the conditional formatting for the column for space left in each room
    Range("G5:G804").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    
    ' Adding the data bars for the data field summary
    Range(Cells(5, 8), Cells(last_row, last_col)).Select
    Selection.FormatConditions.AddDatabar
    Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With
    With Selection.FormatConditions(1).BarColor
        .Color = 13012579
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
    Selection.FormatConditions(1).Direction = xlContext
    Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
    Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
        xlDataBarColor
    With Selection.FormatConditions(1).BarBorder.Color
        .Color = 13012579
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    With Selection.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.Color
        .Color = 255
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.BorderColor
        .Color = 255
        .TintAndShade = 0
    End With
    'Adding conditional formatting for the data field summary
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=LEN(TRIM(H5))=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399945066682943
    End With
    
    Call Module1.UpdateProgress(93 / 100)
    
    'apply thick borders around every row
    Dim row_num As Integer
    For row_num = 4 To last_row
        Range(Cells(row_num, 1), Cells(row_num, last_col)).Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Next row_num
    
    
    
    'Adding the title and background color
    Range("A3").value = "Occupancy and breakdown of nationalities per room"
    Range(Cells(3, 1), Cells(3, last_col)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 20
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    'Adding the filter for each column
    Range("A4").Select
    Selection.AutoFilter
    
End Sub


