Option Explicit



Sub create_data_search()
'
' create_data_search Macro
'

    'remove screen updating to speed up code and hide the changes in the excel sheet from showing while happening
    Application.ScreenUpdating = False
    
    'insert a new worksheet for the data search and renaming the worksheet
    Application.Goto Worksheets("Data").Range("A1")
    Sheets.Add before:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Search DMS"
    Sheets("Data").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Sheets("Search DMS").Select
    Range("A2").Select
    ActiveSheet.Paste
    
    'delete the unnecessary column
    Range("F1").EntireColumn.Delete
    
    'Updating the Progress Indicator bar in the userform
    Call UpdateProgress(32 / 100)
    
    'Add the title and merge the cells in the first row for the title
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Current DMS search"
    Range("A1:Q1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
        .Font.Bold = True
        .Font.Size = 20
    End With
    'Set the background colour for the search page title to yellow
    Range("A1:Q1").Interior.Color = vbYellow
    
    'adjusting each column width to make the values visible
    Columns("A:A").ColumnWidth = 47
    Columns("B:B").ColumnWidth = 28
    Columns("C:C").ColumnWidth = 20
    Columns("D:D").ColumnWidth = 20
    Columns("E:E").ColumnWidth = 22
    Columns("F:F").ColumnWidth = 22
    Columns("G:G").ColumnWidth = 22
    
    
    Columns("H:H").ColumnWidth = 47
    Columns("I:I").ColumnWidth = 8
    
    Columns("J:J").ColumnWidth = 8
    Columns("K:K").ColumnWidth = 8
    Columns("L:L").ColumnWidth = 20
    Columns("M:M").ColumnWidth = 20
    Columns("N:N").ColumnWidth = 20
    Columns("O:O").ColumnWidth = 40
    Columns("P:P").ColumnWidth = 15
    Columns("Q:Q").ColumnWidth = 15
    
    
    'Updating the Progress Indicator bar in the userform
    Call UpdateProgress(33 / 100)
    
    
    Range("B2").value = "ENTER FIN No."
    
    
    'Insert the formula for vlookup for the respective columns
    Range("A2").value = "Name"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],Data!R1C1:R8000C15,2,0)"
    Range("A3").Select
    Selection.AutoFill Destination:=Range("A3:A1000"), Type:=xlFillDefault
    
    
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,3,0)"
    
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,4,0)"
   
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,5,0)"
    
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,7,0)"
    
    Range("G2").value = "WP expiry date" 'Change the column names if necessary
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,8,0)"
    
    Range("H2").value = "Company"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,9,0)"
    
    Range("I2").value = "Blk"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,14,0)"
    
    Range("J2").value = "Level"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,15,0)"
    
    Range("K2").value = "Unit"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,16,0)"
    
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,13,0)"
    
    Range("M2").value = "Company POC"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,10,0)"
    
    Range("N2").value = "POC contact no."
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,11,0)"
    
    Range("O2").value = "POC email address"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,12,0)"
    
    Range("P3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,17,0)"
    
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC2,Data!R1C1:R8000C18,18,0)"
    
    
    
    
    'Updating the Progress Indicator bar in the userform
    Call UpdateProgress(35 / 100)
    
    
    
    
    ' fill the remaining rows with the formulas
    Range("C3:Q3").Select
    Selection.AutoFill Destination:=Range("C3:Q1000"), Type:=xlFillDefault
    Range("A2:Q1000").Select
    'Apply borders
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
    
    
    'Updating the Progress Indicator bar in the userform
    Call UpdateProgress(37 / 100)
    
    
    
    
    'set the date format for the columns with date values
    Columns("E:E").NumberFormat = "m/d/yyyy"
    Columns("G:G").NumberFormat = "m/d/yyyy"
    
    'Updating the Progress Indicator bar in the userform
    Call UpdateProgress(38 / 100)
    
    
    Sheets("Search DMS").Range("B3").Select
    'Sheets("Search DMS").Move before:=Sheets(1)
'
End Sub

