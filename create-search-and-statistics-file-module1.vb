Option Explicit

'object declared at public level for the userform progress indicator bar
'this will allow all modules to have access to it
Public ProgressIndicator As Object


'workbook object declared at public level for all procedures and functions in all modules to have access to it
'to be set as the file to be created for data search and summary statistics
Public dms_report As Workbook


Sub main_procedure()
    'This sub-procedure that will be the main procedure to call the other sub-procedures
    'It is the main procedure to run for this VBA macro
    
    'remove screen updating to speed up code and hide the changes in the excel sheet from showing while happening
    Application.ScreenUpdating = False
    
    'create a copy for the progress bar object
    Set ProgressIndicator = New UserForm1
    
    'show progress bar in modeless state
    ProgressIndicator.Show vbModeless
    'check to ensure that the active selection is a worksheet
    If TypeName(ActiveSheet) <> "Worksheet" Then
        Unload ProgressIndicator
        MsgBox "Please ensure that a worksheet is being selected.", vbCritical + vbOKOnly, "Error encountered"
        Exit Sub
    End If
        
    'Updating the Progress Indicator bar in the userform to start from
    Call UpdateProgress(0 / 100)
    
    'call the procedure to create the file
    Call create_file
    
    'Updating the Progress Indicator bar in the userform
    Call UpdateProgress(10 / 100)
    
    'call the procedure to tidy and modify the database
    Call clean_and_tidy_dataset
    
    'Updating the Progress Indicator bar in the userform
    Call UpdateProgress(30 / 100)
    
    Call Module2.create_data_search
    
    'Updating the Progress Indicator bar in the userform
    Call UpdateProgress(40 / 100)
    
    Call Module3.create_all_statistical_summaries
    
    'Updating the Progress Indicator bar in the userform
    Call UpdateProgress(100 / 100)

    Unload ProgressIndicator
    
    Set ProgressIndicator = Nothing
    

End Sub


Sub UpdateProgress(pct)
    'sub-procedure for updating the progress indicator bar to reflect the current progress
    
    With ProgressIndicator
        .FrameProgress.Caption = Format(pct, "0%")
        .LabelProgress.Width = pct * (.FrameProgress.Width)
    End With
    
    DoEvents ' statement responsible for the form updating

End Sub


Sub create_file()

    Dim dms_file_path As Variant
    Dim datetime_now As String, file_path_name As String
    
    
    'Introduction message box for the instructions and purpose
    Dim Msg As String, Title As String, Config As Integer, Ans As Integer
    Msg = "Welcome to the Dormitory search and statistics file creator."
    
    Msg = Msg & vbNewLine & vbNewLine
    
    Msg = Msg & "Do you wish to proceed?"
    
    Msg = Msg & vbNewLine & vbNewLine
    
    Msg = Msg & "Please click OK to proceed and select the "
    Msg = Msg & vbNewLine & "JTC report to use for the file creation in the next step."
    Title = "Create Dormitory Search and Statistics File"
    Config = vbOKCancel + vbQuestion + vbDefaultButton1
    Ans = MsgBox(Msg, Config, Title)
    
    
    
    If Ans = vbOK Then
    
        'open the dialog box for the user to select the file to import
        dms_file_path = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", Title:="Select the JTC report to use for the file creation.")
        
        'If user clicks cancel, the whole program will terminate
        If dms_file_path = False Then
            End
        End If
        
        'open the file that the user has selected and assign it to the variable dms_report
        Set dms_report = Workbooks.Open(dms_file_path)
        
        
        
        'find the current date and time to be used for the file name
        'Use the excel worksheet function to convert the date type value to string
        datetime_now = Application.WorksheetFunction.Text(Now, "DD-MM-YYYY hh-mm-ss")
        
        'define the file path name for the DMS search and statistics summary file
        file_path_name = Application.ActiveWorkbook.Path & "/Current DMS search and summary file as at " & datetime_now & ".xlsx"
        
        
        'save the workbook as a normal excel workbook
        dms_report.SaveAs Filename:=file_path_name, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    Else
        
        End 'terminate the whole program if user selects cancel
    
    End If
    
    ActiveSheet.Name = "Data"
    
End Sub


Sub clean_and_tidy_dataset()
    
    
    'Make the workbook active to make sure the correct workbook is the active workbook
    Workbooks(dms_report.Name).Activate
    
    'Delete the first 3 rows of the report as they are not needed
    Range("A1:A3").EntireRow.Delete
    
    Call UpdateProgress(12 / 100)
    
    'Find the last row of the dataset
    Dim last_row As Integer
    Range("A1").Select
    Selection.End(xlDown).Select
    last_row = ActiveCell.row
    
    
    'Check for rows with any workers that has no room number
    'Workers with no assigned rooms are not to be considered in this report
    Dim var As Integer, no_room As Integer
    no_room = 0
    For var = 2 To last_row
        If IsEmpty(Range(Cells(var, 23), Cells(var, 23))) = True Then
            no_room = no_room + 1
        End If
    Next var
    If no_room > 0 Then 'If there are any workers with no room number, delete those rows
        For var = 1 To no_room
            Range("W2").EntireRow.Delete
        Next var
        
        'Find the new last row of the dataset, if any rows are deleted
        Range("A1").Select
        Selection.End(xlDown).Select
        last_row = ActiveCell.row
        
    End If
    
    
    
    'Delete the unneccesary columns
    Range("A1,D1,H1,I1,K1,L1,N1,O1,R1,U1,V1,X1,Y1").EntireColumn.Delete
    
    Call UpdateProgress(14 / 100)
    
    'Cut and paste the FIN number column as the first column
    Columns("E:E").Select
    Selection.Copy
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    
    
    
    
    
    
    Dim col_index As Integer
    Dim i As Integer
    
    
    Call UpdateProgress(15 / 100)
    
    'create the column for the cluster names
    col_index = Range(Cells(1, 1), Cells(1, 20)).Find("Proximity Card").column
    Call create_new_column(col_index, "Cluster number")
    'populate the column with the cluster name values
    For i = 2 To last_row
        
        Range(Cells(i, col_index), Cells(i, col_index)) = cluster_name(Left(Range(Cells(i, col_index - 1), Cells(i, col_index - 1)), 1))
    
    Next i
    
    Call UpdateProgress(17 / 100)
    
    'create the column for the unit number
    col_index = Range(Cells(1, 1), Cells(1, 20)).Find("Cluster number").column
    Call create_new_column(col_index, "Unit")
    'populate the column with the unit number values
    For i = 2 To last_row
        
        Range(Cells(i, col_index), Cells(i, col_index)) = unit_number(Range(Cells(i, col_index - 1), Cells(i, col_index - 1)))
    
    Next i
    
    Call UpdateProgress(20 / 100)
    
    'create the column for the level number
    col_index = Range(Cells(1, 1), Cells(1, 20)).Find("Unit").column
    Call create_new_column(col_index, "Level")
    'populate the column with the level number values
    For i = 2 To last_row
        
        Range(Cells(i, col_index), Cells(i, col_index)) = Mid(Range(Cells(i, col_index - 1), Cells(i, col_index - 1)), 5, 1)
    
    Next i
    
    Call UpdateProgress(22 / 100)
    
    'create the column for the block number
    col_index = Range(Cells(1, 1), Cells(1, 20)).Find("Level").column 'select the column index
    Call create_new_column(col_index, "Blk")
    'populate the column with the block number values
    For i = 2 To last_row
        
        Range(Cells(i, col_index), Cells(i, col_index)) = Left(Range(Cells(i, col_index - 1), Cells(i, col_index - 1)), 1)
    
    Next i
    
    Call UpdateProgress(25 / 100)
    
    
    'convert the nationality values to a standardized set of values
    col_index = Range(Cells(1, 1), Cells(1, 20)).Find("Nationality").column 'select the column index
    'ignore any error due to wrong data type entered into nationality column
    On Error Resume Next
    'convert each nationality value into the standardized value
    For i = 2 To last_row
        
        Range(Cells(i, col_index), Cells(i, col_index)) = test_nationality(Range(Cells(i, col_index), Cells(i, col_index)))
    
    Next i
    
    Call UpdateProgress(26 / 100)
    
    
    
    'create the column for the count values
    
    Range("S1").value = "Count"
    col_index = Range(Cells(1, 1), Cells(1, 20)).Find("Count").column 'select the column index
    For i = 2 To last_row
    
        Range(Cells(i, col_index), Cells(i, col_index)) = 1
        
    Next i
    
    
    'renaming the column name to company
    Range("I1").value = "Company"
    
    Call UpdateProgress(27 / 100)
    
End Sub



Function cluster_name(block As Integer)
    'A function to return the cluster number based on the block number given in the argument
    'check the block number for the correct cluster
    Select Case block
        Case 1 To 3
            cluster_name = "Cluster 1"
        Case 4 To 5
            cluster_name = "Cluster 2"
        Case 6 To 8
            cluster_name = "Cluster 3"
    End Select

End Function



Function unit_number(digits)
    'A function to return the unit number based on the room number given in the argument
    
    If Len(digits) = 9 Then
        unit_number = Right(digits, 3)
    Else
        If Left(Right(digits, 2), 1) = 0 Then
            unit_number = Right(digits, 1)
        Else
            unit_number = Right(digits, 2)
        End If
    End If

End Function


'define the function to be private for access only within the module
'function will not be available as user-defined function in worksheet

Private Function test_value(expression As String, value As String) As Boolean
    'A function to test for the string expression in the argument against a regular expression rule
    'It will return True if matches
    Dim regexObject As RegExp
    
    Set regexObject = New RegExp
    
    With regexObject
        .Pattern = expression
    End With
    
    test_value = regexObject.Test(value)

End Function


Function test_nationality(nationality As String) As String
    'A function to test and convert the nationality value given in the argument
    'It will be converted into a set of standardized nationality values
    If test_value("[d,D][e,E][s,S][h,H]|[b,B][a,A][n,N][g,G]", nationality) = True Then
        test_nationality = "Bangladesh"
    ElseIf test_value("[i,I][n,N][d,D]", nationality) = True Then
        test_nationality = "India"
    ElseIf test_value("[c,C][h,H]|[p,P][r,R][c,C]", nationality) = True Then
        test_nationality = "China"
    ElseIf test_value("[m,M][y,Y]|[b,B][u,U]", nationality) = True Then
        test_nationality = "Myanmar"
    ElseIf test_value("[t,T][h,H][a,A][i,I]", nationality) = True Then
        test_nationality = "Thailand"
    ElseIf test_value("[m,M][a,A][l,L][a,A][y,Y]", nationality) = True Then
        test_nationality = "Malaysia"
    ElseIf test_value("[v,V][i,I][e,E][t,T]", nationality) = True Then
        test_nationality = "Vietnam"
    ElseIf test_value("[s,S][i,I][n,N]", nationality) = True Then
        test_nationality = "Singaporean PR"
    ElseIf test_value("[f,F][i,I][p,P][i,I]", nationality) = True Then
        test_nationality = "Filipino"
    ElseIf test_value("[p,P][i,I][l,L][i,I]", nationality) = True Then
        test_nationality = "Filipino"
    Else
        test_nationality = nationality
    End If
    
End Function

Sub create_new_column(col_num As Integer, col_name As String)
    'A function that will insert a new column with the column name provided in the argument
    'MsgBox col_num
    Columns(col_num).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range(Cells(1, col_num), Cells(1, col_num)).value = col_name
                
End Sub


