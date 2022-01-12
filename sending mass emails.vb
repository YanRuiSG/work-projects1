Option Explicit

Sub send_emails():


    Dim WkSht As Worksheet
    

    For Each WkSht In ActiveWorkbook.Worksheets
        
        'Call new_formatter1(WkSht)
        
        Dim xOutApp As Object
        Dim xOutMail As Object
        Dim xOutMsg As String, content1 As String, signature As String, content2 As String, content3 As String
        
            

        Dim newOutMsg As String
        
        'message box for alerting if no email address found in database for a company
        If IsEmpty(ActiveSheet.Range("F2").Value) Then MsgBox "No email address found for " & Range("D2").Value & ", follow-up needed for this company."
        
        
        On Error Resume Next
        
        Set xOutApp = CreateObject("Outlook.Application")
        Set xOutMail = xOutApp.CreateItem(0)
        
    
        'CREATING THE CONTENTS TO BE INCLUDED IN THE EMAIL
        content1 = "Dear Employer ,<br/><br/>" & "The following workers from your company can be discharged from the isolation facility on the <b>" & ActiveSheet.Cells(2, 5) & "</b> ,<br/>"

        content1 = content1 & "based on the nominal roll provided by MOM. Should you have any enquiries for your non-listed worker(s), please check with MOM directly.<br/><br/>"
'
'        content2 = "<br/><br/>Attached in this email are the discharge document for each individual worker in the above table. <br/><br/><br/>"
'
'        content2 = content2 & "<b><u>Please take note of the following important points:</u></b><br/>"
'
'        content2 = content2 & "&nbsp&nbsp - 1 . Discharge hours are to be between 9am to 5pm on the discharge date. <br/>"
'
'        content2 = content2 & "&nbsp&nbsp - 2. Please ensure that transport arrangement is being arranged for the pick-up of your workers at the pick-up point.<br/>"
'
'        content2 = content2 & "&nbsp&nbsp - 3. Workers leaving the facility must present his discharge document for verification in PDF/hardcopy version.<br/><br/><br/>"
'
'        signature = "<b><span style=""color:#1F497D"">Thanks and best regards,<br/>Tan Yan Rui<br/>Director of Operations</span style=""color:#335eff""></b>"
'
        'content1 = GetFileContent("C:\Users\Tuas03\Desktop\Mass email sending tool\Email content part 1.txt")
        
        content2 = GetFileContent("C:\Users\Tuas03\Desktop\Mass email sending tool\Email content part 2.txt")
        
        
        Dim lastRow As Integer
        Dim row As Integer
        Dim Table As String
        
        With ActiveSheet
            lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        End With
        
        
        ' creating the table of the list of workers to be included in the email content
        Table = "<table border=""2""><tr><th><b>Name</b></th><th><b>FIN number</b></th><th><b>Contact</b></th><th><b>Company</b></th><th><b>Discharge date</b></th></tr>"
        For row = 2 To lastRow
            'MsgBox ActiveSheet.Cells(row, 1)
            Table = Table & "<tr><td>" & ActiveSheet.Cells(row, 1) & "</td><td>" & ActiveSheet.Cells(row, 2) & "</td><td>" & ActiveSheet.Cells(row, 3) & "</td><td>" & ActiveSheet.Cells(row, 4) & "</td><td>" & ActiveSheet.Cells(row, 5) & "</td></tr>"
        Next row
        
        Table = Table & "</table>"
        
        newOutMsg = content1 & Table & content2 & signature & content3
    
        With xOutMail
            .To = ActiveSheet.Range("F2").Value
            .CC = "yanrui.work@gmail.com" + ";" + ActiveSheet.Range("G2").Value
            .BCC = ""
            .Subject = "Discharge of workers from isolation facility for " & ActiveSheet.Range("D2").Value
            .HTMLBody = newOutMsg
            .Importance = 2
            .Display
            '.Send
        End With
        
        
        Dim worker_name As String, fin_num As String
        
        'Attach the discharge memo for each worker using a for loop
        For row = 2 To lastRow
            worker_name = Range(Cells(row, 1), Cells(row, 1)).Value
            fin_num = Range(Cells(row, 2), Cells(row, 2)).Value
            ' Attached the discharge memo PDF file
            xOutMail.Attachments.Add Application.ActiveWorkbook.Path & "/Discharge memo for " & worker_name & " " & fin_num & ".pdf"
        Next row
        
        
        
        
        Set xOutMail = Nothing
        Set xOutApp = Nothing
        
    Next
        




End Sub



Function GetFileContent(Name As String) As String
' Function for opening and reading the contents of a text file
    Dim intUnit As Integer
    
    On Error GoTo ErrGetFileContent
    
    intUnit = FreeFile
    Open Name For Input As intUnit
    GetFileContent = Input(LOF(intUnit), intUnit)
    
ErrGetFileContent:
    Close intUnit
    Exit Function
End Function









Sub create_memo_main()
    
    'main sub procedure for creating the discharge memo for workers
    
    Dim worker As Integer
    
    'loop through the defined index of workers to create the discharge memo for each worker in PDF
    For worker = 2 To 100 'specify the desired indexes in the excel dataset
        Call create_memo(worker)
    Next worker


End Sub

Sub create_memo(row As Integer)
    'sub procedure for creating an individual discharge memo for a worker
    Dim file_path_name As String, worker_name As String, fin_num As String
    
    Workbooks("Generate discharge documents.xlsm").Activate
    Sheets("Form").Select
    
    'Transferring the data for name
    Range("B4:D4").Select
    Selection.UnMerge
    Sheets("Data").Select
    worker_name = Range(Cells(row, 1), Cells(row, 1)).Value
    Range(Cells(row, 1), Cells(row, 1)).Select
    Selection.Copy
    Sheets("Form").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("B4:D4").Select
    Application.CutCopyMode = False
    Selection.Merge
    
    
    'Transferring the data for FIN number
    Range("B5:D5").Select
    Selection.UnMerge
    Sheets("Data").Select
    fin_num = Range(Cells(row, 2), Cells(row, 2)).Value
    Range(Cells(row, 2), Cells(row, 2)).Select
    Selection.Copy
    Sheets("Form").Select
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("B5:D5").Select
    Application.CutCopyMode = False
    Selection.Merge
    
    
    
    'Transferring the data for Company
    Range("B6:D6").Select
    Selection.UnMerge
    Sheets("Data").Select
    Range(Cells(row, 4), Cells(row, 4)).Select
    Selection.Copy
    Sheets("Form").Select
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("B6:D6").Select
    Application.CutCopyMode = False
    Selection.Merge
    
    
    
    'Transferring the data for barcode
    Range("A9:E9").Select
    Selection.UnMerge
    Sheets("Data").Select
    Range(Cells(row, 5), Cells(row, 5)).Select
    Selection.Copy
    Sheets("Form").Select
    Range("A9").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("A9:E9").Select
    Application.CutCopyMode = False
    Selection.Merge
    
    
     'Transferring the data for barcode
    Sheets("Data").Select
    Range(Cells(row, 6), Cells(row, 6)).Select
    Selection.Copy
    Sheets("Form").Select
    Range("E12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    
    'defining the file path name for the discharge memo based on the active workbook location
    file_path_name = Application.ActiveWorkbook.Path & "/Discharge memo for " & worker_name & " " & fin_num & ".pdf"
    
    'save the discharge memo in pdf format
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        file_path_name _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False

End Sub

Sub new_formatter1(ws As Worksheet):


ws.Activate
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft ' Deleting the first column
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
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
    
    ' Set the column width to autofit the contents
    Columns("A:L").Select
    Columns("A:L").EntireColumn.AutoFit
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select




End Sub
