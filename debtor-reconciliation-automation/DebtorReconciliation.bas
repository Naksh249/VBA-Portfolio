Dim LastRow As Long

Sub Clear_Content()
    Dim Answer As VbMsgBoxResult
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("RawData")
    
    'Find the last used row in column A
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Answer = MsgBox("Are you sure you wish to proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Clear cells")
    If Answer = vbYes Then
        ws.Range("A2:M" & LastRow).ClearContents 'only clears content not formatting
    Else
        Exit Sub
    End If
End Sub

Sub Copy_Content()
    Dim SourceWB As Workbook
    Dim TargetWB As Workbook
    Dim SourceWS As Worksheet
    Dim TargetWS As Worksheet
    Dim VisibleCells As Range
    Dim StartCell As Range
    Dim filePath As String
    Dim LastRow As Long
    Dim BranchColumn As String
    Dim BranchName As String
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
         
    ' Ask user to select file
    filePath = Application.GetOpenFilename( _
                Title:="Select the Excel file to import from", _
                FileFilter:="Excel Files (*.xls*), *.xls*")
    ' Exit if cancelled
    If filePath = "False" Then Exit Sub
    
    ' Open source workbook
    Set SourceWB = Workbooks.Open(filePath)
    Set SourceWS = SourceWB.Sheets("Outstanding")

    ' Ask user which branch to filter
    BranchColumn = InputBox("Enter the column letter for Branch (e.g., A):", "Branch Column", "A")
    BranchName = InputBox("Enter the Branch name to filter:", "Select Branch")
    
    ' Apply AutoFilter
    LastRow = SourceWS.Cells(SourceWS.Rows.Count, BranchColumn).End(xlUp).Row
    SourceWS.Range("A6:M" & LastRow).AutoFilter Field:=Range(BranchColumn & "6").Column, Criteria1:=BranchName

    'Set the target workbook and sheet
    Set TargetWB = ThisWorkbook
    Set TargetWS = TargetWB.Sheets("RawData")
    
    'Get only visible cells after filter
    On Error Resume Next
    Set VisibleCells = SourceWS.Range("A7:M" & LastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not VisibleCells Is Nothing Then
        'Copy visible cells only
        VisibleCells.Copy
        TargetWS.Range("A1").PasteSpecial xlPasteValues
        MsgBox "Data copied successfully!", vbInformation
    Else
        MsgBox "No visible cells found to copy!", vbInformation
    End If
    
    ' Close source without saving
    SourceWB.Close SaveChanges:=False
    
    ' Clear filter
    SourceWS.AutoFilterMode = False
    
    'Clear clipboard
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub RefreshAllPowerQueries()
    ThisWorkbook.RefreshAll
    MsgBox "All Power Queries have been refreshed!", vbInformation
End Sub
