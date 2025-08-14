Attribute VB_Name = "CopyAndPaste"
Option Explicit

Sub CopyDynamicData()


    Dim wsDest As Worksheet
    Dim wsSource As Worksheet
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    Dim rngSource As Range
    Dim filePath As String
    Dim wbSource As Workbook
    Dim wbDest As Workbook
    Dim FileToOpen As Variant

'Set destination workbook and sheet
Set wbDest = ThisWorkbook
Set wsDest = Sheet7

    Application.ScreenUpdating = False
    Application.EnableEvents = False


'Ask user to select file
filePath = Application.GetOpenFilename( _
    Title:="Select the Excel file to import fromt", _
    FileFilter:="Excel Files (*.xls*), *.xls*")
    
'Exit if they cancel
If filePath = "False" Then Exit Sub

'Open source workbook
Set wbSource = Workbooks.Open(filePath)

'Set source sheet
Set wsSource = wbSource.Sheets(1)

'Find last row in column A of source
lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row


'Define range from A2 to Z(lastRowSource)
If lastRowSource >= 2 Then
    Set rngSource = wsSource.Range("A2:Z" & lastRowSource)
    
'Find next empty row in destination sheet
lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1

'Copy and paste values

        rngSource.Copy
        wsDest.Range("A" & lastRowDest).PasteSpecial xlPasteValues
        Application.CutCopyMode = False


End If

'Close source workbook without saving
wbSource.Close savechanges:=False

MsgBox "Data Imported successfully!", vbInformation

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

