Attribute VB_Name = "SaveFIle2"
Sub SaveAllAttachmentsFromSelectedEmails()
    Dim olItem As Object
    Dim olAttachment As Attachment
    Dim saveFolder As String
    Dim ns As Outlook.NameSpace
    Dim Selection As Outlook.Selection
    Dim safeFileName As String
    Dim fullFilePath As String
    Dim i As Integer

    ' Set your UNC network path
    saveFolder = "\\DTSRVHQFS\Shared\Accts Dept\Accounts\Purchase Ledger\Processing\Coopervision Import\PDF Invoices to upload\"

    ' Ensure the path ends with a backslash
    If Right(saveFolder, 1) <> "\" Then
        saveFolder = saveFolder & "\"
    End If

    ' Validate folder existence
    If Dir(saveFolder, vbDirectory) = "" Then
        MsgBox "ERROR: Folder path does not exist:" & vbCrLf & saveFolder, vbCritical
        Exit Sub
    End If

    Set ns = Application.GetNamespace("MAPI")
    Set Selection = Application.ActiveExplorer.Selection

    For Each olItem In Selection
        If TypeOf olItem Is Outlook.MailItem Then
            For Each olAttachment In olItem.Attachments
                ' Clean and prepare file name
                safeFileName = CleanFileName(olAttachment.FileName)
                fullFilePath = saveFolder & safeFileName

                ' Avoid overwriting existing files
                i = 1
                Do While Dir(fullFilePath) <> ""
                    fullFilePath = saveFolder & "(" & i & ")_" & safeFileName
                    i = i + 1
                Loop

                ' Attempt to save the attachment with error handling
                On Error Resume Next
                olAttachment.SaveAsFile fullFilePath
                If Err.Number <> 0 Then
                    MsgBox "FAILED to save: " & fullFilePath & vbCrLf & _
                           "Error: " & Err.Description, vbExclamation
                    Err.Clear
                End If
                On Error GoTo 0
            Next
        End If
    Next

    MsgBox "Attachments processed and saved to:" & vbCrLf & saveFolder, vbInformation
End Sub

' Helper function to remove invalid characters from file names
Function CleanFileName(strFileName As String) As String
    Dim invalidChars As Variant
    Dim i As Integer

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
3
    For i = LBound(invalidChars) To UBound(invalidChars)
        strFileName = Replace(strFileName, invalidChars(i), "_")
    Next

    CleanFileName = strFileName
End Function

