# Project: Automated Payment Run Proposal Emails (Excel VBA + Outlook)

## Summary
Developed a VBA tool to generate and send weekly payment run approval emails with attached reports and dynamically populated totals.

## Problem
- Manual process took 20–30 minutes weekly.
- Subject line, tables, and notes had to be manually copied from Excel into Outlook.
- High risk of errors (missing values, wrong recipients).

## Solution
VBA button in Excel automates:
- Extracting totals tables.
- Creating Outlook email with dynamic subject line.
- Embedding formatted tables in the body.
- Attaching the payment run report.
- Sending to the correct manager group.

## Impact
- Reduced process time from ~30 minutes to <1 minute.
- Eliminated copy-paste errors.
- Ensured consistent, professional reporting every week.

> 💡 **This project ticks so many boxes:** *VBA + Outlook integration, automation of a real finance process, measurable efficiency gain.*

---

## Example VBA Code

```vba
Sub SendPaymentRunEmail_CleanFormatted()
    Dim OutApp As Object, OutMail As Object
    Dim rngTotals As Range, rngManual As Range
    Dim strBody As String, strSubject As String
    Dim PaymentDate As String
    Application.ScreenUpdating = False ' Stop screen flickering
    '--- Named ranges ---
    Set rngTotals = Range("PTotal")
    Set rngManual = Range("MPayments")
    
    '--- Payment date from New Summary!C1 ---
    PaymentDate = Sheets("New Summary").Range("C1").Value
    
    '--- Subject ---
    strSubject = "Payment Run Proposal " & PaymentDate
    
    '--- Build email body ---
    strBody = "<p>Good afternoon,</p>" & _
              "<p>Please find attached payment run proposal, totals as below.</p>" & _
              BuildHTMLTable(rngTotals) & _
              "<br>" & _
              BuildHTMLTable(rngManual) & _
              "<p>Kind regards,<br>Lauren</p>"
    
    '--- Create email ---
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .To = "manager1@email.com; manager2@email.com" 'update with recipients
        .CC = "" ' optional
        .BCC = "" ' optional
        .Subject = strSubject
        .HTMLBody = strBody
        .Attachments.Add ThisWorkbook.FullName
        .Display   'use .Send to send automatically
    End With
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    Application.ScreenUpdating = True
End Sub

'--- Function to build clean, formatted HTML table ---
Function BuildHTMLTable(rng As Range) As String
    Dim r As Range, c As Range
    Dim HTML As String
    Dim rowIndex As Long, colIndex As Long
    Dim cellValue As String
    
    HTML = "<table style='border-collapse:collapse; font-family:Calibri; font-size:11pt;'>"
    
    rowIndex = 0
    For Each r In rng.Rows
        rowIndex = rowIndex + 1
        If rowIndex = 1 Then
            ' Header row
            HTML = HTML & "<tr style='background-color:#f2f2f2;'>"
        ElseIf rowIndex Mod 2 = 0 Then
            ' Even row shading
            HTML = HTML & "<tr style='background-color:#fafafa;'>"
        Else
            HTML = HTML & "<tr>"
        End If
        
        colIndex = 0
        For Each c In r.Cells
            colIndex = colIndex + 1
            cellValue = c.Value
            
            ' Format numbers if numeric
            If IsNumeric(cellValue) And cellValue <> "" Then
                cellValue = Format(cellValue, "#,##0.00")
                If rowIndex = 1 Then
                    HTML = HTML & "<th style='border:1px solid #999; padding:6px; text-align:right;'>" & cellValue & "</th>"
                Else
                    HTML = HTML & "<td style='border:1px solid #999; padding:6px; text-align:right;'>" & cellValue & "</td>"
                End If
            Else
                ' Treat as text
                If rowIndex = 1 Then
                    HTML = HTML & "<th style='border:1px solid #999; padding:6px; text-align:left;'>" & cellValue & "</th>"
                Else
                    HTML = HTML & "<td style='border:1px solid #999; padding:6px; text-align:left;'>" & cellValue & "</td>"
                End If
            End If
        Next c
        HTML = HTML & "</tr>"
    Next r
    
    HTML = HTML & "</table>"
    
    BuildHTMLTable = HTML
End Function
```

---

*It’s the kind of project recruiters love to see because they instantly picture it being useful in their own teams.*