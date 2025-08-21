# Automated Payment Run Proposal Emails (Excel VBA + Outlook)

This VBA macro automates the process of generating and sending weekly payment run approval emails directly from Excel. Instead of manually copying tables, subject lines, and notes into Outlook, this tool streamlines the workflow—ensuring accuracy, consistency, and significant time savings.

## Features
- **One-Click Emailing:** Send a fully formatted payment run proposal email with a single button press in Excel.
- **Dynamic Content:** Automatically extracts payment totals and manual payment tables from your workbook and embeds them as clean HTML tables in the email body.
- **Automated Subject & Body:** Generates a dynamic subject line and a professional HTML-formatted email, including payment date and all necessary details.
- **Attachment Handling:** Attaches the current Excel workbook to the email for manager review.
- **Correct Recipients:** Pre-fills the correct recipient group to eliminate the risk of sending to the wrong manager.
- **Consistent Formatting:** Ensures every email looks professional and contains all relevant data.

## Usage
1. Update the named ranges (`PTotal` and `MPayments`) and the payment date cell in your Excel workbook.
2. Press the assigned macro button in Excel to generate the email.
3. Review the pre-filled Outlook email and send (or use `.Send` to send automatically).

## Setup
- Open your Excel workbook and press ALT + F11 to open the VBA editor.
- Insert a new module and paste the provided macro code.
- Update the recipient email addresses and named ranges as needed.
- (Optional) Add a button to your worksheet for one-click operation.

## Example
```vba
' Example usage within the macro:
Set rngTotals = Range("PTotal")
Set rngManual = Range("MPayments")
PaymentDate = Sheets("New Summary").Range("C1").Value
```

## Notes
- The macro is designed for use with Microsoft Excel and Outlook.
- All formatting and totals are handled automatically—no manual editing required.
- Saves time and reduces the risk of errors compared to manual email preparation.

## License
This project is open-source and free to use. Attribution appreciated!

Happy automating!
