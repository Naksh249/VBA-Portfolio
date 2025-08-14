# VBA-Portfolio

# Outlook Attachment Saver Macro

This VBA macro automates saving attachments from multiple selected Outlook emails to a specified network folder. Instead of opening each email individually and manually saving attachments, the macro streamlines the process into a single click, ensuring consistent file naming and avoiding overwriting issues.

## Features

- **Batch Processing:** Save attachments from multiple selected emails at once.
- **Custom Save Location:** Specify a network folder or local directory for saving attachments.
- **Consistent Naming:** Automatically names files to prevent overwriting and maintain organization.
- **User-Friendly:** No need to open each email or save attachments manually—just select emails and run the macro.

## Usage

1. Select one or more emails in Outlook.
2. Run the macro from the VBA editor or assign it to a button for one-click use.
3. All attachments from the selected emails will be saved to your chosen folder.

## Setup

1. Open Outlook and press `ALT + F11` to open the VBA editor.
2. Insert a new module and paste the macro code into it.
3. Update the folder path in the macro to your desired network or local folder.
4. (Optional) Add a button to your Outlook ribbon or Quick Access Toolbar for easier access.

## Example

```vba
' Example usage within the macro:
Const saveFolder As String = "\\NetworkPath\Attachments"
```

## Notes

- Ensure you have write permission to the selected network folder.
- The macro handles duplicate file names by appending a number or timestamp to the filename.
- This macro is designed for use with Microsoft Outlook.

## License

This project is open-source and free to use. Attribution appreciated!

---

**Happy automating!**

“Move SaveAttachments.bas to outlook-attachment-saver folder
