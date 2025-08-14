# Excel Data Import Macro

This VBA macro automates importing data from multiple Excel files into a master sheet. Instead of manually copying and pasting, the user selects the source file, and the macro dynamically copies data from `A2:Z` (skipping headers) down to the last filled row. It then pastes the data as values into the next available empty row of the master sheet.

## Features

- **Automated Import:** Select an Excel file, and the macro pulls data without manual copy-paste.
- **Dynamic Range:** Copies from `A2` to `Z` of the last used row, skipping headers.
- **Consolidates Data:** Appends imported data to the next available row in your master sheet.
- **Paste as Values:** Ensures only data—not formulas or formatting—is brought across.
- **User-Friendly:** Simple file selection dialog; no need for advanced Excel skills.

## How It Works

1. **Prompt:** When you run the macro, a dialog asks you to select an Excel file.
2. **Import:** The macro copies all rows from the source file’s first sheet (from A2 to the last filled row in column Z).
3. **Paste:** Data is pasted as values into the next empty row of your master sheet.
4. **Repeat:** You can run the macro multiple times to consolidate data from multiple files.

## Setup

1. Open your master workbook in Excel.
2. Press `ALT + F11` to open the VBA editor.
3. Insert a new module and paste in the macro code.
4. (Optional) Adjust `wsDest = Sheet7` to use your preferred worksheet.
5. Save your workbook as a macro-enabled file (`.xlsm`).

## Usage

- Run the macro (`CopyDynamicData`) from the VBA editor or assign it to a button.
- When prompted, select the Excel file you want to import.
- The data will be appended to your master sheet. A message box will confirm completion.

## Customization

- Change the range `A2:Z` if your data uses different columns.
- Update `wsDest = Sheet7` to match your destination sheet (e.g., `Set wsDest = Worksheets("Master")`).

## License

Open-source for learning and productivity purposes. Attribution appreciated!

---

**Happy automating!**
