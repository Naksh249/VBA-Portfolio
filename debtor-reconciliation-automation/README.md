# Debtor Reconciliation Automation

This project builds upon and enhances an existing Power Query–driven debtor reconciliation workflow by introducing VBA automation to streamline the full data refresh process in Excel. The macros are designed to work alongside your existing Power Query setup in the same workbook, adding the ability to:

- **Clear previous raw data** from a dedicated sheet.
- **Import new data** from a source workbook, with branch-specific filtering.
- **Refresh all Power Query transformations** and reconciliation logic with a single command.

By integrating these VBA macros with your established Power Query solution, the reconciliation process becomes truly end-to-end, allowing users to manage data import, transformation, and automation seamlessly from a single interface.

## Features

- **Clear Data Sheet:** Removes old raw data, readying the sheet for new imports.
- **Import with Filtering:** Lets users select and filter source data by branch before importing.
- **Automated Power Query Refresh:** Triggers all Power Queries and logic for up-to-date reconciliation.

## Usage

1. **Clear_Content:** Clears previous data in the "RawData" sheet (except headers).
2. **Copy_Content:** Prompts user to select the source file and branch, imports filtered data into "RawData".
3. **RefreshAllPowerQueries:** Refreshes all Power Queries in the workbook—integrating seamlessly with your existing Power Query reconciliation setup.

## Setup

- Copy the code from `DebtorReconciliation.bas` into a VBA module in your workbook.
- Ensure sheet names ("RawData", "Outstanding") match your workbook or update them in the code as needed.
- Save the workbook as a macro-enabled file (`.xlsm`).
- These macros are intended to complement and automate your existing Power Query logic in the same workbook.

## Note

- No sensitive or internal network paths are included in this code.
- Modify column ranges and sheet names as required for your specific workflow.

---

**Take your Power Query reconciliation to the next level with automation!**
