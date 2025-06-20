# NWSFContractReview
Automated sheet for new job start up

## Overview
When the workbook opens it reads the job number and name from the file path. If the **PM** (cell `E1`) or **Ton** (cell `E2`) fields are blank, the macro `FillFromJobList` retrieves the information from `F:\JOB LIST\JOB LIST2.xlsx` (sheet **Add Jobs Here**). The source workbook is opened read‑only if necessary and closed again without user interaction.

## Code Files
This project includes VBA code extracted from the Excel workbook for easier version control:

* **WorkbookEvents.bas** – code from the `ThisWorkbook` object containing `Workbook_Open`, `GetFilePath` and `FillFromJobList`.
