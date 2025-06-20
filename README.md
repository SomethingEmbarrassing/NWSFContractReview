# NWSFContractReview
Automated sheet for new job start up

To do
Pull additional data from "F:\JOB LIST\JOB LIST2.xlsx" if cell F1 or E2 are blank
  -if job list2.xlsx is locked, open as read only. Pull data without opening file if possible. if opening is         necessary, screen flicker off from start until finish.
  -always use sheet "Add Jobs Here"
  -Find Job # in column C
  -copy Column A data under "PM" in the table, paste to E1 of destination sheet
  -copy column J data under "Ton" in the table, paste into E2 of destination sheet


This project includes VBA code extracted from an Excel .xlsm file for easier version control and collaboration. The code is organized as follows:

    WorkbookEvents.bas – Code from the ThisWorkbook object
    Contract Review Sheet1.bas – Code behind Sheet1 
