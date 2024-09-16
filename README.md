# SkywardFinDashboard

A PowerShell script that can take data exported from Skyward into two files: Detail Dashboard / Summary Dashboard and put onto a specific Google Sheet that is linked to a financial dashboard.

You will need to create a scheduled job in Skyward to export these .CSVs.

Once the files are configured to automatically export, simply specify the $certPath, $iss, $certPswd, $spreadsheetID, $sheetName, $sourceCSV, $spreadsheetIdSum, $sheetNameSum, $sourceCSVSum

Basically, we ran some cleanup against the source data, cleaned temporary header rows, removed the top row, imported the file, cleared the existing sheet, put the data back on, and then finally, cleared the top row again.

After cleaning the detail dashboard, we import the Summary Dashboard in much the same way.

This is required when going to Google because if you simply upload a CSV to Google, each time a new file is uploaded, the ID will change.  This wouldn't be an issue with OneDrive/Sharepoint since those use relative paths instead.
