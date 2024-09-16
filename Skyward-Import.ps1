# Import the UMN-Google PowerShell module
Import-Module UMN-Google

# Google API Authozation
$scope = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/drive.file"
# Credential cert downloaded from Google
$certPath = "C:\Scripts\yourcredcert.p12"
# Google Service Account Path
$iss = 'gserviceaccount'
# Cert Password
$certPswd = 'certpassword'

# Define the ID of the Detail Dashboard spreadsheet you want to manipulate
$spreadsheetId = "ID of Detail Spreadsheet"
# Define the Spreadsheet information
$sheetName = "Detail Dashboard"
# Source CSV
$sourceCSV = "c:\Financial Dashboard\Detail Dashboard.csv"

# Define the ID of the Summary Dashboard spreadsheet you want to manipulate
$spreadsheetIdSum = "ID of Summary Spreadsheet"
# Define the Spreadsheet information
$sheetNameSum = "Summary Dashboard"
# Source CSV
$sourceCSVSum = "c:\Financial Dashboard\Summary Dashboard.csv"

# Define the logging file path
$LoggingFilePath = "Update.log"

# Function to write to both console and log file
function Write-HostAndFile {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message
    )

    Write-Host $Message
    [System.IO.File]::AppendAllLines($LoggingFilePath, [string []]$Message)
}



function Set-GoogleAPIAuthorization {
    # Set security protocol to TLS 1.2 to avoid TLS errors
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    try {
        $accessToken = Get-GOAuthTokenService -scope $scope -certPath $certPath -certPswd $certPswd -iss $iss
        return $accessToken
    } catch {
        $err = $_.Exception
        $err | Select-Object -Property *
        "Response: "
        $err.Response
    }
}

function Clean-CSVFile {
    param (
        [Parameter(Mandatory=$true)]
        [string]$sourceCSVFunc
    )

    # Import the CSV file as a text array
    $data = Get-Content -Path $sourceCSVFunc

    # Skip the first two lines (old header and second row)
    $data = $data | Select-Object -Skip 1

    # Define new headers
    $newHeaders = '"H1","H2","H3","H4","H5","H6","H7","H8","H9","H10","H11","H12","H13",'

    # Add the new headers at the beginning
    $data = $newHeaders, $data

    # Write the modified data back to the file
    $data | Set-Content -Path $sourceCSVFunc
}

function CleanupCommas-CSVFile {
    param (
        [Parameter(Mandatory=$true)]
        [string]$csvFilePath
    )

    # Check if the file contains any quotes
    if ((Get-Content $csvFilePath) -like '*"*') {
        # Create a new TextFieldParser
        Add-Type -AssemblyName Microsoft.VisualBasic
        $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($csvFilePath)

        # Configure the TextFieldParser
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.Delimiters = ','
        $parser.HasFieldsEnclosedInQuotes = $true

        # Create an ArrayList to store the cleaned lines
        $cleanedCsvFile = New-Object System.Collections.ArrayList($null)

        # Process the CSV file
        while (!$parser.EndOfData) {
            # Read the fields of the current line
            $fields = $parser.ReadFields()

            # Process each field
            for ($i = 0; $i -lt $fields.Length; $i++) {
                # Remove commas from the field
                $fields[$i] = $fields[$i] -replace ',', ''
            }

            # Join the fields with commas and add the cleaned line to the ArrayList
            $cleanedLine = $fields -join ','
            $null = $cleanedCsvFile.Add($cleanedLine)
        }

        # Close the TextFieldParser
        $parser.Close()

        # Write the cleaned lines back to the CSV file
        $cleanedCsvFile | Out-File -FilePath $csvFilePath

        Write-Host "CSV file cleaned successfully."
    } else {
        Write-Host "CSV file does not contain any quotes. Skipping cleanup."
    }
}

function Import-CSVFileWithoutHeader {
    param (
        [Parameter(Mandatory=$true)]
        [string]$csvFilePath
    )

    # Import CSV
    $importNoHeader = New-Object System.Collections.ArrayList($null)

    # Read the CSV file as a text file
    $csvFile = Get-Content -Path $csvFilePath

    # Split the first line into headers
    $headers = $csvFile[0] -split ','

    # Create a hashtable to store the unique headers and their counts
    $headerCounts = @{}

    # Process the headers
    for ($i = 0; $i -lt $headers.Length; $i++) {
        # If the header is already in the hashtable, increment its count and append the count to the header name
        if ($headerCounts.ContainsKey($headers[$i])) {
            $headerCounts[$headers[$i]]++
            $headers[$i] += '_' + $headerCounts[$headers[$i]]
        } else {
            # If the header is not in the hashtable, add it with a count of 0
            $headerCounts[$headers[$i]] = 0
        }
    }

    # Process the rest of the CSV file
    for ($i = 0; $i -lt $csvFile.Length; $i++) {
        # Split the current line into fields
        $fields = $csvFile[$i] -split ','

        # Create an array for the current row
        $rowArray = @()

        # Add each field to the row array with its corresponding header as the key
        for ($j = 0; $j -lt $fields.Length; $j++) {
            if ($j -lt $headers.Length) {
                # Remove the double quotes around the field value
                $fieldValue = $fields[$j] -replace '^"|"$', ''
                $rowArray += $fieldValue
            }
        }

        # Add the row array to the import array
        $null = $importNoHeader.Add($rowArray)
    }

    # Return the imported data
    return $importNoHeader
}

function Clear-GoogleSheetData {
    param (
        [Parameter(Mandatory=$true)]
        [string]$sheetId,
        [Parameter(Mandatory=$true)]
        [string]$token,
        [Parameter(Mandatory=$true)]
        [string]$nameOfSheet
    )

    # Make a request to the Google Sheets API to retrieve the spreadsheet data
    try {
        $sheetResponse = Invoke-RestMethod -Uri "https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?includeGridData=true" -Method Get -Headers @{ 'Authorization' = "Bearer $token" }
        Write-Host "Response: $sheetResponse"
    } catch {
        Write-Host "Error: $_"
        Write-Host "Error details: $($_.Exception.Response)"
    }

    # Get the number of rows
    $numRows = $sheetResponse.sheets[0].data[0].rowData.Count
    Write-Host "There are currently: " $numRows

    # Check if there are rows to clear
    if ($numRows -gt 0) {
        # Define the range to clear
        $clearRange = "'${nameOfSheet}'!1:${numRows}"

        # Make a request to the Google Sheets API to clear the data in the range
        Invoke-RestMethod -Uri "https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${clearRange}:clear" -Method POST -Headers @{ 'Authorization' = "Bearer $token" }
    } else {
        Write-Host "No rows to clear."
    }
 }

function UploadToGSheet {
    param(
        [Parameter(Mandatory=$true)]$data,
        [Parameter(Mandatory=$true)]$token,
        [Parameter(Mandatory=$true)]$name,
        [Parameter(Mandatory=$true)]$ID
    )
    try {
        $data | ForEach-Object { Write-Host "Data for upload: $($_ -join ', ')" }
        Set-GSheetData -accessToken $token -rangeA1 "A1:N$($data.Count)" -sheetName $name -spreadSheetID $ID -values $data -Debug -Verbose
    } catch {
        $errorVar = $_.Exception
        $errorVar | Select-Object -Property *
        "Response: "
        $errorVar.Response
    }
}

function UploadToGSheetNative {
    param(
        [Parameter(Mandatory=$true)]$data,
        [Parameter(Mandatory=$true)]$token,
        [Parameter(Mandatory=$true)]$name,
        [Parameter(Mandatory=$true)]$ID
    )
    try {
        $data | ForEach-Object { Write-Host "Data for upload: $($_ -join ', ')" }

        # Convert all data to strings and handle empty cells
        $data = $data | ForEach-Object {
            $_ | ForEach-Object {
                if ($_ -eq $null -or $_ -eq '') {
                    # Replace empty cells with an empty string
                    $_ = ''
                } else {
                    # Convert non-empty cells to strings
                    $_ = $_.ToString()
                }
            }
        }

        # Calculate the number of rows and columns based on the data
        $numRows = $data.Count
        $numCols = ($data | Measure-Object -Property Count -Maximum).Maximum

        # Ensure $numCols is an integer
        $numCols = [int]$numCols

        # Define the range dynamically
        $range = "A1:$([System.Convert]::ToChar(64 + $numCols))$numRows"
            
        # Define the URL for the Google Sheets API
        $url = "https://sheets.googleapis.com/v4/spreadsheets/$ID/values/$name!$range?valueInputOption=USER_ENTERED"

        # Define the headers for the request
        $headers = @{
            "Authorization" = "Bearer $token"
            "Content-Type" = "application/json"
        }

        # Convert the data to JSON
        $body = $data | ConvertTo-Json

        # Make the request to the Google Sheets API
        Invoke-WebRequest -Uri $url -Method PUT -Headers $headers -Body $body
    } catch {
        $errorVar = $_.Exception
        Write-Host "Exception Message: $($errorVar.Message)"
        Write-Host "Exception Item: $($errorVar.ItemName)"
        Write-Host "Exception Response: $($errorVar.Response)"
    }
}

function ClearTopRow {
    param(
        [Parameter(Mandatory=$true)]
        [string]$sheetId,
        [Parameter(Mandatory=$true)]
        [string]$token,
        [Parameter(Mandatory=$true)]
        [string]$nameOfSheet
    )

    # Define the range to clear
    $clearRange = "'${nameOfSheet}'!A1:Z1"

    Write-Host "Here is the range:" $clearRange

    # Make a request to the Google Sheets API to clear the data in the range
    $clearUri = "https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${clearRange}:clear"
    Write-Host "Clear URI:" $clearUri
    Invoke-RestMethod -Uri $clearUri -Method POST -Headers @{ 'Authorization' = "Bearer $token" }
}

function ClearTopRowCSV {
    param (
        [Parameter(Mandatory=$true)]
        [string]$csvFilePath
    )

    # Read the CSV file as a text file
    $csvFile = Get-Content -Path $csvFilePath

    # Check if the first row is empty or contains only commas surrounded by quotes
    $isEmptyOrCommas = ($csvFile[0] -replace '","', '') -eq '""'

    if ($isEmptyOrCommas) {
        # If the first row is empty or contains only commas surrounded by quotes, remove it
        $csvFile = $csvFile | Select-Object -Skip 1

        # Write the modified content back to the CSV file
        $csvFile | Set-Content -Path $csvFilePath
    }
}

function Remove-TopEmptyRows {
    param (
        [string]$spreadsheetId,
        [string]$sheetName,
        [string]$accessToken
    )

    # Define the URL for the Google Sheets API
    $url = "https://sheets.googleapis.com/v4/spreadsheets/$spreadsheetId/values/$sheetName"

    # Make the GET request to get the data
    $response = Invoke-RestMethod -Uri $url -Method Get -Headers @{"Authorization" = "Bearer $accessToken"}

    # Find the first non-empty row
    $firstNonEmptyRow = $response.values | ForEach-Object { if ($_ -ne $null) { return $_ } }

    # Debug: Print the first non-empty row
    Write-Host "First non-empty row: $firstNonEmptyRow"

    # If the first row is not the first non-empty row, delete the empty rows
    if ($firstNonEmptyRow -ne $response.values[0]) {
        # Define the URL for the batchUpdate method
        $url = "https://sheets.googleapis.com/v4/spreadsheets/$spreadsheetId:batchUpdate"

        # Define the request body
        $body = @{
            "requests" = @(
                @{
                    "deleteDimension" = @{
                        "range" = @{
                            "sheetId" = $spreadsheetId
                            "dimension" = "ROWS"
                            "startIndex" = 0
                            "endIndex" = [array]::IndexOf($response.values, $firstNonEmptyRow)
                        }
                    }
                }
            )
        } | ConvertTo-Json

        # Debug: Print the request body
        Write-Host "Request body: $body"

        # Make the POST request to delete the rows
        try {
            $response = Invoke-RestMethod -Uri $url -Method Post -Headers @{"Authorization" = "Bearer $accessToken"} -Body $body -ContentType "application/json"
        } catch {
            # Debug: Print the error message
            Write-Host "Error: $($_.Exception.Message)"
        }
    }
}

# Call the function and assign the returned value to a variable
$accessToken = Set-GoogleAPIAuthorization

# Cleanup and Import Detail Dashboard
    # Clean the source CSV so it can be imported (add temp header rows - will need to be removed later)
    Clean-CSVFile -sourceCSV $sourceCSV
     
    # Remove the top row from the CSV
    ClearTopRowCSV -csvFilePath $sourceCSV

    # Cleanup Commas from CSV
    CleanupCommas-CSVFile -csvFilePath $sourceCSV

    # Import the CSV
    $importNoHeaderDet = Import-CSVFileWithoutHeader -csvFilePath $sourceCSV

    Write-Host "Data from Detail Dashboard:"
    $importNoHeaderDet | ForEach-Object { Write-Host "Row: $($_ -join ', ')" }

    # Clear all existing data from the sheet
    Clear-GoogleSheetData -sheetId $spreadsheetId -token $accessToken -nameOfSheet $sheetName

    # Upload the new data to the Google Sheet
    UploadToGSheet -data $importNoHeaderDet -token $accessToken -name $sheetName -ID $spreadsheetId

    # Fix the header row on the Detail Dashboard
    ClearTopRow -sheetId $spreadsheetId -token $accessToken -nameOfSheet $sheetName
    #Remove-TopEmptyRows -spreadsheetId $spreadsheetId -sheetName $sheetName -accessToken $accessToken

# Cleanup and Import Summary Dashboard
   
    # Cleanup Commas from CSV
    CleanupCommas-CSVFile -csvFilePath $sourceCSVSum
    
    # Import the CSV
    $importNoHeaderSum = Import-CSVFileWithoutHeader -csvFilePath $sourceCSVSum

    # Clear all existing data from the sheet
    Clear-GoogleSheetData -sheetId $spreadsheetIdSum -token $accessToken -nameOfSheet $sheetNameSum
        
    Write-Host "Data from Summ Dashboard:"
    $importNoHeaderSum | ForEach-Object { Write-Host "Row: $($_ -join ', ')" }

    # Upload the new data to the Google Sheet
    UploadToGSheet -data $importNoHeaderSum -token $accessToken -name $sheetNameSum -ID $spreadsheetIdSum