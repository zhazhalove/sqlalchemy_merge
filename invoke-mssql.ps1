# Install the ImportExcel module if not already installed
# Install-Module -Name ImportExcel -Force -Scope CurrentUser

################### Global variables #######################
$serverName = "localhost\DB2016"
$databaseName = "TSQL2012"
$connString = "Server=$serverName;Database=$databaseName;Integrated Security=True;"
$sqlTemplate = @"
            INSERT INTO [HR].[Employees_copy]
            (empid, lastname, firstname, title, titleofcourtesy, birthdate, hiredate, address, city, region, postalcode, country, phone, mgrid) 
            VALUES (@empid, @lastname, @firstname, @title, @titleofcourtesy, @birthdate, @hiredate, @address, @city, @region, @postalcode, @country, @phone, @mgrid)
"@
$excelFilePath = "$PSScriptRoot\TestData.xlsx"

##################### Load Data from Excel ################################

# Read the data from the Excel file
$excelData = Import-Excel -Path $excelFilePath

# Ensure the data is loaded correctly
if (-not $excelData) {
    Write-Host "Error: Unable to read data from Excel file." -ForegroundColor Red
    exit
}

Write-Host "Data read from Excel file successfully." -ForegroundColor Green


######################## SQL Operations ############################

# Create a connection object
$conn = New-Object System.Data.SqlClient.SqlConnection
$conn.ConnectionString = $connString

try {
    # Open the connection
    $conn.Open()
    Write-Host "Connection opened successfully."

    # Start transaction
    $transaction = $conn.BeginTransaction()
    Write-Host "Transaction started."

    # Create a SQL Command object outside of the loop (singleton pattern)
    $cmd = $conn.CreateCommand()
    $cmd.Transaction = $transaction  # Assign the transaction to the command

    # Loop through each row in the Excel data and insert it into the database
    $excelData | ForEach-Object {
        try {
            $cmd.CommandText = $sqlTemplate

            # Clear previous parameter values
            $cmd.Parameters.Clear()

            # Add parameters for each column
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@empid", [System.Data.SqlDbType]::Int))).Value = $_.empid
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@lastname", [System.Data.SqlDbType]::NVarChar, 20))).Value = $_.lastname
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@firstname", [System.Data.SqlDbType]::NVarChar, 10))).Value = $_.firstname
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@title", [System.Data.SqlDbType]::NVarChar, 30))).Value = $_.title
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@titleofcourtesy", [System.Data.SqlDbType]::NVarChar, 25))).Value = $_.titleofcourtesy
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@birthdate", [System.Data.SqlDbType]::DateTime))).Value = [DateTime]::Parse($_.birthdate)
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@hiredate", [System.Data.SqlDbType]::DateTime))).Value = [DateTime]::Parse($_.hiredate)
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@address", [System.Data.SqlDbType]::NVarChar, 60))).Value = $_.address
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@city", [System.Data.SqlDbType]::NVarChar, 15))).Value = $_.city

            # Check for null value in region and assign DBNull if necessary
            if ($null -eq $_.region) {
                $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@region", [System.Data.SqlDbType]::NVarChar, 15))).Value = [DBNull]::Value
            } else {
                $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@region", [System.Data.SqlDbType]::NVarChar, 15))).Value = $_.region
            }

            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@postalcode", [System.Data.SqlDbType]::NVarChar, 10))).Value = $_.postalcode
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@country", [System.Data.SqlDbType]::NVarChar, 15))).Value = $_.country
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@phone", [System.Data.SqlDbType]::NVarChar, 24))).Value = $_.phone

            # Check for null value in mgrid and assign DBNull if necessary
            if ($null -eq $_.mgrid) {
                $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@mgrid", [System.Data.SqlDbType]::Int))).Value = [DBNull]::Value
            } else {
                $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@mgrid", [System.Data.SqlDbType]::Int))).Value = $_.mgrid
            }

            # Execute the insert command
            $cmd.ExecuteNonQuery() | Out-Null
            Write-Host "Inserted record: $($_.firstname) $($_.lastname)`n`r" -ForegroundColor Green

        } catch {
            Write-Host "SQL Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "$($_.ScriptStackTrace)" -ForegroundColor Red

            # Rollback transaction in case of error
            $transaction.Rollback()
            Write-Host "Transaction rolled back." -ForegroundColor Red
            break
        }
    }

    # Commit the transaction if all records inserted successfully
    $transaction.Commit()
    Write-Host "Transaction committed." -ForegroundColor Green

} catch {
    # Handle connection-level or overall errors
    Write-Host "Connection error" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "$($_.ScriptStackTrace)" -ForegroundColor Red

    # Rollback transaction in case of any overall error
    if ($null -ne $transaction) {
        $transaction.Rollback()
        Write-Host "Transaction rolled back due to overall error." -ForegroundColor Red
    }

} finally {
    # Ensure the connection is closed
    if ($conn.State -eq 'Open') {
        $conn.Close()
        Write-Host "Connection closed."
    }
}

