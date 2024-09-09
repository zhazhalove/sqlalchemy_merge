################### Global variables #######################
$serverName = "localhost\DB2016"
$databaseName = "TSQL2012"
$connString = "Server=$serverName;Database=$databaseName;Integrated Security=True;"
$sqlTemplate = @"
            INSERT INTO [HR].[Employees_copy]
            (lastname, firstname, title, titleofcourtesy, birthdate, hiredate, address, city, region, postalcode, country, phone, mgrid) 
            VALUES (@lastname, @firstname, @title, @titleofcourtesy, @birthdate, @hiredate, @address, @city, @region, @postalcode, @country, @phone, @mgrid)
"@


##################### Test Data ################################

# generate test data
$random = New-Object System.Random
$firstNames = @("John", "Jane", "Mike", "Emily", "Jake", "Laura", "David", "Mia", "Tom", "Anna")
$lastNames = @("Doe", "Smith", "Johnson", "Brown", "Davis", "Wilson", "Anderson", "Taylor", "Moore", "Clark")
$titles = @("Manager", "Engineer", "Technician", "Clerk", "Director")
$titleOfCourtesy = @("Mr.", "Mrs.", "Ms.", "Dr.")
$cities = @("New York", "Los Angeles", "Chicago", "Houston", "Phoenix")
$countries = @("USA", "Canada", "Mexico", "UK", "Germany")
$addresses = @("123 Main St", "456 Oak St", "789 Pine St", "101 Maple Ave", "202 Elm St")

$sqlData = @()

Write-Host "Generating Test Data" -ForegroundColor Green

for ($i = 1; $i -le 50000; $i++) {
    $randomFirst = $firstNames[$random.Next(0, $firstNames.Length)]
    $randomLast = $lastNames[$random.Next(0, $lastNames.Length)]
    $randomTitle = $titles[$random.Next(0, $titles.Length)]
    $randomCourtesy = $titleOfCourtesy[$random.Next(0, $titleOfCourtesy.Length)]
    $randomCity = $cities[$random.Next(0, $cities.Length)]
    $randomCountry = $countries[$random.Next(0, $countries.Length)]
    $randomAddress = $addresses[$random.Next(0, $addresses.Length)]

    # Generate phone number
    $phoneAreaCode = $random.Next(100, 999).ToString()  # 3 digits
    $phonePrefix = $random.Next(100, 999).ToString()    # 3 digits
    $phoneLineNumber = $random.Next(1000, 9999).ToString()  # 4 digits
    $randomPhone = "$phoneAreaCode-$phonePrefix-$phoneLineNumber"

    # Generating data for each row
    $row = @{
        "lastname"        = $randomLast
        "firstname"       = $randomFirst
        "title"           = $randomTitle
        "titleofcourtesy" = $randomCourtesy
        "birthdate"       = (Get-Date).AddYears(-$random.Next(25, 60))  # Random birthdate
        "hiredate"        = (Get-Date).AddYears(-$random.Next(1, 10))  # Random hiredate
        "address"         = $randomAddress
        "city"            = $randomCity
        "region"          = $null  # No region by default
        "postalcode"      = $random.Next(10000, 99999).ToString()
        "country"         = $randomCountry
        "phone"           = $randomPhone
        "mgrid"           = $null  # No manager for simplicity
    }

    Write-Host $row -ForegroundColor Yellow
    Write-Host "----------------------------------------------"
    $sqlData += $row
}


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
    #$transaction = $conn.BeginTransaction([System.Data.IsolationLevel]::ReadCommitted)
    Write-Host "Transaction started."

    # Create a SQL Command object outside of the loop (singleton pattern)
    $cmd = $conn.CreateCommand()
    $cmd.Transaction = $transaction  # Assign the transaction to the command

    $sqlData | ForEach-Object {
        try {
            $cmd.CommandText = $sqlTemplate

            # clear pervious parm values
            $cmd.Parameters.Clear()

            # Add parameters for each column
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@lastname", [System.Data.SqlDbType]::NVarChar, 20))).Value = $_["lastname"]
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@firstname", [System.Data.SqlDbType]::NVarChar, 10))).Value = $_["firstname"]
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@title", [System.Data.SqlDbType]::NVarChar, 30))).Value = $_["title"]
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@titleofcourtesy", [System.Data.SqlDbType]::NVarChar, 25))).Value = $_["titleofcourtesy"]
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@birthdate", [System.Data.SqlDbType]::DateTime))).Value = $_["birthdate"]
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@hiredate", [System.Data.SqlDbType]::DateTime))).Value = $_["hiredate"]
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@address", [System.Data.SqlDbType]::NVarChar, 60))).Value = $_["address"]
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@city", [System.Data.SqlDbType]::NVarChar, 15))).Value = $_["city"]

            # Check for null value in region and assign DBNull if necessary
            if ($null -eq $_["region"]) {
                $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@region", [System.Data.SqlDbType]::NVarChar, 15))).Value = [DBNull]::Value
            } else {
                $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@region", [System.Data.SqlDbType]::NVarChar, 15))).Value = $_["region"]
            }

            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@postalcode", [System.Data.SqlDbType]::NVarChar, 10))).Value = $_["postalcode"]
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@country", [System.Data.SqlDbType]::NVarChar, 15))).Value = $_["country"]
            $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@phone", [System.Data.SqlDbType]::NVarChar, 24))).Value = $_["phone"]

            # Check for null value in mgrid and assign DBNull if necessary
            if ($null -eq $_["mgrid"]) {
                $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@mgrid", [System.Data.SqlDbType]::Int))).Value = [DBNull]::Value
            } else {
                $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@mgrid", [System.Data.SqlDbType]::Int))).Value = $_["mgrid"]
            }

            # Execute the insert command
            $cmd.ExecuteNonQuery() | Out-Null
            Write-Host "Inserted record: `n`r$($_.Keys) $($_.Values)`n`r" -ForegroundColor Green

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
