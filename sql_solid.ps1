# Load the IDatabase module
#Import-Module -Name "$PSScriptRoot\IDatabaseModule.psm1"


class SqlServerDatabase : DatabaseNamespace.IDatabase {
    [string]$serverName
    [string]$databaseName
    [System.Data.SqlClient.SqlConnection]$conn
    [System.Data.SqlClient.SqlCommand]$cmd
    [array]$dataList

    SqlServerDatabase([string]$serverName, [string]$databaseName, [array]$dataList) {
        $this.serverName = $serverName
        $this.databaseName = $databaseName
        $this.dataList = $dataList
        $this.conn = New-Object System.Data.SqlClient.SqlConnection
        $this.conn.ConnectionString = "Server=$($this.serverName);Database=$($this.databaseName);Integrated Security=True;"
        $this.cmd = $this.conn.CreateCommand()
    }

    [void] OpenConnection() {
        try {
            $this.conn.Open()
        } catch {
            throw $_
        }
    }

    [void] CloseConnection() {
        try {
            if ($this.conn.State -eq 'Open') {
                $this.conn.Close()
            }
        } catch {
            throw $_
        }
    }

    [void] ExecuteDml() {
        $sqlCommand = @"
            INSERT INTO [HR].[Employees_copy]
            (lastname, firstname, title, titleofcourtesy, birthdate, hiredate, address, city, region, postalcode, country, phone, mgrid) 
            VALUES (@lastname, @firstname, @title, @titleofcourtesy, @birthdate, @hiredate, @address, @city, @region, @postalcode, @country, @phone, @mgrid)
"@
        foreach ($data in $this.dataList) {
            try {
                $this.cmd.CommandText = $sqlCommand
                $this.cmd.Parameters.Clear()
    
                # Add parameters for each column
                $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@lastname", [System.Data.SqlDbType]::NVarChar, 20))).Value = $data["lastname"]
                $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@firstname", [System.Data.SqlDbType]::NVarChar, 10))).Value = $data["firstname"]
                $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@title", [System.Data.SqlDbType]::NVarChar, 30))).Value = $data["title"]
                $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@titleofcourtesy", [System.Data.SqlDbType]::NVarChar, 25))).Value = $data["titleofcourtesy"]
                $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@birthdate", [System.Data.SqlDbType]::DateTime))).Value = $data["birthdate"]
                $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@hiredate", [System.Data.SqlDbType]::DateTime))).Value = $data["hiredate"]
                $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@address", [System.Data.SqlDbType]::NVarChar, 60))).Value = $data["address"]
                $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@city", [System.Data.SqlDbType]::NVarChar, 15))).Value = $data["city"]
    
                # Check for null value in region and assign DBNull if necessary
                if ($null -eq $data["region"]) {
                    $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@region", [System.Data.SqlDbType]::NVarChar, 15))).Value = [DBNull]::Value
                } else {
                    $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@region", [System.Data.SqlDbType]::NVarChar, 15))).Value = $data["region"]
                }
    
                $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@postalcode", [System.Data.SqlDbType]::NVarChar, 10))).Value = $data["postalcode"]
                $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@country", [System.Data.SqlDbType]::NVarChar, 15))).Value = $data["country"]
                $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@phone", [System.Data.SqlDbType]::NVarChar, 24))).Value = $data["phone"]
    
                # Check for null value in mgrid and assign DBNull if necessary
                if ($null -eq $data["mgrid"]) {
                    $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@mgrid", [System.Data.SqlDbType]::Int))).Value = [DBNull]::Value
                } else {
                    $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@mgrid", [System.Data.SqlDbType]::Int))).Value = $data["mgrid"]
                }
    
                # Execute the insert command for this data item
                $this.cmd.ExecuteNonQuery()
            } catch {
                throw $_
            }
        }
    }       
}


function Invoke-DatabaseOperation {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [DatabaseNamespace.IDatabase]$DatabaseObject
    )

    begin {

        Write-Host "Starting the processing of database objects..."
    }

    process {
        if ($DatabaseObject -is [DatabaseNamespace.IDatabase]) {
            try {
                # Open the database connection
                $DatabaseObject.OpenConnection()
                Write-Host "Connection opened successfully."

                # Execute the DML operation on the database
                $DatabaseObject.ExecuteDml()

                # Update statistics based on the number of successful operations
                Write-Host "Operation executed successfully for the provided data list." -ForegroundColor Green

            } catch {
                Write-Host "Error during database operation: $($_.Exception.Message)" -ForegroundColor Red
                throw  # Propagate the error further if necessary
            }
        } else {
            Write-Host "Invalid object in pipeline. Expected object to implement Database interface." -ForegroundColor Red
        }
    }

    end {
        try {
            # Close the database connection
            $DatabaseObject.CloseConnection()
            Write-Host "Connection closed successfully."
        } catch {
            Write-Host "Error closing connection: $($_.Exception.Message)" -ForegroundColor Red
        }

        Write-Host "Finished processing all database objects."
    }
}


######################## Main Program #############################


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

for ($i = 1; $i -le 500; $i++) {
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

    $sqlData += $row
    
    Write-Host $row
    Write-Host "---------------------------------------------------"
}

$sqlServerObj = [SqlServerDatabase]::new("localhost\DB2016", "TSQL2012", $sqlData)


$sqlServerObj | Invoke-DatabaseOperation

