# base class for database operations
class Database {
    [void] OpenConnection() { throw "NotImplementedException" }
    [void] CloseConnection() { throw "NotImplementedException" }
    [void] InsertData([hashtable]$data) { throw "NotImplementedException" }
}

# concrete SQL Server implementation
class SqlServerDatabase : Database {
    [string]$serverName
    [string]$databaseName
    [System.Data.SqlClient.SqlConnection]$conn
    [System.Data.SqlClient.SqlCommand]$cmd

    SqlServerDatabase([string]$serverName, [string]$databaseName) {
        $this.serverName = $serverName
        $this.databaseName = $databaseName
        $this.conn = New-Object System.Data.SqlClient.SqlConnection
        $this.conn.ConnectionString = "Server=$($this.serverName);Database=$($this.databaseName);Integrated Security=True;"
        $this.cmd = $this.conn.CreateCommand()
    }

    [void] OpenConnection() {
        try {
            $this.conn.Open()
            Write-Host "SQL Server connection opened successfully."
        } catch {
            Write-Host "Error opening connection: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    [void] CloseConnection() {
        if ($this.conn.State -eq 'Open') {
            $this.conn.Close()
            Write-Host "SQL Server connection closed."
        }
    }

    [void] InsertData([hashtable]$data) {
        try {
            $sqlTemplate = "INSERT INTO your_table (Name, Age) VALUES (@Name, @Age)"
            $this.cmd.CommandText = $sqlTemplate
            $this.cmd.Parameters.Clear()
            $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Name", $data.Name)))
            $this.cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Age", $data.Age)))
            $this.cmd.ExecuteNonQuery()
            Write-Host "Inserted record for Name: $($data.Name)" -ForegroundColor Green
        } catch {
            Write-Host "Failed to insert record for Name: $($data.Name)" -ForegroundColor Red
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Main logic for handling database insertion
function InsertDataToDatabase {
    param(
        [Database]$database,  # Base class type
        [array]$dataList
    )

    try {
        # Open the database connection
        $database.OpenConnection()

        # Loop through each data item and insert it
        $dataList | ForEach-Object {
            $database.InsertData($_)
        }
    } finally {
        # Close the database connection
        $database.CloseConnection()
    }
}

# Example data to insert
$data = @(
    @{Name="John"; Age=30},
    @{Name="Jane"; Age=25},
    @{Name="Doe"; Age=35}
)

# Use SQL Server implementation
$sqlServer = [SqlServerDatabase]::new("your_sql_server", "your_database")

# Insert data into SQL Server
InsertDataToDatabase -database $sqlServer -dataList $data
