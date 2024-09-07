# Connection details
$serverName = "your_sql_server"
$databaseName = "your_database"
$connString = "Server=$serverName;Database=$databaseName;Integrated Security=True;"
$sqlTemplate = "INSERT INTO your_table (Name, Age) VALUES (@Name, @Age)"

  # Data to insert, replace with your actual data
  $data = @(
    @{Name="John"; Age=30},
    @{Name="Jane"; Age=25},
    @{Name="Doe"; Age=35}
)

# Create a connection object
$conn = New-Object System.Data.SqlClient.SqlConnection
$conn.ConnectionString = $connString

try {
    # Open the connection
    $conn.Open()
    Write-Host "Connection opened successfully."

    # Create a SQL Command object outside of the loop (singleton pattern)
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $sqlTemplate
    
    # Create the parameters once and reuse them
    $paramName = $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Name", [System.Data.SqlDbType]::NVarChar)))
    $paramAge = $cmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@Age", [System.Data.SqlDbType]::Int)))

    # Insert data for each object in $data array
    $data | ForEach-Object {
        try {
            # Set the values for the parameters
            $paramName.Value = $_.Name
            $paramAge.Value = $_.Age

            # Execute the insert command
            $cmd.ExecuteNonQuery()
            Write-Host "Inserted record for Name: $($_.Name)" -ForegroundColor Green
        } catch {
            # Handle errors during individual insert execution
            Write-Host "Failed to insert record for Name: $($_.Name)" -ForegroundColor Red
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
} catch {
    # Handle connection-level or overall errors
    Write-Host "Failed to open connection or run query." -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # Ensure the connection is closed
    if ($conn.State -eq 'Open') {
        $conn.Close()
        Write-Host "Connection closed."
    }
}
