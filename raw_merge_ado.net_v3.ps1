# Load necessary .NET assemblies
Add-Type -AssemblyName System.Data
Add-Type -Path "$PSScriptRoot\oracle.manageddataaccess.core.23.4.0\lib\netstandard2.1\Oracle.ManagedDataAccess.dll"

# Define Microsoft SQL Server connection parameters
$SQL_SERVER_INSTANCE = "Server\Instance"
$SQL_DATABASE = "Database"
$SQL_CONNECTION_STRING = "Server=$SQL_SERVER_INSTANCE;Database=$SQL_DATABASE;Integrated Security=True;TrustServerCertificate=True;"

# Oracle connection details
$ORACLE_USERNAME = ''
$ORACLE_PASSWORD = ''
$DB_SERVER_IP = ''
$PLUGGABLE_DB = ''

$tnsSource = @"
(DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = $DB_SERVER_IP)(PORT = 1521))
    (CONNECT_DATA =
        (SERVICE_NAME = $PLUGGABLE_DB)
    )
)
"@

$connectionString = 'User Id=' + $ORACLE_USERNAME + ';Password=' + $ORACLE_PASSWORD + ';Data Source=' + $tnsSource

# Create and open Oracle connection
$conn = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionString)

try {
    $conn.Open()

    # Define the query to retrieve data from Oracle
    $queryStatement = @"
SELECT e.employee_id, e.first_name, e.last_name, e.email, e.phone_number, e.hire_date, e.job_id, e.salary,
        e.commission_pct, e.department_id, d.department_name, j.job_title, j.min_salary, j.max_salary
FROM employees e
INNER JOIN departments d ON e.department_id = d.department_id
INNER JOIN jobs j ON e.job_id = j.job_id
"@

    # Create and configure the Oracle command
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $queryStatement
    $cmd.CommandTimeout = 3600 # Seconds

    # Create a data adapter and fill the DataTable
    $da = New-Object Oracle.ManagedDataAccess.Client.OracleDataAdapter($cmd)
    $resultSet = New-Object System.Data.DataTable
    [void]$da.Fill($resultSet)

} catch {
    Write-Error "Error fetching data from Oracle: $_"
} finally {
    if ($conn.State -eq [System.Data.ConnectionState]::Open) {
        $conn.Close()
    }
}

# Function to check if a row exists using ADO.NET
function RowExists {
    param (
        [int]$employee_id
    )
    
    $query = "SELECT COUNT(1) FROM dbo.combined_employees WHERE employee_id = @employee_id"
    $count = 0
    
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection($SQL_CONNECTION_STRING)
    
    try {
        $sqlConnection.Open()
        $sqlCommand = $sqlConnection.CreateCommand()
        $sqlCommand.CommandText = $query
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@employee_id", [Data.SqlDbType]::Decimal))).Value = $employee_id
        $count = $sqlCommand.ExecuteScalar()
    } catch {
        Write-Error "Error checking row existence: $_"
    } finally {
        if ($sqlConnection.State -eq [System.Data.ConnectionState]::Open) {
            $sqlConnection.Close()
        }
    }
    
    return $count -ne 0
}

# Helper function to handle DBNull for empty or null values
function Get-DBValue {
    param (
        [object]$value,
        [Type]$type
    )
    if ([string]::IsNullOrEmpty([string]$value)) {
        return [System.DBNull]::Value
    } else {
        return [Convert]::ChangeType($value, $type)
    }
}

# Iterate over each row in the resultSet and use parameterized queries
foreach ($row in $resultSet.Rows) {

    $row # debug

    [int]$employee_id = $row["employee_id"]

    # Open SQL Server connection
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection($SQL_CONNECTION_STRING)

    try {
        $sqlConnection.Open()
        
        # Begin a transaction with the desired isolation level - ReadCommitted (default)
        $transaction = $sqlConnection.BeginTransaction([System.Data.IsolationLevel]::ReadCommitted)

        # Merge query to update existing records or insert new records
        $mergeQuery = @"
MERGE INTO dbo.combined_employees AS target
USING (SELECT 
    @employee_id AS employee_id,
    @first_name AS first_name,
    @last_name AS last_name,
    @email AS email,
    @phone_number AS phone_number,
    @hire_date AS hire_date,
    @job_id AS job_id,
    @salary AS salary,
    @commission_pct AS commission_pct,
    @department_id AS department_id,
    @department_name AS department_name,
    @job_title AS job_title,
    @min_salary AS min_salary,
    @max_salary AS max_salary
) AS source
ON target.employee_id = source.employee_id
WHEN MATCHED THEN
    UPDATE SET 
        first_name = source.first_name,
        last_name = source.last_name,
        email = source.email,
        phone_number = source.phone_number,
        hire_date = source.hire_date,
        job_id = source.job_id,
        salary = source.salary,
        commission_pct = source.commission_pct,
        department_id = source.department_id,
        department_name = source.department_name,
        job_title = source.job_title,
        min_salary = source.min_salary,
        max_salary = source.max_salary
WHEN NOT MATCHED THEN
    INSERT (employee_id, first_name, last_name, email, phone_number, hire_date, job_id, salary, commission_pct, department_id, department_name, job_title, min_salary, max_salary)
    VALUES (source.employee_id, source.first_name, source.last_name, source.email, source.phone_number, source.hire_date, source.job_id, source.salary, source.commission_pct, source.department_id, source.department_name, source.job_title, source.min_salary, source.max_salary);
"@

        $sqlCommand = $sqlConnection.CreateCommand()
        $sqlCommand.CommandText = $mergeQuery
        $sqlCommand.Transaction = $transaction

        # Add parameters using the Get-DBValue function
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@employee_id", [Data.SqlDbType]::Decimal))).Value = $employee_id
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@first_name", [Data.SqlDbType]::NVarChar, 20))).Value = Get-DBValue $row["first_name"] ([System.String])
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@last_name", [Data.SqlDbType]::NVarChar, 25))).Value = Get-DBValue $row["last_name"] ([System.String])
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@email", [Data.SqlDbType]::NVarChar, 25))).Value = Get-DBValue $row["email"] ([System.String])
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@phone_number", [Data.SqlDbType]::NVarChar, 20))).Value = Get-DBValue $row["phone_number"] ([System.String])
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@hire_date", [Data.SqlDbType]::DateTime))).Value = Get-DBValue $row["hire_date"] ([System.Datetime])
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@job_id", [Data.SqlDbType]::NVarChar, 10))).Value = Get-DBValue $row["job_id"] ([System.String])
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@salary", [Data.SqlDbType]::Decimal))).Value = Get-DBValue $row["salary"] ([System.Decimal])
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@commission_pct", [Data.SqlDbType]::Decimal))).Value = Get-DBValue $row["commission_pct"] ([System.Decimal])
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@department_id", [Data.SqlDbType]::Decimal))).Value = Get-DBValue $row["department_id"] ([System.Decimal])
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@department_name", [Data.SqlDbType]::NVarChar, 30))).Value = Get-DBValue $row["department_name"] ([System.String])
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@job_title", [Data.SqlDbType]::NVarChar, 35))).Value = Get-DBValue $row["job_title"] ([System.String])
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@min_salary", [Data.SqlDbType]::Decimal))).Value = Get-DBValue $row["min_salary"] ([System.Decimal])
        $sqlCommand.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@max_salary", [Data.SqlDbType]::Decimal))).Value = Get-DBValue $row["max_salary"] ([System.Decimal])

        Write-Host "Executing MERGE Statement for Employee ID: $employee_id"
        $sqlCommand.ExecuteNonQuery() | Out-Null
        
        # Commit the transaction
        $transaction.Commit()

    } catch {
        Write-Error "Error executing SQL command for Employee ID $employee_id : $_"
        # Rollback the transaction in case of error
        $transaction.Rollback()
    } finally {
        if ($sqlConnection.State -eq [System.Data.ConnectionState]::Open) {
            $sqlConnection.Close()
        }
    }
}
