# **** Only works with PowerShell 7.4 because of oracle assembly ****

# CREATE PROCEDURE sp_updateCombinedEmployees
#     @employee_id NUMERIC(6,0),
#     @first_name VARCHAR(20),
#     @last_name VARCHAR(25),
#     @email VARCHAR(25),
#     @phone_number VARCHAR(20),
#     @hire_date DATE,
#     @job_id VARCHAR(10),
#     @salary NUMERIC(8, 2),
#     @commission_pct NUMERIC(2, 2),
#     @department_id NUMERIC(4,0),
#     @department_name VARCHAR(30),
#     @job_title VARCHAR(35),
#     @min_salary NUMERIC(6, 0),
#     @max_salary NUMERIC(6, 0)
# AS
# BEGIN
#     -- Handle NULL or empty string values
#     IF @first_name = '' SET @first_name = NULL
#     IF @last_name = '' SET @last_name = NULL
#     IF @email = '' SET @email = NULL
#     IF @phone_number = '' SET @phone_number = NULL
#     IF @job_id = '' SET @job_id = NULL
#     IF @department_name = '' SET @department_name = NULL
#     IF @job_title = '' SET @job_title = NULL

#     -- Update the employee record
#     UPDATE dbo.combined_employees
#     SET 
#         first_name = @first_name,
#         last_name = @last_name,
#         email = @email,
#         phone_number = @phone_number,
#         hire_date = @hire_date,
#         job_id = @job_id,
#         salary = @salary,
#         commission_pct = @commission_pct,
#         department_id = @department_id,
#         department_name = @department_name,
#         job_title = @job_title,
#         min_salary = @min_salary,
#         max_salary = @max_salary
#     WHERE employee_id = @employee_id
# END

# CREATE PROCEDURE sp_insertCombinedEmployees
#     @employee_id NUMERIC(6,0),
#     @first_name VARCHAR(20),
#     @last_name VARCHAR(25),
#     @email VARCHAR(25),
#     @phone_number VARCHAR(20),
#     @hire_date DATE,
#     @job_id VARCHAR(10),
#     @salary NUMERIC(8, 2),
#     @commission_pct NUMERIC(2, 2),
#     @department_id NUMERIC(4,0),
#     @department_name VARCHAR(30),
#     @job_title VARCHAR(35),
#     @min_salary NUMERIC(6, 0),
#     @max_salary NUMERIC(6, 0)
# AS
# BEGIN
#     -- Handle NULL or empty string values
#     IF @first_name = '' SET @first_name = NULL
#     IF @last_name = '' SET @last_name = NULL
#     IF @email = '' SET @email = NULL
#     IF @phone_number = '' SET @phone_number = NULL
#     IF @job_id = '' SET @job_id = NULL
#     IF @department_name = '' SET @department_name = NULL
#     IF @job_title = '' SET @job_title = NULL

#     -- Insert the new employee record
#     INSERT INTO dbo.combined_employees (
#         employee_id,
#         first_name,
#         last_name,
#         email,
#         phone_number,
#         hire_date,
#         job_id,
#         salary,
#         commission_pct,
#         department_id,
#         department_name,
#         job_title,
#         min_salary,
#         max_salary
#     )
#     VALUES (
#         @employee_id,
#         @first_name,
#         @last_name,
#         @email,
#         @phone_number,
#         @hire_date,
#         @job_id,
#         @salary,
#         @commission_pct,
#         @department_id,
#         @department_name,
#         @job_title,
#         @min_salary,
#         @max_salary
#     )
# END

# Load necessary .NET assemblies
Add-Type -AssemblyName System.Data
Add-Type -Path "$PSScriptRoot\oracle.manageddataaccess.core.23.4.0\lib\netstandard2.1\Oracle.ManagedDataAccess.dll"

# Define SQL Server connection parameters
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
    $exists = RowExists -employee_id $employee_id

    # Open SQL Server connection
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection($SQL_CONNECTION_STRING)

    try {
        $sqlConnection.Open()
        if ($exists) {

            # Update the existing record
            $updateQuery = @"
UPDATE dbo.combined_employees
SET 
    first_name = @first_name,
    last_name = @last_name,
    email = @email,
    phone_number = @phone_number,
    hire_date = @hire_date,
    job_id = @job_id,
    salary = @salary,
    commission_pct = @commission_pct,
    department_id = @department_id,
    department_name = @department_name,
    job_title = @job_title,
    min_salary = @min_salary,
    max_salary = @max_salary
WHERE employee_id = @employee_id
"@

            $sqlCommand = $sqlConnection.CreateCommand()
            $sqlCommand.CommandText = $updateQuery

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

            Write-Host "Executing Update Statement for Employee ID: $employee_id"
            $sqlCommand.ExecuteNonQuery() | Out-Null
        } 
        else {
            # Insert a new record
            $insertQuery = @"
INSERT INTO dbo.combined_employees (
    employee_id,
    first_name,
    last_name,
    email,
    phone_number,
    hire_date,
    job_id,
    salary,
    commission_pct,
    department_id,
    department_name,
    job_title,
    min_salary,
    max_salary
)
VALUES (
    @employee_id,
    @first_name,
    @last_name,
    @email,
    @phone_number,
    @hire_date,
    @job_id,
    @salary,
    @commission_pct,
    @department_id,
    @department_name,
    @job_title,
    @min_salary,
    @max_salary
)
"@

            $sqlCommand = $sqlConnection.CreateCommand()
            $sqlCommand.CommandText = $insertQuery

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
            
            Write-Host "Executing Insert Statement for Employee ID: $employee_id"
            $sqlCommand.ExecuteNonQuery() | Out-Null
        }
    } catch {
        Write-Error "Error executing SQL command for Employee ID $employee_id : $_"
    } finally {
        if ($sqlConnection.State -eq [System.Data.ConnectionState]::Open) {
            $sqlConnection.Close()
        }
    }
}
