# Install the ImportExcel module if not already installed
# Install-Module -Name ImportExcel -Force -Scope CurrentUser

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
    $row = [PSCustomObject]@{
        "LastName"        = $randomLast
        "FirstName"       = $randomFirst
        "Title"           = $randomTitle
        "TitleOfCourtesy" = $randomCourtesy
        "BirthDate"       = (Get-Date).AddYears(-$random.Next(25, 60)).ToString("yyyy-MM-dd")  # Random birthdate
        "HireDate"        = (Get-Date).AddYears(-$random.Next(1, 10)).ToString("yyyy-MM-dd")   # Random hiredate
        "Address"         = $randomAddress
        "City"            = $randomCity
        "Region"          = $null  # No region by default
        "PostalCode"      = $random.Next(10000, 99999).ToString()
        "Country"         = $randomCountry
        "Phone"           = $randomPhone
        "MgrID"           = $null  # No manager for simplicity
    }

    $sqlData += $row
}

Write-Host "Writing Test Data to Excel" -ForegroundColor Green

# Export to Excel
#$sqlData | Export-Excel -Path "$PSScriptRoot\TestData.xlsx" -AutoSize -Verbose

$sqlData | ForEach-Object {
    # Print each row to the console
    Write-Host "Processing row: $($_.FirstName) $($_.LastName)" -ForegroundColor Yellow
    $_  # Return the object back to the pipeline for Export-Excel
} | Export-Excel -Path "$PSScriptRoot\TestData.xlsx" -AutoSize

Write-Host "Test Data written to C:\temp\TestData.xlsx" -ForegroundColor Cyan
