# Install the ImportExcel module if not already installed
# Install-Module -Name ImportExcel -Force -Scope CurrentUser

# Generate test data
$random = New-Object System.Random
$empIdCounter = 1  # Start empid counter

$sqlData = @()
$generatedEmployees = @()

# Define hierarchical levels for manager mapping with realistic limits
$hierarchy = @{
    "President" = @()
    "Chief Executive Officer (CEO)" = @("President")
    "Senior Vice President" = @("Chief Executive Officer (CEO)")
    "Vice President" = @("Senior Vice President")
    "Director" = @("Vice President")
    "Senior Director" = @("Vice President")
    "Manager" = @("Director", "Senior Director")
    "Team Lead" = @("Manager")
    "Technician" = @("Team Lead", "Manager")
    "Junior Technician" = @("Manager", "Team Lead")
    "Senior Technician" = @("Manager")
    "Clerk" = @("Manager")
    "Junior Clerk" = @("Manager")
    "Coordinator" = @("Manager")
}

# Titles to be used with their corresponding counts for real-world proportions
$positionsWithCount = @(
    @{ title = "President"; count = 1 },
    @{ title = "Chief Executive Officer (CEO)"; count = 1 },
    @{ title = "Senior Vice President"; count = 2 },
    @{ title = "Vice President"; count = 3 },
    @{ title = "Director"; count = 5 },
    @{ title = "Senior Director"; count = 5 },
    @{ title = "Manager"; count = 15 },
    @{ title = "Team Lead"; count = 20 },
    @{ title = "Technician"; count = 30 },
    @{ title = "Junior Technician"; count = 20 },
    @{ title = "Senior Technician"; count = 10 },
    @{ title = "Clerk"; count = 25 },
    @{ title = "Junior Clerk"; count = 20 },
    @{ title = "Coordinator"; count = 10 }
)


# Employees data
$employees = @(
        @{ lastname = "Doe"; firstname = "John"; titleofcourtesy = "Mr."; birthdate = "1980-05-12"; hiredate = "2010-06-15" },
        @{ lastname = "Smith"; firstname = "Jane"; titleofcourtesy = "Ms."; birthdate = "1985-07-18"; hiredate = "2012-03-10" },
        @{ lastname = "Brown"; firstname = "Emily"; titleofcourtesy = "Mrs."; birthdate = "1978-09-23"; hiredate = "2011-01-22" },
        @{ lastname = "Wilson"; firstname = "Jake"; titleofcourtesy = "Mr."; birthdate = "1982-11-30"; hiredate = "2014-07-11" },
        @{ lastname = "Davis"; firstname = "Laura"; titleofcourtesy = "Ms."; birthdate = "1990-04-14"; hiredate = "2016-10-18" },
        @{ lastname = "Johnson"; firstname = "Tom"; titleofcourtesy = "Mr."; birthdate = "1975-02-09"; hiredate = "2008-12-01" },
        @{ lastname = "Anderson"; firstname = "Anna"; titleofcourtesy = "Mrs."; birthdate = "1987-06-21"; hiredate = "2015-08-23" },
        @{ lastname = "Taylor"; firstname = "David"; titleofcourtesy = "Mr."; birthdate = "1981-01-27"; hiredate = "2009-04-17" },
        @{ lastname = "Moore"; firstname = "Mia"; titleofcourtesy = "Ms."; birthdate = "1992-03-19"; hiredate = "2017-11-05" },
        @{ lastname = "Clark"; firstname = "Russell"; titleofcourtesy = "Mr."; birthdate = "1979-08-15"; hiredate = "2013-05-07" },
        @{ lastname = "Miller"; firstname = "Sarah"; titleofcourtesy = "Ms."; birthdate = "1983-12-11"; hiredate = "2014-01-29" },
        @{ lastname = "Adams"; firstname = "Paul"; titleofcourtesy = "Mr."; birthdate = "1976-10-03"; hiredate = "2010-09-12" },
        @{ lastname = "Baker"; firstname = "Zoe"; titleofcourtesy = "Mrs."; birthdate = "1989-07-22"; hiredate = "2015-06-01" },
        @{ lastname = "Carter"; firstname = "Chris"; titleofcourtesy = "Mr."; birthdate = "1985-02-26"; hiredate = "2012-12-15" },
        @{ lastname = "Evans"; firstname = "Liam"; titleofcourtesy = "Mr."; birthdate = "1980-11-18"; hiredate = "2009-03-20" },
        @{ lastname = "Green"; firstname = "Sophia"; titleofcourtesy = "Ms."; birthdate = "1991-06-08"; hiredate = "2016-04-11" },
        @{ lastname = "Hall"; firstname = "Oliver"; titleofcourtesy = "Mr."; birthdate = "1978-12-30"; hiredate = "2007-10-09" },
        @{ lastname = "King"; firstname = "Charlotte"; titleofcourtesy = "Mrs."; birthdate = "1982-08-25"; hiredate = "2013-09-03" },
        @{ lastname = "Wright"; firstname = "Benjamin"; titleofcourtesy = "Mr."; birthdate = "1977-05-16"; hiredate = "2010-02-11" },
        @{ lastname = "Hill"; firstname = "Isabella"; titleofcourtesy = "Ms."; birthdate = "1988-03-04"; hiredate = "2015-07-09" },
        @{ lastname = "Cooper"; firstname = "William"; titleofcourtesy = "Mr."; birthdate = "1984-09-07"; hiredate = "2011-11-30" },
        @{ lastname = "Parker"; firstname = "Grace"; titleofcourtesy = "Mrs."; birthdate = "1979-01-12"; hiredate = "2010-06-24" },
        @{ lastname = "Collins"; firstname = "Lucas"; titleofcourtesy = "Mr."; birthdate = "1987-05-03"; hiredate = "2014-03-17" },
        @{ lastname = "Cook"; firstname = "Amelia"; titleofcourtesy = "Ms."; birthdate = "1990-10-29"; hiredate = "2016-09-18" },
        @{ lastname = "Bell"; firstname = "James"; titleofcourtesy = "Mr."; birthdate = "1981-12-16"; hiredate = "2009-07-14" },
        @{ lastname = "Bailey"; firstname = "Elijah"; titleofcourtesy = "Mr."; birthdate = "1982-06-11"; hiredate = "2013-01-22" },
        @{ lastname = "Rivera"; firstname = "Harper"; titleofcourtesy = "Ms."; birthdate = "1991-04-20"; hiredate = "2015-05-15" },
        @{ lastname = "Gonzalez"; firstname = "Mason"; titleofcourtesy = "Mr."; birthdate = "1985-07-07"; hiredate = "2011-03-02" },
        @{ lastname = "Hughes"; firstname = "Evelyn"; titleofcourtesy = "Mrs."; birthdate = "1979-02-05"; hiredate = "2008-11-21" },
        @{ lastname = "Flores"; firstname = "Logan"; titleofcourtesy = "Mr."; birthdate = "1986-05-19"; hiredate = "2012-04-30" },
        @{ lastname = "Morgan"; firstname = "Ava"; titleofcourtesy = "Ms."; birthdate = "1992-11-13"; hiredate = "2017-10-08" },
        @{ lastname = "Lee"; firstname = "Alexander"; titleofcourtesy = "Mr."; birthdate = "1983-03-21"; hiredate = "2009-09-05" },
        @{ lastname = "Murphy"; firstname = "Eleanor"; titleofcourtesy = "Mrs."; birthdate = "1980-08-29"; hiredate = "2010-05-10" },
        @{ lastname = "Martinez"; firstname = "Daniel"; titleofcourtesy = "Mr."; birthdate = "1988-01-03"; hiredate = "2015-12-02" },
        @{ lastname = "Martin"; firstname = "Chloe"; titleofcourtesy = "Ms."; birthdate = "1991-09-25"; hiredate = "2016-11-19" },
        @{ lastname = "Walker"; firstname = "Henry"; titleofcourtesy = "Mr."; birthdate = "1978-10-10"; hiredate = "2007-02-14" },
        @{ lastname = "Perez"; firstname = "Victoria"; titleofcourtesy = "Ms."; birthdate = "1993-06-15"; hiredate = "2018-03-28" },
        @{ lastname = "Robinson"; firstname = "Michael"; titleofcourtesy = "Mr."; birthdate = "1984-05-24"; hiredate = "2011-09-22" },
        @{ lastname = "Turner"; firstname = "Stella"; titleofcourtesy = "Mrs."; birthdate = "1977-07-19"; hiredate = "2006-12-05" },
        @{ lastname = "Campbell"; firstname = "Sebastian"; titleofcourtesy = "Mr."; birthdate = "1981-11-12"; hiredate = "2008-08-15" },
        @{ lastname = "Sanders"; firstname = "Lily"; titleofcourtesy = "Ms."; birthdate = "1994-02-27"; hiredate = "2018-09-14" },
        @{ lastname = "Reed"; firstname = "Ethan"; titleofcourtesy = "Mr."; birthdate = "1985-12-21"; hiredate = "2012-05-11" },
        @{ lastname = "Foster"; firstname = "Penelope"; titleofcourtesy = "Ms."; birthdate = "1990-08-06"; hiredate = "2017-07-04" },
        @{ lastname = "Powell"; firstname = "Matthew"; titleofcourtesy = "Mr."; birthdate = "1979-04-10"; hiredate = "2009-01-31" },
        @{ lastname = "Howard"; firstname = "Aria"; titleofcourtesy = "Ms."; birthdate = "1991-12-09"; hiredate = "2016-06-27" },
        @{ lastname = "Ward"; firstname = "Owen"; titleofcourtesy = "Mr."; birthdate = "1987-01-25"; hiredate = "2013-08-20" }
)

# Locations data
$locations = @(
    @{ address = "123 Main St"; city = "New York"; state = "NY"; region = "Northeast"; postalcode = "10001"; country = "USA"; phone = "(212) 555-1234" },
    @{ address = "456 Oak St"; city = "Los Angeles"; state = "CA"; region = "West"; postalcode = "90001"; country = "USA"; phone = "(310) 555-5678" },
    @{ address = "789 Pine St"; city = "Chicago"; state = "IL"; region = "Midwest"; postalcode = "60601"; country = "USA"; phone = "(312) 555-7890" },
    @{ address = "101 Maple Ave"; city = "Houston"; state = "TX"; region = "South"; postalcode = "77001"; country = "USA"; phone = "(713) 555-0101" },
    @{ address = "202 Elm St"; city = "Phoenix"; state = "AZ"; region = "West"; postalcode = "85001"; country = "USA"; phone = "(602) 555-1212" },
    @{ address = "5678 Old Redmond Rd."; city = "Redmond"; state = "WA"; region = "West"; postalcode = "98052"; country = "USA"; phone = "(425) 555-3456" },
    @{ address = "2345 Moss Bay Blvd."; city = "Kirkland"; state = "WA"; region = "West"; postalcode = "98033"; country = "USA"; phone = "(425) 555-7891" },
    @{ address = "7890 - 20th Ave. E."; city = "Seattle"; state = "WA"; region = "West"; postalcode = "98101"; country = "USA"; phone = "(206) 555-4567" },
    @{ address = "3456 Coventry House, Miner Rd."; city = "London"; state = $null; region = $null; postalcode = "EC1A 1BB"; country = "UK"; phone = "+44 20 7123 4567" },
    @{ address = "8901 Garrett Hill"; city = "London"; state = $null; region = $null; postalcode = "SW1A 1AA"; country = "UK"; phone = "+44 20 7984 7890" },
    @{ address = "1600 Amphitheatre Parkway"; city = "Mountain View"; state = "CA"; region = "West"; postalcode = "94043"; country = "USA"; phone = "(650) 555-7890" },
    @{ address = "1 Infinite Loop"; city = "Cupertino"; state = "CA"; region = "West"; postalcode = "95014"; country = "USA"; phone = "(408) 555-1234" },
    @{ address = "350 5th Ave"; city = "New York"; state = "NY"; region = "Northeast"; postalcode = "10118"; country = "USA"; phone = "(212) 555-9876" },
    @{ address = "500 S. Buena Vista St."; city = "Burbank"; state = "CA"; region = "West"; postalcode = "91521"; country = "USA"; phone = "(818) 555-6543" },
    @{ address = "111 8th Ave"; city = "New York"; state = "NY"; region = "Northeast"; postalcode = "10011"; country = "USA"; phone = "(212) 555-1122" },
    @{ address = "555 Pennsylvania Ave NW"; city = "Washington"; state = "DC"; region = "Mid-Atlantic"; postalcode = "20001"; country = "USA"; phone = "(202) 555-2233" },
    @{ address = "1000 Wilshire Blvd"; city = "Los Angeles"; state = "CA"; region = "West"; postalcode = "90017"; country = "USA"; phone = "(213) 555-3344" },
    @{ address = "1 Queen St"; city = "Auckland"; state = $null; region = "Auckland"; postalcode = "1010"; country = "New Zealand"; phone = "+64 9 555 5678" },
    @{ address = "10 Downing St"; city = "London"; state = $null; region = $null; postalcode = "SW1A 2AA"; country = "UK"; phone = "+44 20 555 6789" },
    @{ address = "4 Rue de la Paix"; city = "Paris"; state = $null; region = $null; postalcode = "75002"; country = "France"; phone = "+33 1 555 4321" },
    @{ address = "221B Baker St"; city = "London"; state = $null; region = $null; postalcode = "NW1 6XE"; country = "UK"; phone = "+44 20 555 1212" },
    @{ address = "1600 Pennsylvania Ave NW"; city = "Washington"; state = "DC"; region = "Mid-Atlantic"; postalcode = "20500"; country = "USA"; phone = "(202) 555-6789" },
    @{ address = "100 George St"; city = "Sydney"; state = "NSW"; region = $null; postalcode = "2000"; country = "Australia"; phone = "+61 2 555 3456" },
    @{ address = "121 King St"; city = "Melbourne"; state = "VIC"; region = $null; postalcode = "3000"; country = "Australia"; phone = "+61 3 555 7891" },
    @{ address = "300 N. LaSalle St"; city = "Chicago"; state = "IL"; region = "Midwest"; postalcode = "60654"; country = "USA"; phone = "(312) 555-6543" },
    @{ address = "75 State St"; city = "Boston"; state = "MA"; region = "Northeast"; postalcode = "02109"; country = "USA"; phone = "(617) 555-0987" },
    @{ address = "1 Microsoft Way"; city = "Redmond"; state = "WA"; region = "West"; postalcode = "98052"; country = "USA"; phone = "(425) 555-3456" },
    @{ address = "500 Terry Francois Blvd"; city = "San Francisco"; state = "CA"; region = "West"; postalcode = "94158"; country = "USA"; phone = "(415) 555-4567" },
    @{ address = "1 Hacker Way"; city = "Menlo Park"; state = "CA"; region = "West"; postalcode = "94025"; country = "USA"; phone = "(650) 555-5678" },
    @{ address = "200 Vesey St"; city = "New York"; state = "NY"; region = "Northeast"; postalcode = "10281"; country = "USA"; phone = "(212) 555-6543" },
    @{ address = "1-1-1 Marunouchi"; city = "Tokyo"; state = $null; region = "Kanto"; postalcode = "100-0005"; country = "Japan"; phone = "+81 3 555 6789" },
    @{ address = "350 Ellis St"; city = "Mountain View"; state = "CA"; region = "West"; postalcode = "94043"; country = "USA"; phone = "(650) 555-1234" },
    @{ address = "200 Park Ave"; city = "New York"; state = "NY"; region = "Northeast"; postalcode = "10166"; country = "USA"; phone = "(212) 555-4567" },
    @{ address = "303 Collins St"; city = "Melbourne"; state = "VIC"; region = $null; postalcode = "3000"; country = "Australia"; phone = "+61 3 555 5678" },
    @{ address = "77 Water St"; city = "New York"; state = "NY"; region = "Northeast"; postalcode = "10005"; country = "USA"; phone = "(212) 555-7890" },
    @{ address = "333 Wacker Dr"; city = "Chicago"; state = "IL"; region = "Midwest"; postalcode = "60606"; country = "USA"; phone = "(312) 555-6543" },
    @{ address = "1 London Bridge St"; city = "London"; state = $null; region = $null; postalcode = "SE1 9GF"; country = "UK"; phone = "+44 20 555 9876" },
    @{ address = "221 Yonge St"; city = "Toronto"; state = "ON"; region = "Central Canada"; postalcode = "M5B 1N1"; country = "Canada"; phone = "(416) 555-4321" },
    @{ address = "10 St. Mary Axe"; city = "London"; state = $null; region = $null; postalcode = "EC3A 8BF"; country = "UK"; phone = "+44 20 555 3456" },
    @{ address = "500 Wellington St"; city = "Toronto"; state = "ON"; region = "Central Canada"; postalcode = "M5V 3P6"; country = "Canada"; phone = "(416) 555-6789" },
    @{ address = "101 Collins St"; city = "Melbourne"; state = "VIC"; region = $null; postalcode = "3000"; country = "Australia"; phone = "+61 3 555 0987" },
    @{ address = "800 Rue Sherbrooke"; city = "Montreal"; state = "QC"; region = "Eastern Canada"; postalcode = "H3A 2K6"; country = "Canada"; phone = "(514) 555-6543" },
    @{ address = "151 Front St"; city = "Toronto"; state = "ON"; region = "Central Canada"; postalcode = "M5J 2N1"; country = "Canada"; phone = "(416) 555-7890" },
    @{ address = "222 Exhibition St"; city = "Melbourne"; state = "VIC"; region = $null; postalcode = "3000"; country = "Australia"; phone = "+61 3 555 9876" },
    @{ address = "50 Grosvenor Hill"; city = "London"; state = $null; region = $null; postalcode = "W1K 3JH"; country = "UK"; phone = "+44 20 555 8765" },
    @{ address = "6101 Long Prairie Rd"; city = "Flower Mound"; state = "TX"; region = "South"; postalcode = "75028"; country = "USA"; phone = "(972) 555-1234" },
    @{ address = "555 California St"; city = "San Francisco"; state = "CA"; region = "West"; postalcode = "94104"; country = "USA"; phone = "(415) 555-5678" },
    @{ address = "101 California St"; city = "San Francisco"; state = "CA"; region = "West"; postalcode = "94111"; country = "USA"; phone = "(415) 555-3456" },
    @{ address = "20 Queens St"; city = "Brisbane"; state = "QLD"; region = $null; postalcode = "4000"; country = "Australia"; phone = "+61 7 555 7891" },
    @{ address = "30 St Mary Axe"; city = "London"; state = $null; region = $null; postalcode = "EC3A 8BF"; country = "UK"; phone = "+44 20 555 4567" },
    @{ address = "88 Phillip St"; city = "Sydney"; state = "NSW"; region = $null; postalcode = "2000"; country = "Australia"; phone = "+61 2 555 7890" },
    @{ address = "31 Rue Cambon"; city = "Paris"; state = $null; region = $null; postalcode = "75001"; country = "France"; phone = "+33 1 555 6789" },
    @{ address = "333 George St"; city = "Sydney"; state = "NSW"; region = $null; postalcode = "2000"; country = "Australia"; phone = "+61 2 555 0987" },
    @{ address = "350 Madison Ave"; city = "New York"; state = "NY"; region = "Northeast"; postalcode = "10017"; country = "USA"; phone = "(212) 555-5678" },
    @{ address = "300 Montgomery St"; city = "San Francisco"; state = "CA"; region = "West"; postalcode = "94104"; country = "USA"; phone = "(415) 555-7890" },
    @{ address = "2121 Avenue of the Stars"; city = "Los Angeles"; state = "CA"; region = "West"; postalcode = "90067"; country = "USA"; phone = "(310) 555-6789" },
    @{ address = "1 Place Ville Marie"; city = "Montreal"; state = "QC"; region = "Eastern Canada"; postalcode = "H3B 3Y1"; country = "Canada"; phone = "(514) 555-4567" },
    @{ address = "45 Rockefeller Plaza"; city = "New York"; state = "NY"; region = "Northeast"; postalcode = "10111"; country = "USA"; phone = "(212) 555-4321" },
    @{ address = "60 Wall St"; city = "New York"; state = "NY"; region = "Northeast"; postalcode = "10005"; country = "USA"; phone = "(212) 555-6543" },
    @{ address = "5 Rue Daunou"; city = "Paris"; state = $null; region = $null; postalcode = "75002"; country = "France"; phone = "+33 1 555 1234" },
    @{ address = "400 Capitol Mall"; city = "Sacramento"; state = "CA"; region = "West"; postalcode = "95814"; country = "USA"; phone = "(916) 555-5678" },
    @{ address = "101 E Kennedy Blvd"; city = "Tampa"; state = "FL"; region = "South"; postalcode = "33602"; country = "USA"; phone = "(813) 555-0987" },
    @{ address = "777 S Figueroa St"; city = "Los Angeles"; state = "CA"; region = "West"; postalcode = "90017"; country = "USA"; phone = "(213) 555-1234" },
    @{ address = "1800 G St NW"; city = "Washington"; state = "DC"; region = "Mid-Atlantic"; postalcode = "20006"; country = "USA"; phone = "(202) 555-7891" },
    @{ address = "1 Bligh St"; city = "Sydney"; state = "NSW"; region = $null; postalcode = "2000"; country = "Australia"; phone = "+61 2 555 4321" },
    @{ address = "555 Mission St"; city = "San Francisco"; state = "CA"; region = "West"; postalcode = "94105"; country = "USA"; phone = "(415) 555-1234" },
    @{ address = "66 Wellington St"; city = "Toronto"; state = "ON"; region = "Central Canada"; postalcode = "M5K 1J3"; country = "Canada"; phone = "(416) 555-6789" },
    @{ address = "745 7th Ave"; city = "New York"; state = "NY"; region = "Northeast"; postalcode = "10019"; country = "USA"; phone = "(212) 555-7890" },
    @{ address = "41 Rue Vivienne"; city = "Paris"; state = $null; region = $null; postalcode = "75002"; country = "France"; phone = "+33 1 555 3456" },
    @{ address = "415 Madison Ave"; city = "New York"; state = "NY"; region = "Northeast"; postalcode = "10017"; country = "USA"; phone = "(212) 555-5678" },
    @{ address = "110 Bishopsgate"; city = "London"; state = $null; region = $null; postalcode = "EC2N 4AY"; country = "UK"; phone = "+44 20 555 7890" },
    @{ address = "222 S Riverside Plaza"; city = "Chicago"; state = "IL"; region = "Midwest"; postalcode = "60606"; country = "USA"; phone = "(312) 555-0987" },
    @{ address = "100 Wellington St"; city = "Ottawa"; state = "ON"; region = "Central Canada"; postalcode = "K1A 0A9"; country = "Canada"; phone = "(613) 555-1234" },
    @{ address = "1538 K St NW"; city = "Washington"; state = "DC"; region = "Mid-Atlantic"; postalcode = "20005"; country = "USA"; phone = "(202) 555-6543" },
    @{ address = "700 19th St NW"; city = "Washington"; state = "DC"; region = "Mid-Atlantic"; postalcode = "20431"; country = "USA"; phone = "(202) 555-0987" },
    @{ address = "50 Market St"; city = "Brisbane"; state = "QLD"; region = $null; postalcode = "4000"; country = "Australia"; phone = "+61 7 555 7891" },
    @{ address = "500 Bay St"; city = "Toronto"; state = "ON"; region = "Central Canada"; postalcode = "M5G 1B1"; country = "Canada"; phone = "(416) 555-1234" },
    @{ address = "125 High St"; city = "Boston"; state = "MA"; region = "Northeast"; postalcode = "02110"; country = "USA"; phone = "(617) 555-5678" },
    @{ address = "100 N Tryon St"; city = "Charlotte"; state = "NC"; region = "South"; postalcode = "28202"; country = "USA"; phone = "(704) 555-0987" },
    @{ address = "111 S Wacker Dr"; city = "Chicago"; state = "IL"; region = "Midwest"; postalcode = "60606"; country = "USA"; phone = "(312) 555-4321" },
    @{ address = "350 Bush St"; city = "San Francisco"; state = "CA"; region = "West"; postalcode = "94104"; country = "USA"; phone = "(415) 555-5678" },
    @{ address = "12 Rue Cambon"; city = "Paris"; state = $null; region = $null; postalcode = "75001"; country = "France"; phone = "+33 1 555 6789" }
)


function Get-ManagerId {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Title,

        [Parameter(Mandatory = $true, Position = 1)]
        [PSObject[]]$EmployeeList,

        [Parameter(Mandatory = $true, Position = 2)]
        [hashtable]$Hierarchy
    )
    
    # Initialize random object
    $random = New-Object System.Random

    # Find possible manager titles for the given employee title
    $possibleMgrTitles = $Hierarchy[$Title]
    if ($possibleMgrTitles.Count -eq 0) {
        Write-Verbose "No manager hierarchy found for title: $Title"
        return $null
    }

    # Find available employees with titles that match possible managers
    $possibleMgrs = $EmployeeList | Where-Object {
        $possibleMgrTitles -contains $_.title
    }

    if ($possibleMgrs.Count -eq 0) {
        Write-Verbose "No potential managers found for $Title"
        return $null
    }

    # Randomly pick a manager from the list of available managers
    $selectedManager = $possibleMgrs[$random.Next(0, $possibleMgrs.Count)]
    Write-Verbose "Selected manager ID: $($selectedManager.empid) for title: $Title"
    return $selectedManager.empid
}



Write-Host "Generating Test Data" -ForegroundColor Green

# Generate employees based on the positions and their counts
foreach ($positionInfo in $positionsWithCount) {
    $title = $positionInfo.title
    $count = $positionInfo.count

    for ($i = 0; $i -lt $count; $i++) {
        # Randomly select a name and location
        $randomEmployee = $employees[$random.Next(0, $employees.Length)]
        $randomLocation = $locations[$random.Next(0, $locations.Length)]
        
        # Generate unique empid
        $currentEmpId = $empIdCounter
        $empIdCounter++

        # Generate data for the employee
        $row = [PSCustomObject]@{
            "empid"           = $currentEmpId
            "lastname"        = $randomEmployee.lastname
            "firstname"       = $randomEmployee.firstname
            "title"           = $title
            "titleofcourtesy" = $randomEmployee.titleofcourtesy
            "birthDate"       = $randomEmployee.birthdate
            "hiredate"        = $randomEmployee.hiredate
            "address"         = $randomLocation.address
            "city"            = $randomLocation.city
            "region"          = $randomLocation.state
            "postalcode"      = $randomLocation.postalcode
            "country"         = $randomLocation.country
            "phone"           = $randomLocation.phone
            "mgrid"           = $null  # Set to null temporarily
        }

        # Add the generated employee to the list for future manager lookup
        $generatedEmployees += $row

        # Find a suitable manager for this employee
        $mgrId = Get-ManagerId -Title $title -EmployeeList $generatedEmployees -Hierarchy $hierarchy
        $row.mgrid = $mgrId  # Update the mgrid after manager is found

        Write-Host "Generated data - $($row.lastname) $($row.firstname)`r`n"

        # Add row to sqlData for export
        $sqlData += $row
    }
}

Write-Host "Writing Test Data to Excel" -ForegroundColor Green

# Export to Excel (optional)
# $sqlData | Export-Excel -Path "$PSScriptRoot\TestData.xlsx" -AutoSize -Verbose

$sqlData | ForEach-Object {
    Write-Host "Exporting to Excel: $($_.FirstName) $($_.LastName)" -ForegroundColor Yellow
    $_
} | Export-Excel -Path "$PSScriptRoot\TestData.xlsx" -AutoSize

Write-Host "Test Data written to C:\temp\TestData.xlsx" -ForegroundColor Cyan