# Path to the log file
$logFilePath = ".\Log.txt"

# Function to write log messages to the log file
function Write-Log($message) {
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $message"
    $logMessage | Out-File -FilePath $logFilePath -Append
}

# Add a trap to handle script termination
$ErrorActionPreference = 'Stop'
trap {
    Write-Log "Script terminated with error: $($_.Exception.Message)"
    Exit 1
}

try {

Write-Log "Log Started : Loading the HTML Agility Pack assembly..."
# Load the HTML Agility Pack assembly
Add-Type -Path "HtmlAgilityPack.dll"

Write-Log "Importing the required ADO.NET assembly..."
# Import the required ADO.NET assembly
Add-Type -Path "System.Data.SqlClient.dll"

# Access the COMPUTERNAME environment variable
$computerName = $env:COMPUTERNAME
Write-Log "Computer Name: $computerName"


# Specify the path to the executable
$exePath = ".\LookInMyPC.exe"

Write-Log "Opening LookInMyPC..."
# Start the executable
$process = Start-Process -FilePath $exePath -PassThru

# Wait for the application to load
Start-Sleep -Seconds 5

# Optional: Wait for the report generation to complete (adjust the sleep time as needed)
Start-Sleep -Seconds 10

# Wait for the application to exit
$process.WaitForExit()
Write-Log "LookInMyPC exits..."

# Check if XAMPP service is running
$xamppService = $services | Where-Object { $_.Name -eq "XAMPP" -and $_.Status -eq "Running" }

# Check if IIS service is running
$iisService = $services | Where-Object { $_.Name -eq "W3SVC" -and $_.Status -eq "Running" }

$SqlServerServiceName = "MSSQLSERVER"  # Add more service names as needed

# Specify the path to the HTML file
$htmlFilePath = ".\Reports\$computerName\Page1.htm"
$html = Get-Content -Path $htmlFilePath -Raw

  $iisStatus = ""
  $xamppService = ""
  $databaseServerRunning = ""
  $LocaldatabaseName = "Sql Server"
  $xamppServiceNew = "NO"

# Function to get the general server details
function GetServerGeneralInfo($sysid) {
    try {
        
            Write-Log "Getting General information of server"
            $sqlServerService = Get-Service -Name $SqlServerServiceName -ErrorAction SilentlyContinue

            if ($sqlServerService -eq $null) {
                $databaseServerRunning = "Not Installed"
                #Write-Host "SQL Server is not installed or the service name is incorrect."
                }
            elseif ($sqlServerService.Status -eq "Running") {
                $databaseServerRunning = "Running"
                #Write-Host "SQL Server is running."
                }
           else {
                    $databaseServerRunning = "Not Running"
                    #Write-Host "SQL Server is installed but not running."
                }

        # Check if XAMPP service is running
        $xamppService = $services | Where-Object { $_.Name -eq "XAMPP" -and $_.Status -eq "Running" }
        if ($xamppService)
        {
            $xamppServiceNew = "YES"
        }
        # Get IIS status
        $iisStatus = (Get-Service -Name W3SVC).Status

        # Insert command for ServerGenralInfo table
$InsertGeneralSystemdataSql = @"
                INSERT INTO ServerGenralInfo (sys_id, IsDatabaseServerAvailable, DatabaseName, IsIISActive, IsXamppActive)
                VALUES ('$sysid', '$databaseServerRunning', '$LocaldatabaseName', '$iisStatus', '$xamppServiceNew' );
"@
        # Insert data to table
        InsertData $InsertGeneralSystemdataSql
    }
    catch {
        Write-Log "Error While getting Server GenralInfo"
        Write-Log $_
        Exit 1
    }
}

# Database name
$databaseName = "SectionInformation_db"

# Connection string for the Local SQL Server
#$connectionString = "Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=$databaseName;Integrated Security=True;Connect Timeout=30;Encrypt=False;Application Name=FinalHTMLTODB"


# Connection string for the Azure SQL Server
$connectionString = "Server=tcp:getintoit1dbserver.database.windows.net,1433;Initial Catalog=SectionInformation_db;Persist Security Info=False;User ID=AshwaniGusain;Password=*****;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

# Create a new connection for database operations
try {
    $useDatabaseConnection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $useDatabaseConnection.Open()
    Write-Log "Connected to the database."
}
catch {
    Write-Log "Error connecting to the database: $_"
    Exit 1
}

# Function to execute a SQL query and return a single value
function GetSingleValueFromQuery($query) {
    try {
        $command = $useDatabaseConnection.CreateCommand()
        $command.CommandText = $query
        $result = $command.ExecuteScalar()
        $command.Dispose()
        return $result
    }
    catch {
        Write-Log "Error executing query: $query"
        Write-Log $_
        Exit 1
    }
}

# Function to execute a SQL insert query and return the last inserted ID
function InsertAndGetLastID($query) {
    try {
        $command = $useDatabaseConnection.CreateCommand()
        $command.CommandText = $query
        $command.ExecuteNonQuery()
        $command.Dispose()

        # Fetch the last inserted ID using SCOPE_IDENTITY()
        $identityQuery = "SELECT SCOPE_IDENTITY() AS LastInsertedID"
        $identityCommand = $useDatabaseConnection.CreateCommand()
        $identityCommand.CommandText = $identityQuery
        $lastInsertedID = $identityCommand.ExecuteScalar().Value
        $identityCommand.Dispose()

        return $lastInsertedID
    }
    catch {
        Write-Log "Error executing query: $query"
        Write-Log $_
        Exit 1
    }
}



# Function to insert data into the database
function InsertData($query) {
    try {
        $command = $useDatabaseConnection.CreateCommand()
        $command.CommandText = $query
        $command.ExecuteNonQuery()
        $command.Dispose()
        }
    catch {
        Write-Log "Error Inserting the data: $query"
        Write-Log $_
        Exit 1
    }
}

# Create an HTML document object
$htmlDocument = New-Object HtmlAgilityPack.HtmlDocument
$htmlDocument.LoadHtml($html)

$tableHeadingHTML = "Network Adapter Information - What's This?"
$tableUserId = "Windows Information - What's This?"

# Fetch desiredUserDetails and desiredIPValue
$desiredUserDetails = ""
$desiredIPValue = ""

$tables = $htmlDocument.DocumentNode.SelectNodes("//table[@style='width: 900px;WORD-BREAK:BREAK-ALL']")

foreach ($table in $tables) {
    # Extract the table heading
    $tableHeading = $table.SelectSingleNode("tr/td").InnerText

    if ($tableHeading -eq $tableHeadingHTML -or $tableHeading -eq $tableUserId) {
        $tableRows = $table.SelectNodes("tr[position() > 1]")

        foreach ($row in $tableRows[1..($tableRows.Count - 1)]) {
            $columnValues = $row.SelectNodes("td") | ForEach-Object { $_.InnerText }

            if ($columnValues -like "Logged In User") {
                $desiredUserDetails = $columnValues[1]
                break
            }
        }

        $desiredIPValue = $columnValues[3]
    }
}

$currentTime = Get-Date
$dateTimeObject = [DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss")

# Fetch the latest iteration number for the specified IP address
$sysiteration = GetSingleValueFromQuery("SELECT ISNULL(MAX(IterationNumber), 0) FROM systems WHERE sys_IP LIKE '%$desiredIPValue%'")

# Increment the iteration number
$sysiteration++

# Insert or update data in the systems table
$InsertSysIpSql = @"
    IF NOT EXISTS (SELECT 1 FROM systems WHERE sys_IP = '$desiredIPValue' AND IterationNumber = '$sysiteration')
    BEGIN
        INSERT INTO systems (sys_IP, sys_name, IterationNumber, sys_time)
        VALUES ('$desiredIPValue', '$desiredUserDetails', '$sysiteration', '$dateTimeObject')
    END
    ELSE
    BEGIN
        UPDATE systems
        SET sys_name = '$desiredUserDetails', sys_time = '$dateTimeObject'
        WHERE sys_IP = '$desiredIPValue' AND IterationNumber = '$sysiteration'
    END
"@
InsertData $InsertSysIpSql

# Fetch the sysid from systems table
$sysid = GetSingleValueFromQuery("SELECT Id FROM systems WHERE sys_IP = '$desiredIPValue' AND IterationNumber = '$sysiteration'")


# call to get initial system information
GetServerGeneralInfo $sysid


# Fetch the sectioninfo IDs and sys_headings from systemsInformation table
$systemsInformationQuery = "SELECT Id, sys_heading FROM sectioninfo"
$systemsInformationCommand = $useDatabaseConnection.CreateCommand()
$systemsInformationCommand.CommandText = $systemsInformationQuery
$systemsInformationReader = $systemsInformationCommand.ExecuteReader()

Write-Log "This is IterationNumber : $sysiteration processing on  : $desiredIPValue"
# Create a DataTable to store the results from the SQL query
$dataTable = New-Object System.Data.DataTable

# Close the DataReader before filling the DataTable
$systemsInformationReader.Close()
$systemsInformationReader.Dispose()

# Use a SqlDataAdapter to fill the DataTable with data from the SQL query
$adapter = New-Object System.Data.SqlClient.SqlDataAdapter
$adapter.SelectCommand = $systemsInformationCommand
$adapter.Fill($dataTable)

$systemsInformationCommand.Dispose()

foreach ($row in $dataTable.Rows) {
    $id = $row["Id"]
    $sysHeading = $row["sys_heading"]

    foreach ($table in $tables) {
        # Extract the table heading
        $tableHeading = $table.SelectSingleNode("tr/td").InnerText
        $trimmedHeading = ($tableHeading -split '-')[0].Trim()

        # Compare the sys_heading value with the table heading
        if ($sysHeading -eq $trimmedHeading) {
            Write-Host "Table heading matches sys_heading in sectioninfo table."
            $tableRows = $table.SelectNodes("tr[position() > 1]")

            # Extract column names from the first row
            $columnNames = $tableRows[0].SelectNodes("td") | ForEach-Object { $_.InnerText }
                 $wordCount = $columnNames.Split(' ').Count
            for ($i = 0; $i -lt $columnNames.Length; $i++) {

            #Write-Host "$wordCount"

        if ($wordCount -gt 1) {
                    
                    $coln = $columnNames[$i]
                }
                else {
        $coln = $columnNames
        }
            # Insert column names into sectioninfo_columns table
            $columnsSql = @"
                INSERT INTO sectioninfo_columns (sys_id, section_id, column_name, IterationNumber)
                VALUES ('$sysid', '$id', '$coln', '$sysiteration');
"@
         #write-host "$columnsSql" 
        $lastInsertedColumnID = InsertAndGetLastID $columnsSql
        #write-host "$lastInsertedColumnID" 
# Iterate through the data rows and insert row data into sectioninfo_rowsdata table
foreach ($row in $tableRows[1..($tableRows.Count - 1)]) {
    $columnValues = $row.SelectNodes("td") | ForEach-Object { $_.InnerText -replace "'", "''" }

    $wordCount1 = $columnValues.Split(' ').Count

                 if ($wordCount -gt 1) {
                    
                    $col = $columnValues[$i]
                }
                else {
                $col = $columnValues
                }


    # Generate the insert SQL statement
    $insertDataSql = @"
        INSERT INTO sectioninfo_rowsdata (system_id, sectioninfo_id, Column_id, rowdata, IterationNumber)
        VALUES ('$sysid', '$id', '$lastInsertedColumnID', '$coln', '$sysiteration');
"@
    InsertData $insertDataSql
    }
   }
 }
}
}


$systemsInformationReader.Close()
$systemsInformationReader.Dispose()
$systemsInformationCommand.Dispose()

# Close and dispose of the database connection
$useDatabaseConnection.Close()
$useDatabaseConnection.Dispose()
Write-Log "Log Ended :Script completed successfully."
}
catch {
    Write-Log "Error in the script: $_"
    Exit 1
}


finally {
    # Close and dispose of the database connection
    if ($useDatabaseConnection -ne $null) {
        if ($useDatabaseConnection.State -eq 'Open') {
            try {
                $useDatabaseConnection.Close()
                Write-Log "May be user cancelled the operation."
            }
            catch {
                Write-Log "Error closing the database connection: $_"
            }
        }
        $useDatabaseConnection.Dispose()
    }
}
