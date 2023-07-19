# Load the HTML Agility Pack assembly
Add-Type -Path "HtmlAgilityPack.dll"

# Install the required SQL Server module
#Install-Module -Name SqlServer -AllowClobber -Force

# Import the required ADO.NET assembly
Add-Type -Path "System.Data.SqlClient.dll"

# Access the COMPUTERNAME environment variable
$computerName = $env:COMPUTERNAME

# Specify the path to the executable
$exePath = ".\LookInMyPC.exe"

# Start the executable
$process = Start-Process -FilePath $exePath -PassThru

# Wait for the application to load
Start-Sleep -Seconds 5
    Start-Sleep -Seconds 10
    # Wait for the application to exit
    $process.WaitForExit()

# Specify the path to the HTML file
$htmlFilePath = ".\Reports\$computerName\Page1.htm"

$html = Get-Content -Path $htmlFilePath -Raw

# Database name
$databaseName = "sys_infonew"

# Connection string for the Local SQL Server
$connectionString = "Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=$databaseName;Integrated Security=True;Connect Timeout=30;Encrypt=False;Application Name=HTML2DATABASE"

# Connection string for the Azure SQL Server
#$connectionString = "Server=tcp:getintoit1dbserver.database.windows.net,1433;Initial Catalog=SectionInformation_db;Persist Security Info=False;User ID=AshwaniGusain;Password=*****;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

# Switch the context to the existing database
#$useDatabaseSql = "USE [$databaseName]"
$useDatabaseConnection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
$useDatabaseConnection.Open()
#$useDatabaseCommand = $useDatabaseConnection.CreateCommand()
#$useDatabaseCommand.CommandText = $useDatabaseSql
#$useDatabaseCommand.ExecuteNonQuery()


$tableHeadingHTML = "Network Adapter Information - What's This?"
$tableUserId = "Windows Information - What's This?"
# Create an HTML document object
$htmlDocument = New-Object HtmlAgilityPack.HtmlDocument
$htmlDocument.LoadHtml($html)

$tables = $htmlDocument.DocumentNode.SelectNodes("//table[@style='width: 900px;WORD-BREAK:BREAK-ALL']")

foreach ($table in $tables) {
    # Extract the table heading
    $tableHeading = $table.SelectSingleNode("tr/td").InnerText

    if ($tableHeading -eq $tableHeadingHTML -or $tableHeading -eq $tableUserId) 
    {
      $tableRows = $table.SelectNodes("tr[position() > 1]")

    foreach ($row in $tableRows[1..($tableRows.Count - 1)]) {
         $columnValues = $row.SelectNodes("td") | ForEach-Object { $_.InnerText }

         if ($columnValues -like "Logged In User" )
         {
            $desiredUserDetails = $columnValues[1]
            break

         }

        }

        $desiredIPValue = $columnValues[3]
        }
                         
        }

        # Update sys_IP and sys_name column in systems table
        $InsertSysIpSql = "INSERT INTO systems (sys_IP, sys_name) VALUES ('$desiredIPValue', '$desiredUserDetails');"
        #write-Host "$InsertSysIpSql"
        $InsertSysIpCommand = $useDatabaseConnection.CreateCommand()
        $InsertSysIpCommand.CommandText = $InsertSysIpSql
        $InsertSysIpCommand.ExecuteNonQuery()
                    
        Write-Host "Inserted sys_IP in systems table with IP Address value."


         # Fetch id and sys_heading columns from systems table
        $systemsQuery = "SELECT Id, sys_IP, sys_name FROM systems where sys_IP = '$desiredIPValue'"
        $systemsCommand = $useDatabaseConnection.CreateCommand()
        $systemsCommand.CommandText = $systemsQuery
        $systemsReader = $systemsCommand.ExecuteReader()

            while ($systemsReader.Read()) {
            $sysid = $systemsReader["Id"]
            }

            $systemsReader.Close()
            $systemsCommand.Dispose()

         # Fetch id and sys_heading columns from systemsInformation table
        $systemsInformationQuery = "SELECT Id, sys_heading FROM sectioninfo"
        $systemsInformationCommand = $useDatabaseConnection.CreateCommand()
        $systemsInformationCommand.CommandText = $systemsInformationQuery
        $systemsInformationReader = $systemsInformationCommand.ExecuteReader()

while ($systemsInformationReader.Read()) {
    $id = $systemsInformationReader["Id"]
    $sysHeading = $systemsInformationReader["sys_heading"]

    foreach ($table in $tables) {
        #Write-Host "Im here"

        # Extract the table heading
        $tableHeading = $table.SelectSingleNode("tr/td").InnerText
        $trimmedHeading = ($tableHeading -split '-')[0].Trim()

        #Write-Host "$tableHeading"

        # Select all table rows except the first one (heading)
        $tableRows = $table.SelectNodes("tr[position() > 1]")

        #Getting columnsNames
        $columnNames = $tableRows[0].SelectNodes("td") | ForEach-Object { $_.InnerText }

        for ($i = 0; $i -lt $columnNames.Length; $i++) {


        # Update sys_IP and sys_name column in systems table

        $wordCount = $columnNames.Split(' ').Count
        Write-Host "$wordCount"

        if ($wordCount -gt 1) {
                    
                    $coln = $columnNames[$i]
                }
                else {
        $coln = $columnNames
        }
       $InsertColumns = @"
        INSERT INTO sectioninfo_columns (sys_id,section_id,column_name ) VALUES ('$sysid', '$id', '$coln');

        SELECT SCOPE_IDENTITY() AS LastInsertedID;
"@

                $insertConnection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
                $insertConnection.Open()

                try {
                    # Execute the insert SQL statement
                    $InsertColumnsCommand = $insertConnection.CreateCommand()
                    $InsertColumnsCommand.CommandText = $InsertColumns
                    $lastInsertedColumnID = $InsertColumnsCommand.ExecuteScalar()
                }
                finally {
                    # Close and dispose of the insert command and connection
                    $InsertColumnsCommand.Dispose()
                    $insertConnection.Close()
                    $insertConnection.Dispose()
                }
        


        # Compare the sys_heading value with the table heading
        if ($sysHeading -eq $trimmedHeading) {
            Write-Host "Table heading matches sys_heading in sectioninfo table."
            $tableRows = $table.SelectNodes("tr[position() > 1]")

            foreach ($row in $tableRows[1..($tableRows.Count - 1)]) {
                $columnValues = $row.SelectNodes("td") | ForEach-Object { $_.InnerText }
                #$columnValues = $row.SelectNodes("td") | Where-Object { $_.InnerText -eq $coln } | ForEach-Object { $_.InnerText }

                # Generate the column and value SQL statements
                $columnsSql = "[" + ($columnNames -join "], [") + "]"
                $valuesSql = "'" + ($columnValues -join "', '") + "'"

                $wordCount1 = $columnValues.Split(' ').Count

                 if ($wordCount -gt 1) {
                    
                    $col = $columnValues[$i]
                }
                else {
                $col = $columnValues
                }

                #$col = $columnValues

                # Generate the insert SQL statement
                $insertDataSql = "INSERT INTO sectioninfo_rowsdata (system_id, sectioninfo_id, Column_id, rowdata) 
                                  VALUES ('$sysid', '$id', '$lastInsertedColumnID', '$col')"
                
                # Create a new connection for insert command
                $insertConnection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
                $insertConnection.Open()


                try {
                    # Execute the insert SQL statement
                    $insertDataCommand = $insertConnection.CreateCommand()
                    $insertDataCommand.CommandText = $insertDataSql
                    $insertDataCommand.ExecuteNonQuery()
                }
                finally {
                    # Close and dispose of the insert command and connection
                    $insertDataCommand.Dispose()
                    $insertConnection.Close()
                    $insertConnection.Dispose()
                }
            }
        }
    }
}
}

$systemsInformationReader.Close()
$systemsInformationReader.Dispose()
$systemsInformationCommand.Dispose()
