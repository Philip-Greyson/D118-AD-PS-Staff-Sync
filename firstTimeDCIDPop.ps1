# Using SQL query code from https://tsql.tech/how-to-read-data-from-oracle-database-via-powershell-without-using-odbc-or-installing-oracle-client-and-import-it-to-sql-server-too/

#Parameters
$OracleDLLPath = ".\Oracle.ManagedDataAccess.dll"
#The oracle DataSource as you would compile it in TNSNAMES.ORA
$datasource = " (DESCRIPTION = 
                (ADDRESS = (PROTOCOL = TCP)(HOST  = " + $Env:POWERSCHOOL_PROD_DB_IP + ")(PORT = "+ $Env:POWERSCHOOL_PROD_DB_PORT +"))
                (CONNECT_DATA = 
                (SERVER =  DEDICATED)
                (SERVICE_NAME = " + $Env:POWERSCHOOL_PROD_DB_NAME + ")
                (FAILOVER_MODE = (TYPE = SELECT)
                (METHOD =  BASIC)
                (RETRIES = 180)
                (DELAY = 5))))"
$username = $Env:POWERSCHOOL_READ_USER # get the username of read-only account from environment variables
$password = $Env:POWERSCHOOL_DB_PASSWORD # get the password from environment variable

#Load Required Types and modules
Add-Type -Path $OracleDLLPath
Import-Module SqlServer
Import-Module ActiveDirectory

# Clear out log file from previous run
Clear-Content -Path .\dcidLog.txt


#Create the connection string
$connectionstring = 'User Id=' + $username + ';Password=' + $password + ';Data Source=' + $datasource 
#Create the connection object
$con = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionstring)
#Create a command and configure it

$querySchools = "SELECT name, school_number, abbreviation FROM schools" # make a query to find a list of schools

$cmd = $con.CreateCommand()
$cmd.CommandText = $querySchools
$cmd.CommandTimeout = 3600 #Seconds
$cmd.FetchSize = 10000000 #10MB
#Creates a data adapter for the command
$da = New-Object Oracle.ManagedDataAccess.Client.OracleDataAdapter($cmd);
#The Data adapter will fill this DataTable
$schools = New-Object System.Data.DataTable # each row is one entry, the columns can be accessed with []
#Only here the query is sent and executed in Oracle 
[void]$da.fill($schools)

$spacer = "--------------------------------------------------"
$constantOU = $Env:CONSTANT_OU_ENDING
$badNames = 'Use', 'Training1','Trianing2','Trianing3','Trianing4','Planning','Admin','Nurse','User', 'Use ', 'Test', 'Testtt', 'Do Not', 'Do', 'Not', 'Tbd', 'Lunch' # define list of names to ignore

foreach ($school in $Schools)
{
    $schoolName = $school[0].ToString().ToUpper()
    $schoolNum = $school[1]
    $schoolAbbrev = $school[2]
    $OUPath = "OU=Staff,OU=$schoolName,$constantOU"
    $schoolInfo = "STARTING BUILDING: $schoolName | $schoolNum | $OUPath  | $schoolAbbrev"

    # print out a space line and the school info header to console and log file
    Write-Output $spacer
    Write-Output $spacer | Out-File -FilePath .\dcidLog.txt -Append
    Write-Output $schoolInfo 
    Write-Output $schoolInfo | Out-File -FilePath .\dcidLog.txt -Append
    Write-Output $spacer
    Write-Output $spacer | Out-File -FilePath .\dcidLog.txt -Append

    $userQuery = "SELECT users.last_name, users.first_name, users.email_addr, users.dcid, users.preferredname, schoolstaff.status `
                        FROM users INNER JOIN schoolstaff ON users.dcid = schoolstaff.users_dcid WHERE users.email_addr IS NOT NULL AND users.homeschoolid = $schoolNum AND schoolstaff.schoolid = $schoolNum ORDER BY users.dcid"

    $cmd.CommandText = $userQuery
    $da = New-Object Oracle.ManagedDataAccess.Client.OracleDataAdapter($cmd);
    #The Data adapter will fill this DataTable
    $resultSet = New-Object System.Data.DataTable # each row is one entry, the columns can be accessed with []
    #Only here the query is sent and executed in Oracle 
    [void]$da.fill($resultSet)
    #Close the connection
    $con.Close()
    foreach ($result in $resultSet)
    {
        $lastName = $result[0]
        $firstName = $result[1]
        # check to see if their first or last name is in the list of "bad names"
        if (($badNames -notcontains $firstName) -and ($badNames -notcontains $lastName))
        {
            $email = $result[2]
            $uDCID = $result[3]
            $preferredName = $result[4]
            if (![string]::IsNullOrEmpty($preferredName)) # if they have a preferred name in powerschool, overwrite the first name with it
            {
                $firstName = $preferredName
            }
            $firstInitial = $firstName.Substring(0,1) # get the first initial by getting the first character of the first name
            $samAccountName = $lastName.ToLower().replace(" ", "-").replace("'", "") + $firstInitial.ToLower()
            $userInfo = "Processing User: First: $firstName | Last: $lastName | Email: $email | DCID: $uDCID"
            Write-Output $userInfo
            $userInfo | Out-File -FilePath .\dcidLog.txt -Append #
            $adUser = Get-ADUser -Filter "pSuDCID -eq $uDCID" -Properties mail
            # see if we have a match for the user based on DCID, if not we want to search for an account that might be correct but not populated with the DCID
            if (!$adUser) # if no match for the DCID
            {
                $message = "    WARNING: No AD account found for $uDCID, searching for $samAccountName and $email"
                Write-Output $message # write to console
                $message | Out-File -FilePath .\dcidLog.txt -Append # write to log file
                $adUser = Get-ADUser -Filter {(EmailAddress -eq $email) -or (SamAccountName -eq $samAccountName)} -Properties mail,pSuDCID # search for a user that either has a matching email or SamAccountName
                if ($adUser)
                {
                    $currentEmail = $adUser.mail
                    $currentSamAccountName = $adUser.SamAccountName
                    $currentDCID = $adUser.pSuDCID
                    if (!$currentDCID)
                    {
                        if ($currentSamAccountName -eq $samAccountName)
                        {
                            if ($currentEmail -eq $email)
                            {
                                $message = "        ACTION: samAccountName $samAccountName with $currentEmail matches, updating with DCID $uDCID"
                                Write-Output $message # write to console
                                $message | Out-File -FilePath .\dcidLog.txt -Append # write to log file
                                # Read-Host "Press ENTER to CONFIRM..." # pauses before the actual change is made
                                # Set-ADUser $adUser -Add @{pSuDCID=$uDCID} # adds the custom attribute PSuDCID to the user
                            }
                            else 
                            {
                                $message = "        UNCERTAINTY: Desired email $email does not match current one of $currentEmail, NOT UPDATING"
                                Write-Output $message
                                $message | Out-File -FilePath .\dcidLog.txt -Append
                            }
                        }
                        else 
                        {
                            $message = "        UNCERTAINTY: Desired samAccountName $samAccountName does not match current one of $currentSamAccountName, NOT UPDATING"
                            Write-Output $message
                            $message | Out-File -FilePath .\dcidLog.txt -Append
                        }
                    }
                    else
                    {
                        $message = "        UNCERTAINTY: Found account with matching info $currentSamAccountName and $currentEmail but with existing DCID $currentDCID"
                        Write-Output $message
                        $message | Out-File -FilePath .\dcidLog.txt -Append
                    }
                }
            }
            else 
            {
                $samAccountName = $adUser.SamAccountName
                $message = "    SUCCESS: User $uDCID already found under $samAccountName"
                Write-Output $message
                $message | Out-File -FilePath .\dcidLog.txt -Append 
            }
        }
        else # otherwise if we found a match for the DCID
        {
            $message = "INFO: found user matching name in bad names list: $firstName $LastName"
            Write-Output $message
            $message | Out-File -FilePath .\dcidLog.txt -Append
        }
    }
}