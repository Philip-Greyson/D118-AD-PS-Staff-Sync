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
Clear-Content -Path .\syncLog.txt
# repadmin.exe /showrepl *
repadmin.exe /syncall D118-DIST-OFF /Aed # synchronize the controllers so they all have updated data
# break

#Create the connection string
$connectionstring = 'User Id=' + $username + ';Password=' + $password + ';Data Source=' + $datasource 
#Create the connection object
$con = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionstring)

# make a query to find a list of schools
$querySchools = "SELECT name, school_number, abbreviation FROM schools" 

#Create a command and configure it
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

$spacer = "--------------------------------------------------" # just a spacer for printing easily repeatedly
$constantOU = $Env:CONSTANT_OU_ENDING # define the constant parts of our AD OU structure. Essentially everything after our building level
$defaultPassword = ConvertTo-SecureString $Env:AD_NEW_USER_PASSWORD -AsPlainText -Force  # define the default password used for new accounts

$staffJobTypes = "Not Assigned","Teacher","Staff","Lunch Staff","Substitute"
$badNames = 'Use', 'Training1','Trianing2','Trianing3','Trianing4','Planning','Admin','Nurse','User', 'Use ', 'Test', 'Testtt', 'Do Not', 'Do', 'Not', 'Tbd', 'Lunch', 'Formbuilder', 'Human' # define list of names to ignore

# define our district wide employee AD groups
$districtTeacherGroup = "D118 Teachers"
$districtStaffGroup = "D118 Staff"
$districtSubGroup = "D118 Substitutes"
$papercutGroup = "Papercut Staff Group"
# find the members of these district wide groups so we only have to do it once and then can reference them later
$districtStaffMembers = Get-ADGroupMember -Identity $districtStaffGroup -Recursive | Select-Object -ExpandProperty samAccountName
$districtTeacherMembers = Get-ADGroupMember -Identity $districtTeacherGroup -Recursive | Select-Object -ExpandProperty samAccountName
$districtSubMembers = Get-ADGroupMember -Identity $districtSubGroup -Recursive | Select-Object -ExpandProperty samAccountName
$papercutStaffMembers = Get-ADGroupMember -Identity $papercutGroup -Recursive | Select-Object -ExpandProperty samAccountName

foreach ($school in $Schools)
{
    $schoolName = $school[0].ToString().ToUpper()
    $schoolNum = $school[1]
    $schoolAbbrev = $school[2]
    $OUPath = "OU=Staff,OU=$schoolName,$constantOU"
    $schoolInfo = "STARTING BUILDING: $schoolName | $schoolNum | $OUPath  | $schoolAbbrev"

    # print out a space line and the school info header to console and log file
    Write-Output $spacer
    Write-Output $spacer | Out-File -FilePath .\syncLog.txt -Append
    Write-Output $schoolInfo 
    Write-Output $schoolInfo | Out-File -FilePath .\syncLog.txt -Append
    Write-Output $spacer
    Write-Output $spacer | Out-File -FilePath .\syncLog.txt -Append

    # get the members of the teachers and staff groups at the current building, for reference in each user without querying every time. Ignoring buildings where these groups do not exist
    if (($schoolAbbrev -ne "O-HR") -and ($schoolAbbrev -ne "SUM") -and ($schoolAbbrev -notlike "DNU *") -and ($schoolAbbrev -ne "Graduated Students") -and ($schoolAbbrev -ne "AUX"))
    {
        $schoolTeacherGroup = $schoolAbbrev + " Teachers"
        $schoolStaffGroup = $schoolAbbrev + " Staff"
        $schoolTeacherMembers = Get-ADGroupMember -Identity $schoolTeacherGroup -Recursive | Select-Object -ExpandProperty samAccountName
        $schoolStaffMembers = Get-ADGroupMember -Identity $schoolStaffGroup -Recursive | Select-Object -ExpandProperty samAccountName
        # debug group memberships
        # Write-Output $schoolInfo | Out-File -FilePath .\syncLog.txt -Append
        # "Teachers:" | Out-File -FilePath .\syncLog.txt -Append
        # Write-Output $schoolInfo | Out-File -FilePath .\syncLog.txt -Append
        # $schoolTeacherMembers | Out-File -FilePath .\syncLog.txt -Append # output to the syncLog.txt file
        # Write-Output $schoolInfo | Out-File -FilePath .\syncLog.txt -Append
        # "Staff:" | Out-File -FilePath .\syncLog.txt -Append
        # Write-Output $schoolInfo | Out-File -FilePath .\syncLog.txt -Append
        # $schoolStaffMembers | Out-File -FilePath .\syncLog.txt -Append # output to the syncLog.txt file
    }
    # create a new query to find the users in the current building
    $userQuery = "SELECT users.last_name, users.first_name, users.email_addr, users.teachernumber, schoolstaff.status, schoolstaff.staffstatus, users.dcid, users.preferredname, u_humanresources.jobtitle `
                        FROM users INNER JOIN schoolstaff ON users.dcid = schoolstaff.users_dcid LEFT OUTER JOIN u_humanresources ON users.dcid = u_humanresources.usersdcid`
                        WHERE users.email_addr IS NOT NULL AND users.homeschoolid = $schoolNum AND schoolstaff.schoolid = $schoolNum ORDER BY users.dcid" # query to get all the staff info in that building
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
        $OUPath = "OU=Staff,OU=$schoolName,$constantOU" # reset the OUPath to the normal one otherwise it will stay on suspended once we get one
        $lastName = $result[0].ToLower() # take the last name and convert to all lower case
        $lastName = (Get-Culture).TextInfo.ToTitleCase($lastName) # take the last name all lowercase string and now convert to title case
        $firstName = $result[1].ToLower()
        $firstName = (Get-Culture).TextInfo.ToTitleCase($firstName) # take the last name all lowercase string and now convert to title case
        $email = $result[2]
        $teachNumber = $result[3]
        $active = $result[4] # integer from PS, 1 for active, 2 for inactive
        $staffType = $result[5] # integer 0-4, maps to the staffJobTypes array above
        $uDCID = $result[6] # unique user DCID from PS, used to track which user is which even if their name and email changes
        $preferredName = $result[7]
        $jobTitle = $result[8]
        if (![string]::IsNullOrEmpty($preferredName)) # if they have a preferred name in powerschool, overwrite the first name with it
        {
            $preferredName = $preferredName.ToLower()
            $preferredName = (Get-Culture).TextInfo.ToTitleCase($preferredName)
            $firstName = $preferredName
        }
        if (($badNames -notcontains $firstName) -and ($badNames -notcontains $lastName))
        {
            $firstInitial = $firstName.Substring(0,1) # get the first initial by getting the first character of the first name
            $samAccountName = $lastName.ToLower().replace(" ", "-").replace("'", "") + $firstInitial.ToLower()
            $jobType = $staffJobTypes[$staffType]
            $userInfo = "INFO: Processing User: First: $firstName | Last: $lastName | Email: $email | Active: $active | Type: $jobType | Teacher ID: $teachNumber | DCID: $uDCID"
            Write-Output $userInfo
            $userInfo | Out-File -FilePath .\syncLog.txt -Append # output to the syncLog.txt file
            if ($active -eq 1)
            { # if they have a 1 they are active, anything else is inactive
                $adUser = Get-ADUser -Filter "pSuDCID -eq $uDCID" -Properties mail,title,department,description,homedirectory # do a query for existing users with the custom attribute pSuDCID that equals the users DCID
                if ($adUser)
                { # if we find a user with a matchind DCID, just update their info
                    $currentSamAccountName = $adUser.SamAccountName
                    $currentFullName = $adUser.name
                    $message = "  User with DCID $uDCID already exists under samname $currentSamAccountName, object full name $currentFullName. Updating any info"
                    Write-Output $message # write to console
                    $message | Out-File -FilePath .\syncLog.txt -Append # write to log file

                    # Check to see if their name has changed, update the name fields and the sam account name
                    if (($firstName -ne $adUser.GivenName) -or ($lastName -ne $adUser.Surname) )
                    {
                        $currentFirst = $adUser.GivenName
                        $currentLast = $adUser.Surname
                        $message = "      ACTION: NAME: User $uDCID changed names, updating from $currentFirst $currentLast to $firstName $lastName"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                        Set-ADUser $adUser -GivenName $firstName -Surname $lastName
                        if ($currentSamAccountName -ne $samAccountName)
                        {
                            $message = "      ACTION: SAMNAME: User $uDCID changed names, updating account name from $currentSamAccountName to $samAccountName and 'full name' from $currentFullName to $samAccountName"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                            try # try to set the unique name, but catch if it fails and try some other permutations
                            {
                                Set-ADUser $adUser -SamAccountName $samAccountName -Name $samAccountName # update the actual samAccountName and the object name
                                $currentSamAccountName = $samAccountName
                            }
                            catch
                            {
                                $message = "          ERROR: Could not change $currentSamAccountName to $samAccountName, trying with full first name"
                                Write-Output $message # write to console
                                $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                                Write-Output $_ # write out the actual error
                                $_ | Out-File -FilePath .\syncLog.txt -Append
                                # add their full first name after a period after the last name
                                $samAccountName = $lastName.ToLower().replace(" ", "-").replace("'", "") + "." + $firstName.ToLower().replace(" ", "-").replace("'", "")
                                try # try to set the name again
                                {
                                    Set-ADUser $adUser -SamAccountName $samAccountName -Name $samAccountName # update the actual samAccountName and the object name
                                    $currentSamAccountName = $samAccountName
                                }
                                catch 
                                {
                                    $message =  "          ERROR: Could not change $currentSamAccountName to $samAccountName, out of tries, stopping"
                                    Write-Output $message # write to console
                                    $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                                    Write-Output $_ # write out the actual error
                                    $_ | Out-File -FilePath .\syncLog.txt -Append
                                }
                            }
                        }
                    }

                    # Check to make sure their user account is enabled
                    if (!$adUser.Enabled)
                    {
                        $message = "      ACTION: ENABLE: Enabling user $currentSamAccountName - $uDCID - $email"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                        Enable-ADAccount $adUser # enables the selected account
                    }

                    # Check to see if their email has changed (due to a name change), update all relevant fields
                    if ($email -ne $adUser.mail)
                    {
                        $oldEmail = $adUser.mail
                        $message = "      ACTION: EMAIL: User $firstName $lastName - $uDCID - has had their email change from $oldEmail to $email, changing"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                        # Set-ADUser $adUser -EmailAddress $email -UserPrincipalName $email # update the user's email and principal name which is also their email
                    }
                    
                    # Check to see if they are in the right OU, move them if not
                    $properDistinguished = "CN=$currentFullName,$OUPath" # construct what the desired/proper distinguished name should be based on their samaccount name and the OU they should be in
                    if ($properDistinguished -ne $adUser.DistinguishedName)
                    {
                        $currentDistinguished =  $adUser.DistinguishedName
                        try
                        {
                            $message = "      ACTION: OU: User $currentSamAccountName NOT in correct OU, moving from $currentDistinguished to $properDistinguished"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                            Move-ADObject $adUser -TargetPath $OUPath # moves the targeted AD user account to the correct OU
                        }
                        catch 
                        {
                            $message =  "          ERROR: Could not move $currentSamAccountName to $OUPath"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                            Write-Output $_ # write out the actual error
                            $_ | Out-File -FilePath .\syncLog.txt -Append
                        }
                    }

                    # Check to see if their teacher number and staff type are correct, update if not
                    if (($adUser.title -ne $teachNumber) -or ($adUser.department -ne $jobType))
                    {
                        $oldTitle = $adUser.title
                        $oldDept = $adUser.department
                        $message = "      ACTION: TITLE-DEPARTMENT: Updating user $uDCID's title from $oldTitle to $teachNumber and department from $oldDept to $jobType"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                        Set-ADUser $adUser -Title $teachNumber -Department $jobType
                    }

                    # Check to see if their description (job title) exists in PS, and if it is correct in AD, update if not
                    if (($adUser.description -ne $jobTitle) -and ![string]::IsNullOrEmpty($jobTitle))
                    {
                        $oldDescription = $adUser.description
                        $message = "      ACTION: DESCRIPTION: Updating user $uDCID's description from $oldDescription to $jobTitle"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                        Set-ADUser $adUser -Description $jobTitle
                    }

                    # Check to ensure the user is a member of the papercut staff group
                    if ($papercutStaffMembers -notcontains $adUser.samAccountName)
                    {
                        $message =  "       ACTION: GROUP: User $currentSamAccountName - $email - $jobType is not a member of $papercutGroup, will add them"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                        try 
                        {
                            Add-ADGroupMember -Identity $papercutGroup -Members $adUser.samAccountName # add the user to the group
                        }
                        catch
                        {
                            $message = "     ERROR: Could not add $currentSameAccountName to $papercutGroup"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                        }
                    }

                    # Check to ensure the user is a member of the d118-staff/teachers and school specific staff/teacher groups they are in, ignoring the onboarding, summer school, graduated students, and any of our do not use (DNU) buildings
                    if (($schoolAbbrev -ne "O-HR") -and ($schoolAbbrev -ne "SUM") -and ($schoolAbbrev -notlike "DNU *") -and ($schoolAbbrev -ne "Graduated Students") -and ($schoolAbbrev -ne "AUX"))
                    {
                        if ($jobType -eq "Teacher")
                        {
                            # check the district wide teacher group
                            if ($districtTeacherMembers -notcontains $adUser.samAccountName)
                            {
                                $message =  "       ACTION: GROUP: User $currentSamAccountName - $email - $jobType is not a member of $districtTeacherGroup, will add them"
                                Write-Output $message # write to console
                                $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                                try 
                                {
                                    Add-ADGroupMember -Identity $districtTeacherGroup -Members $adUser.samAccountName # add the user to the group
                                }
                                catch
                                {
                                    $message = "     ERROR: Could not add $currentSameAccountName to $districtTeacherGroup"
                                    Write-Output $message # write to console
                                    $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                                }
                            }
                            # check the school teacher group
                            if ($schoolTeacherMembers -notcontains $adUser.samAccountName)
                            {
                                $message =  "       ACTION: GROUP: User $currentSamAccountName - $email - $jobType is not a member of $schoolTeacherGroup, will add them"
                                Write-Output $message # write to console
                                $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                                try 
                                {
                                    Add-ADGroupMember -Identity $schoolTeacherGroup -Members $adUser.samAccountName # add the user to the group
                                }
                                catch 
                                {
                                    $message = "     ERROR: Could not add $currentSameAccountName to $schoolTeacherGroup"
                                    Write-Output $message # write to console
                                    $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                                }
                            }
                        }
                        elseif (($jobType -eq "Staff") -or ($jobType -eq "Lunch Staff")) # non-teaching staff
                        {
                            # check the district wide staff group
                            if ($districtStaffMembers -notcontains $adUser.samAccountName)
                            {
                                    $message =  "       ACTION: GROUP: User $currentSamAccountName - $email - $jobType is not a member of $districtStaffGroup, will add them"
                                    Write-Output $message # write to console
                                    $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                                    try 
                                    {
                                        Add-ADGroupMember -Identity $districtStaffGroup -Members $adUser.samAccountName # add the user to the group
                                    }
                                    catch 
                                    {
                                        $message = "     ERROR: Could not add $currentSameAccountName to $districtStaffGroup"
                                        Write-Output $message # write to console
                                        $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                                    }
                            }
                            # check the school staff group
                            if ($schoolStaffMembers -notcontains $adUser.samAccountName)
                            {
                                $message =  "       ACTION: GROUP: User $currentSamAccountName - $email - $jobType is not a member of $schoolStaffGroup, will add them"
                                Write-Output $message # write to console
                                $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                                try 
                                {
                                    Add-ADGroupMember -Identity $schoolStaffGroup -Members $adUser.samAccountName # add the user to the group
                                }
                                catch 
                                {
                                    $message = "     ERROR: Could not add $currentSameAccountName to $schoolStaffGroup"
                                    Write-Output $message # write to console
                                    $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                                }
                            }
                        }
                        elseif ($jobType -eq "Substitute")# subs
                        {
                            # check district wide sub group
                            if ($districtSubMembers -notcontains $adUser.samAccountName)
                            {
                                $message =  "       ACTION: GROUP: User $currentSamAccountName - $email - $jobType is not a member of $districtSubGroup, will add them"
                                Write-Output $message # write to console
                                $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                                try 
                                {
                                    Add-ADGroupMember -Identity $districtSubGroup -Members $adUser.samAccountName # add the user to the group
                                }
                                catch 
                                {
                                    $message = "     ERROR: Could not add $currentSameAccountName to $districtSubGroup"
                                    Write-Output $message # write to console
                                    $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                                }
                            }
                        }
                    }

                    # Check to see if the "Full Name" is the same as their samAccountName, if not, change it to match
                    if ($currentFullName -ne $currentSamAccountName)
                    {
                        $message = "      ACTION: FULL NAME: Updating user $uDCID's 'full name' from $currentFullName to $currentSamAccountName"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                        Rename-ADObject $adUser -NewName $currentSamAccountName
                    }

                    # Check to see if they have a homedrive populated, if not we want to assign them one in their current building, but not for the onboarding building
                    if([string]::IsNullOrEmpty($adUser.homedirectory) -and ($schoolAbbrev -ne "O-HR"))
                    {
                        # find their home drive path from the school abbrev
                        switch -Wildcard ($schoolAbbrev)
                        {
                            "??S" {$schoolHomedrive = $schoolAbbrev + "_Teachers$\"}
                            "SSO OFF" {$schoolHomedrive = "SSO_Staff$\"}
                            "CUR OFF" {$schoolHomedrive = "Curriculum_Staff$\"}
                            "TRAN" {$schoolHomedrive = "Transportation_Staff$\"}
                            "MNT" {$schoolHomedrive = "Maintenance_Staff$\"}
                            "CO" {$schoolHomedrive = "District_Office_Home_Drives$\"}
                            Default {$schoolHomedrive = "Other_Staff$\"}
                        }
                        $newHomedirectory = $env:SHARED_DRIVE_BASE_PATH + $schoolHomedrive + $currentSamAccountName
                        $message = "      ACTION: HOMEDRIVE: User $currentSamAccountName - $uDCID's does not have a home directory mapped, will be assigned one at $newHomedirectory"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                        try 
                        {
                            Set-ADUser $adUser -HomeDirectory $newHomedirectory -HomeDrive "H:" # set their home drive to be H: and mapped to the directory constructed from their building and name
                        }
                        catch 
                        {
                            $message =  "          ERROR: Could not map homedrive for $currentSamAccountName to $newHomedirectory"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                            Write-Output $_ # write out the actual error
                            $_ | Out-File -FilePath .\syncLog.txt -Append
                        }
                        
                    }
                }
                # otherwise a user was not found, and we need to create them
                else 
                {
                    $message =  "  ACTION: CREATION: User with DCID $uDCID does not exist, will try to create them as $samAccountName"
                    Write-Output $message # write to console
                    $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                    try
                    {
                        New-ADUser -SamAccountName $samAccountName -Name $samAccountName -DisplayName ($firstName + " " + $lastName) -GivenName $firstName -Surname $lastName -EmailAddress $email -UserPrincipalName $email -Path $OUPath -AccountPassword $defaultPassword -ChangePasswordAtLogon $False -PasswordNeverExpires $true -CannotChangePassword $true -Enabled $true -Title $teachNumber -Department $jobType -Description $jobTitle -OtherAttributes @{'pSuDCID' = $uDCID}
                    }
                    catch
                    {
                        $message =  "       ERROR: Could not create user $samAccountName, trying again with full first name appended"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                        Write-Output $_ # write out the actual error
                        $_ | Out-File -FilePath .\syncLog.txt -Append
                        # add their full first name after a period after the last name
                        $samAccountName = $lastName.ToLower().replace(" ", "-").replace("'", "") + "." + $firstName.ToLower().replace(" ", "-").replace("'", "")
                        try
                        {
                            New-ADUser -SamAccountName $samAccountName -Name $samAccountName -DisplayName ($firstName + " " + $lastName) -GivenName $firstName -Surname $lastName -EmailAddress $email -UserPrincipalName $email -Path $OUPath -AccountPassword $defaultPassword -ChangePasswordAtLogon $False -PasswordNeverExpires $true -CannotChangePassword $true -Enabled $true -Title $teachNumber -Department $jobType -Description $jobTitle -OtherAttributes @{'pSuDCID' = $uDCID}
                        }
                        catch
                        {
                            $message =  "       ERROR: Could not create user $samAccountName, stopping"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                            Write-Output $_ # write out the actual error
                            $_ | Out-File -FilePath .\syncLog.txt -Append
                        }
                    }
                }
            }
            # start the inactive staff block, should be disabled and moved to the suspended accounts OU. We dont care otherwise about incorrect info
            else
            {
                $OUPath = "OU=Staff,OU=SUSPENDED ACCOUNTS,$ConstantOU"
                $properDistinguised = "CN=$samAccountName,$OUPath"
                $adUser = Get-ADUser -Filter "pSuDCID -eq $uDCID" # do a query for existing users with the custom attribute pSuDCID that equals the users DCID
                if ($adUser)
                { # if we find a user with a matchind DCID, just update their info
                    $currentFullName = $adUser.name
                    $properDistinguised = "CN=$currentFullName,$OUPath"
                    $currentSamAccountName = $adUser.SamAccountName
                    $message = "  User with DCID $uDCID already exists under $currentSamAccountName, ensuring they are suspended"
                    Write-Output $message # write to console
                    $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                    # check and see if their account is in the right OU
                    # check to see if the account is enabled, if so we need to disable it
                    if ($adUser.Enabled)
                    {
                        try 
                        {
                            $message = "      ACTION: SUSPENDED DISABLE: Disabling user $currentSamAccountName - $uDCID - $email"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\syncLog.txt -Append # write to log file
                            Disable-ADAccount $adUser # disables the selected account
                        }
                        catch 
                        {
                            $message =  "          ERROR: Could not suspend $currentSamAccountName - $uDCID - $currentSamAccountName"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                            Write-Output $_ # write out the actual error
                            $_ | Out-File -FilePath .\syncLog.txt -Append
                        }
                        
                    }
                    if ($properDistinguised -ne $adUser.DistinguishedName)
                    {
                        try 
                        {
                            $message = "      ACTION: SUSPENDED OU: Moving user $currentSamAccountName - $uDCID - $email to the Suspended Users Staff OU"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                            Move-ADObject $adUser -TargetPath $OUPath # moves the targeted AD user account to the correct suspended accounts OU
                        }
                        catch 
                        {
                            $message =  "          ERROR: Could not move $uDCID - $currentSamAccountName to the suspended users - staff OU"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                            Write-Output $_ # write out the actual error
                            $_ | Out-File -FilePath .\syncLog.txt -Append
                        }
                        
                    }
                }
                else
                {
                    $message = "  WARNING: Found inactive user DCID $uDCID without matching AD account. Should be $samAccountName"
                    Write-Output $message # write to console
                    $message | Out-File -FilePath .\syncLog.txt -Append # write to log file 
                }
            }
        }
        else # otherwise if their name was found in the bad names list, just give a warning
        {
            $message = "INFO: found user matching name in bad names list: $firstName $LastName"
            Write-Output $message
            $message | Out-File -FilePath .\syncLog.txt -Append
        }
    }
}

# repadmin.exe /syncall D118-DIST-OFF /Aed

