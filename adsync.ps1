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

$queryStatment = "SELECT lastfirst, email_addr, users_dcid FROM teachers ORDER BY users_dcid" #Be careful not to terminate it with a semicolon, it doesn't like it
#Actual Code
#Load Required Types and modules
Add-Type -Path $OracleDLLPath
Import-Module SqlServer
#Create the connection string
$connectionstring = 'User Id=' + $username + ';Password=' + $password + ';Data Source=' + $datasource 
#Create the connection object
$con = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionstring)
#Create a command and configure it
$cmd = $con.CreateCommand()
$cmd.CommandText = $queryStatment
$cmd.CommandTimeout = 3600 #Seconds
$cmd.FetchSize = 10000000 #10MB
#Creates a data adapter for the command
$da = New-Object Oracle.ManagedDataAccess.Client.OracleDataAdapter($cmd);
#The Data adapter will fill this DataTable
$resultSet = New-Object System.Data.DataTable # each row is one entry, the columns can be accessed with []
#Only here the query is sent and executed in Oracle 
[void]$da.fill($resultSet)
#Close the connection
$con.Close()


foreach ($result in $resultSet){
    Write-Output $result[1]
}

Import-Module ActiveDirectory