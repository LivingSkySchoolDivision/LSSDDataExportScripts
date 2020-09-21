param (
    [Parameter(Mandatory=$true)][string]$OutputFileName,
    [string]$ConfigFilePath
 )

##############################################
# Script configuration                       #
##############################################

# SQL Query to run
# The output CSV file will use column names from your SQL query.
# Rename them using "as" - example: "SELECT cFirstName as FirstName FROM Students"
$SqlQuery = "SELECT DISTINCT
                S.iSchoolID AS SchoolID,
                T.iTrackID AS ScheduleID,
                'REG' AS BSID,
                AB.iAttendanceBlocksID AS PeriodID,
                LTRIM(RTRIM(AB.cName)) AS PeriodName,
                FORMAT(AB.tStartTime,'HH:mm') AS stime,
                FORMAT(AB.tEndTime,'HH:mm') AS etime
            FROM 
                School S
                INNER JOIN Track T ON S.iSchoolID = T.iSchoolID
                LEFT OUTER JOIN Term TE ON T.iTrackID = TE.iTrackID
                INNER JOIN AttendanceBlocks AB ON T.iTrackID = AB.iTrackID

            UNION
                ALL

            SELECT DISTINCT
                S.iSchoolID AS SchoolID,
                T.iTrackID AS ScheduleID,
                'ALT' AS BSID,
                AB.iAttendanceBlocksID AS PeriodID,
                LTRIM(RTRIM(AB.cName)) AS PeriodName,
                FORMAT(AB.tStartTime1,'HH:mm') AS stime,
                FORMAT(AB.tEndTime1,'HH:mm') AS etime
            FROM 
                School S
                INNER JOIN Track T ON S.iSchoolID = T.iSchoolID
                LEFT OUTER JOIN Term TE ON T.iTrackID = TE.iTrackID
                INNER JOIN AttendanceBlocks AB ON T.iTrackID = AB.iTrackID
            WHERE 
                AB.tStartTime1 != '1900-01-01 00:00:00.000' AND AB.iInstructionalMinutes1 !=0

            UNION 
                ALL

            SELECT DISTINCT
                S.iSchoolID AS SchoolID,
                T.iTrackID AS ScheduleID,
                'REG' AS BSID,
                B.IBlocksID AS PeriodID,
                LTRIM(RTRIM(B.cName)) AS PeriodName,
                FORMAT(B.tStartTime,'HH:mm') AS stime,
                FORMAT(B.tEndTime,'HH:mm') AS etime
            FROM 
                School S
                INNER JOIN Track T ON S.iSchoolID = T.iSchoolID
                LEFT OUTER JOIN Term TE ON T.iTrackID = TE.iTrackID
                INNER JOIN Blocks B ON T.iTrackID = B.iTrackID                

            UNION
                ALL

            SELECT DISTINCT
                S.iSchoolID AS SchoolID,
                T.iTrackID AS ScheduleID,
                'ALT' AS BSID,
                B.IBlocksID AS PeriodID,
                LTRIM(RTRIM(B.cName)) AS PeriodName,
                FORMAT(B.tStartTime1,'HH:mm') AS stime,
                FORMAT(B.tEndTime1,'HH:mm') AS etime
            FROM 
                School S
                INNER JOIN Track T ON S.iSchoolID = T.iSchoolID
                LEFT OUTER JOIN Term TE ON T.iTrackID = TE.iTrackID
                INNER JOIN Blocks B ON T.iTrackID = B.iTrackID
            WHERE 
                B.tStartTime1 != '1900-01-01 00:00:00.000' AND B.iInstructionalMinutes1 !=0;"

# CSV Delimeter
# Some systems expect this to be a tab "`t" or a pipe "|".
$Delimeter = ','

# Should all columns be quoted, or just those that contains characters to escape?
# Note: This has no effect on systems with PowerShell versions <7.0 (all fields will be quoted otherwise)
$QuoteAllColumns = $false

##############################################
# No configurable settings beyond this point #
##############################################

# Find the config file
$AdjustedConfigFilePath = $ConfigFilePath
if ($AdjustedConfigFilePath.Length -le 0) {
    $AdjustedConfigFilePath = join-path -Path $(Split-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) -Parent) -ChildPath "config.xml"
}

# Retreive the connection string from config.xml
if ((test-path -Path $AdjustedConfigFilePath) -eq $false) {
    Throw "Config file not found. Specify using -ConfigFilePath. Defaults to config.xml in the directory above where this script is run from."
}
$configXML = [xml](Get-Content $AdjustedConfigFilePath)
$ConnectionString = $configXML.Settings.SchoolLogic.ConnectionString

# Set up the SQL connection
$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $ConnectionString
$SqlCommand = New-Object System.Data.SqlClient.SqlCommand
$SqlCommand.CommandText = $SqlQuery
$SqlCommand.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCommand
$SqlDataSet = New-Object System.Data.DataSet

# Run the SQL query
$SqlConnection.open()
$SqlAdapter.Fill($SqlDataSet)
$SqlConnection.close()

# Output to a CSV file
foreach($DSTable in $SqlDataSet.Tables) {
    if (($PSVersionTable.PSVersion.Major -ge 7) -and ($QuoteAllColumns -eq $false)) {
        $DSTable | export-csv $OutputFileName -notypeinformation -Delimiter $Delimeter -UseQuotes AsNeeded
    } else {        
        $DSTable | export-csv $OutputFileName -notypeinformation -Delimiter $Delimeter
    }
}