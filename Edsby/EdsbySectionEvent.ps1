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
                HR.iSchoolID AS SchoolID,
                CONCAT(HR.iSchoolID,'-',HR.iHomeroomID) AS SectionGUID,
                D.iDaysID AS DayID,
                CASE WHEN HR.iRoomID = 0 THEN NULL ELSE HR.iRoomID END AS RoomID,
                T.iTrackID AS ScheduleID,
                TE.iTermID AS TermID,
                AB.iAttendanceBlocksID AS PeriodID,
                'T' AS TakeAttendance,
                '' AS TeacherGUID
            FROM 
                Homeroom HR
                INNER JOIN Student S ON HR.iHomeroomID = S.iHomeroomID
                INNER JOIN Track T ON S.iTrackID = T.iTrackID
                LEFT OUTER JOIN Days D ON T.iTrackID =D.iTrackID
                INNER JOIN Term TE ON T.iTrackID = TE.iTrackID
                INNER JOIN AttendanceBlocks AB ON T.iTrackID = AB.iTrackID
            WHERE 
                T.lDaily = 1
            
            UNION 
                ALL

            SELECT DISTINCT
                C.iSchoolID AS SchoolID,
                CONCAT(C.iSchoolID,'-',C.iClassID) AS SectionGUID,
                D.iDaysID AS DayID,
                CASE WHEN CR.iRoomID = 0 THEN NULL ELSE R.iRoomID END AS RoomID,
                T.iTrackID AS ScheduleID,
                CS.iTermID AS TermID,
                B.IBlocksID AS PeriodID,
                'T' AS TakeAttendance,
                '' AS TeacherGUID
            FROM 
                Class C
                INNER JOIN ClassResource CR ON C.iClassID = CR.iClassID
                LEFT OUTER JOIN ROOM R ON R.iRoomID = (SELECT TOP 1(iRoomID) FROM ClassResource WHERE iClassID = C.iClassID)
                LEFT OUTER JOIN ClassSchedule CS ON CR.iClassResourceID = CS.iClassResourceID
                INNER JOIN Track T ON C.iTrackID = T.iTrackID
                INNER JOIN Days D ON T.iTrackID = D.iTrackID AND CS.iDayNumber = D.iDayNumber
                INNER JOIN Blocks B ON cs.iBlockNumber = b.iBlockNumber and t.iTrackID = b.iTrackID
            WHERE 
                T.lDaily = 0;"

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