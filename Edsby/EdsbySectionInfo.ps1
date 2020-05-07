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

##### IMPORTANT
# If you modify this query, the post-processing code below may need to be modified as well.
##### IMPORTANT
$SqlQuery = "SELECT DISTINCT
                C.iSchoolID AS SchoolID,
                CONCAT(C.iSchoolID,'-',C.iClassID) AS SectionGUID,
                LEFT(C.iClassID, 3) + RIGHT(C.iClassID, 2) AS SectionID,
                C.cSection AS SubSection,
                TE.iTermID AS TermID,
                CO.cGovernmentCode AS CourseID,
                '' AS TeacherGUID,
                CASE WHEN CR.iRoomID = 0 THEN NULL ELSE R.iRoomID END AS RoomID,
                '' AS GradeLevel,
                SUB.cName AS Subject,
                CO.iCourseID AS CourseCode,
                CO.cName AS CourseTitle,
                CASE WHEN T.lDaily = 1 THEN 0 ELSE 1 END AS Attendance,
                CASE WHEN T.lDaily = 1 THEN 4 ELSE 0 END AS ScheduleMode,
                T.iTrackID AS ScheduleID,
                LTRIM(RTRIM(LOW.cName)) AS LOWGRADE,
                LTRIM(RTRIM(HIGH.cName)) AS HIGHGRADE
            FROM Class C
                INNER JOIN ClassResource CR ON C.iClassID = CR.iClassID
                LEFT OUTER JOIN ClassSchedule CS ON CR.iClassResourceID = CS.iClassResourceID
                LEFT OUTER JOIN ROOM R ON R.iRoomID = (SELECT TOP 1(iRoomID) FROM ClassResource WHERE iClassID = C.iClassID)
                INNER JOIN Grades LOW ON C.iLow_GradesID = LOW.iGradesID
                INNER JOIN Grades HIGH ON C.iHigh_GradesID = HIGH.iGradesID
                INNER JOIN Course CO ON C.iCourseID = CO.iCourseID
                LEFT OUTER JOIN LookupValues SUB ON CO.iLV_SubjectID = SUB.iLookupValuesID
                INNER JOIN Track T ON C.iTrackID = T.iTrackID
                LEFT OUTER JOIN TERM TE ON T.iTrackID = TE.iTrackID AND CS.iTermID = TE.iTermID
            ORDER BY
                CONCAT(C.iSchoolID,'-',C.iClassID);"

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

# Post-processing - format the grade range
function Get-GradesBetween  {
    param (
        [Parameter(Mandatory=$true)][string]$first,
        [Parameter(Mandatory=$true)][string]$second
    )
    $returnMe = ""

    $first = $first.Trim()
    $second = $second.Trim()

    $includeK = $false
    $includePK = $false

    if (($first.ToLower() -eq "pk") -AND ($second.ToLower() -eq "pk")) {
        return "PK"
    }  
    if (($first.ToLower() -eq "0k") -AND ($second.ToLower() -eq "0k")) {
        return "K"
    }
    if (($first.ToLower() -eq "pk") -AND ($second.ToLower() -eq "0k")) {
        return "PK,K"
    }
    if (($first.ToLower() -eq "0k") -AND ($second.ToLower() -eq "pk")) {
        return "PK,K"
    }
       
    # If $first is k or pk, set that aside and set it to 1
    if ($first.ToLower() -eq "pk") {
        $first = 1
        $includePK = $true
    }

    if ($first.ToLower() -eq "0k") {
       $first = 1
       $includeK = $true
    }

    if ($second.ToLower() -eq "pk") {        
        $second = 1
        $includePK = $true
    }
    
    if ($second.ToLower() -eq "0k") {
        $second = 1
        $includeK = $true
    }

    # cast to integers
    $firstNum = [int]$first.Trim()
    $secondNum = [int]$second.Trim()

    if ($firstNum -gt $secondNum) { 
        $tempNum = $firstNum
        $firstNum = $secondNum
        $secondNum = $tempNum
    }

    foreach($x in $firstNum..$secondNum) {
        if ($returnMe.Length -gt 0) {
            $returnMe = "$returnMe,$x"
        } else {
            $returnMe = "$x"
        }
    }

    if ($includeK -eq $true) {
        $returnMe = "K,$returnMe"
    }

    if ($includePK -eq $true) {
        $returnMe = "PK,$returnMe"
    }

    return $returnMe
}

foreach($DSTable in $SqlDataSet.Tables) {
    foreach($DataRow in $DSTable) {
        $DataRow["GradeLevel"] = Get-GradesBetween -first $DataRow["LOWGRADE"] -second $DataRow["HIGHGRADE"]
        $DataRow["LOWGRADE"] = ""
        $DataRow["HIGHGRADE"] = ""
    }
    $DSTable.Columns.Remove("LOWGRADE")
    $DSTable.Columns.Remove("HIGHGRADE")
}

# Output to a CSV file
foreach($DSTable in $SqlDataSet.Tables) {
    if (($PSVersionTable.PSVersion.Major -ge 7) -and ($QuoteAllColumns -eq $false)) {
        $DSTable | export-csv $OutputFileName -notypeinformation -Delimiter $Delimeter -UseQuotes AsNeeded
    } else {        
        $DSTable | export-csv $OutputFileName -notypeinformation -Delimiter $Delimeter
    }
}


