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
                HR.iSchoolID AS SchoolID,
                CONCAT(HR.iSchoolID,'-',HR.iHomeroomID) AS SectionGUID,
                LEFT(HR.iHomeroomID,5) AS SectionID,
                HR.cName AS SubSection,
                REPLACE(REPLACE(REPLACE(STUFF((SELECT DISTINCT iTermID FROM Term TE2 WHERE TE2.iTrackID = T.iTrackID FOR XML PATH ('')) , 1,1,''),'ITERMID>',''),'</',''),'<',',') AS TermID,
                'SK.NAC' AS CourseID,
                '' AS TeacherGUID,
                REPLACE(REPLACE(REPLACE(STUFF((SELECT DISTINCT iRoomID FROM Homeroom HR2 WHERE HR2.iHomeroomID = HR.iHomeroomID AND HR2.iRoomID > 0 FOR XML PATH ('')) , 1,1,''),'iRoomID>',''),'</',''),'<',',') AS RoomID,
                '' AS GradeLevel,
                'Homeroom Attendance' AS Subject,
                '814' AS CourseCode,
                'Homeroom ' + HR.cName AS CourseTitle,
                '1' Attendance,
                '0' ScheduleMode,
                T.iTrackID AS ScheduleID,
                '' AS LOWGRADE,
                '' AS HIGHGRADE
            FROM 
                Homeroom HR
                INNER JOIN Student S ON HR.iHomeroomID = S.iHomeroomID
                INNER JOIN StudentStatus SS ON S.iStudentID = SS.iStudentID
                LEFT OUTER JOIN ROOM R ON HR.iRoomID = R.iRoomID
                INNER JOIN Grades G ON S.iGradesID = G.iGradesID
                INNER JOIN Track T ON S.iTrackID = T.iTrackID
                INNER JOIN TERM ELMTERM ON T.iTrackID = ELMTERM.iTrackID
            WHERE
                T.lDaily = 1 AND
                (SS.dInDate <= getDate() + 1) AND
                ((SS.dOutDate < '1901-01-01') OR (SS.dOutDate >=  { fn CURDATE() })) 

            UNION 
                ALL
                
            SELECT DISTINCT
                C.iSchoolID AS SchoolID,
                CONCAT(C.iSchoolID,'-',C.iClassID) AS SectionGUID,
                LEFT(C.iClassID,5) AS SectionID,
                C.cSection AS SubSection,
                CASE WHEN
                    CR.iClassResourceID NOT IN (SELECT iClassResourceID FROM ClassSchedule)
                THEN
                    REPLACE(REPLACE(REPLACE(STUFF((SELECT DISTINCT iTermID FROM Term TE2 WHERE TE2.iTrackID = ELMTERM.ITRACKID FOR XML PATH ('')) , 1,1,''),'ITERMID>',''),'</',''),'<',',')
                ELSE
                    REPLACE(REPLACE(REPLACE(STUFF((SELECT DISTINCT iTermID FROM ClassSchedule CS2 WHERE CS2.iClassResourceID = CS.iClassResourceID FOR XML PATH ('')) , 1,1,''),'ITERMID>',''),'</',''),'<',',')
                END AS TermID,
                CASE WHEN
                    CO.cGovernmentCode != 'NAC' AND CO.cGovernmentCode < 999 
                THEN 
                    'SK.0' + CO.cGovernmentCode 
                ELSE 
                    'SK.' + CO.cGovernmentCode 
                END AS CourseID,
                '' AS TeacherGUID,
                REPLACE(REPLACE(REPLACE(STUFF((SELECT DISTINCT iRoomID FROM ClassResource CR2 WHERE CR2.iClassID = CR.iClassID AND CR2.iRoomID > 0 FOR XML PATH ('')) , 1,1,''),'iRoomID>',''),'</',''),'<',',') AS RoomID,
                '' AS GradeLevel,
                SUB.cName AS Subject,
                CO.iCourseID AS CourseCode,
                CO.cName AS CourseTitle,
                CASE WHEN T.lDaily = 0 THEN 1  ELSE 0 END AS Attendance,
                '0' AS ScheduleMode,
                T.iTrackID AS ScheduleID,
                LTRIM(RTRIM(LOW.cName)) AS LOWGRADE,
                LTRIM(RTRIM(HIGH.cName)) AS HIGHGRADE
            FROM 
                Class C
                INNER JOIN ClassResource CR ON C.iClassID = CR.iClassID
                LEFT OUTER JOIN ClassSchedule CS ON CR.iClassResourceID = CS.iClassResourceID
                LEFT OUTER JOIN ROOM R ON CR.iRoomID = R.iRoomID
                INNER JOIN Grades LOW ON C.iLow_GradesID = LOW.iGradesID
                INNER JOIN Grades HIGH ON C.iHigh_GradesID = HIGH.iGradesID
                INNER JOIN Course CO ON C.iCourseID = CO.iCourseID
                LEFT OUTER JOIN LookupValues SUB ON CO.iLV_SubjectID = SUB.iLookupValuesID
                INNER JOIN Track T ON C.iTrackID = T.iTrackID
                INNER JOIN TERM ELMTERM ON T.iTrackID = ELMTERM.iTrackID
            WHERE
                C.iLV_SessionID != '4720' --Session set to No Edsby;"

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


