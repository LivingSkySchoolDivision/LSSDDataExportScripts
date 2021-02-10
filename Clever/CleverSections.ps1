param (
    [Parameter(Mandatory=$true)][string]$OutputFileName,
    [string]$ConfigFilePath
 )

##############################################
# About this file
##############################################
# The "teachers.csv" file should only include teachers who have classes associated with them.
# Staff that are duplicates (for multi-school purposes) should be skipped

##############################################
# Script configuration                       #
##############################################

# SQL Query to run
# The output CSV file will use column names from your SQL query.
# Rename them using "as" - example: "SELECT cFirstName as FirstName FROM Students"
$SqlQuery = "SELECT
                School.cCode AS School_id,
                Class.iClassID AS Section_id,                
				UserStaff.UF_2085 AS Teacher_id,
				(SELECT CASE WHEN UserStaff.iBaseStaffIDid > 0 THEN BaseUserStaff.UF_2085 ELSE UserStaff.UF_2085 END FROM (SELECT iStaffID FROM ClassResource WHERE ClassResource.iClassID=Class.iClassID UNION SELECT iDefault_StaffID FROM Class as c WHERE c.iClassID=Class.iClassID) as t LEFT OUTER JOIN Staff ON t.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID LEFT OUTER JOIN UserStaff AS BaseUserStaff ON UserStaff.iBaseStaffIDid=BaseUserStaff.iStaffID WHERE Staff.iStaffID<>0 AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY Staff.iStaffID OFFSET 0 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_2_id,
                (SELECT CASE WHEN UserStaff.iBaseStaffIDid > 0 THEN BaseUserStaff.UF_2085 ELSE UserStaff.UF_2085 END FROM (SELECT iStaffID FROM ClassResource WHERE ClassResource.iClassID=Class.iClassID UNION SELECT iDefault_StaffID FROM Class as c WHERE c.iClassID=Class.iClassID) as t LEFT OUTER JOIN Staff ON t.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID LEFT OUTER JOIN UserStaff AS BaseUserStaff ON UserStaff.iBaseStaffIDid=UserStaff.iStaffID WHERE Staff.iStaffID<>0 AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY Staff.iStaffID OFFSET 1 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_3_id,
                (SELECT CASE WHEN UserStaff.iBaseStaffIDid > 0 THEN BaseUserStaff.UF_2085 ELSE UserStaff.UF_2085 END FROM (SELECT iStaffID FROM ClassResource WHERE ClassResource.iClassID=Class.iClassID UNION SELECT iDefault_StaffID FROM Class as c WHERE c.iClassID=Class.iClassID) as t LEFT OUTER JOIN Staff ON t.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID LEFT OUTER JOIN UserStaff AS BaseUserStaff ON UserStaff.iBaseStaffIDid=UserStaff.iStaffID WHERE Staff.iStaffID<>0 AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY Staff.iStaffID OFFSET 2 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_4_id,
                (SELECT CASE WHEN UserStaff.iBaseStaffIDid > 0 THEN BaseUserStaff.UF_2085 ELSE UserStaff.UF_2085 END FROM (SELECT iStaffID FROM ClassResource WHERE ClassResource.iClassID=Class.iClassID UNION SELECT iDefault_StaffID FROM Class as c WHERE c.iClassID=Class.iClassID) as t LEFT OUTER JOIN Staff ON t.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID LEFT OUTER JOIN UserStaff AS BaseUserStaff ON UserStaff.iBaseStaffIDid=UserStaff.iStaffID WHERE Staff.iStaffID<>0 AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY Staff.iStaffID OFFSET 3 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_5_id,
                (SELECT CASE WHEN UserStaff.iBaseStaffIDid > 0 THEN BaseUserStaff.UF_2085 ELSE UserStaff.UF_2085 END FROM (SELECT iStaffID FROM ClassResource WHERE ClassResource.iClassID=Class.iClassID UNION SELECT iDefault_StaffID FROM Class as c WHERE c.iClassID=Class.iClassID) as t LEFT OUTER JOIN Staff ON t.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID LEFT OUTER JOIN UserStaff AS BaseUserStaff ON UserStaff.iBaseStaffIDid=UserStaff.iStaffID WHERE Staff.iStaffID<>0 AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY Staff.iStaffID OFFSET 4 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_6_id,
                (SELECT CASE WHEN UserStaff.iBaseStaffIDid > 0 THEN BaseUserStaff.UF_2085 ELSE UserStaff.UF_2085 END FROM (SELECT iStaffID FROM ClassResource WHERE ClassResource.iClassID=Class.iClassID UNION SELECT iDefault_StaffID FROM Class as c WHERE c.iClassID=Class.iClassID) as t LEFT OUTER JOIN Staff ON t.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID LEFT OUTER JOIN UserStaff AS BaseUserStaff ON UserStaff.iBaseStaffIDid=UserStaff.iStaffID WHERE Staff.iStaffID<>0 AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY Staff.iStaffID OFFSET 5 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_7_id,
                (SELECT CASE WHEN UserStaff.iBaseStaffIDid > 0 THEN BaseUserStaff.UF_2085 ELSE UserStaff.UF_2085 END FROM (SELECT iStaffID FROM ClassResource WHERE ClassResource.iClassID=Class.iClassID UNION SELECT iDefault_StaffID FROM Class as c WHERE c.iClassID=Class.iClassID) as t LEFT OUTER JOIN Staff ON t.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID LEFT OUTER JOIN UserStaff AS BaseUserStaff ON UserStaff.iBaseStaffIDid=UserStaff.iStaffID WHERE Staff.iStaffID<>0 AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY Staff.iStaffID OFFSET 6 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_8_id,
                (SELECT CASE WHEN UserStaff.iBaseStaffIDid > 0 THEN BaseUserStaff.UF_2085 ELSE UserStaff.UF_2085 END FROM (SELECT iStaffID FROM ClassResource WHERE ClassResource.iClassID=Class.iClassID UNION SELECT iDefault_StaffID FROM Class as c WHERE c.iClassID=Class.iClassID) as t LEFT OUTER JOIN Staff ON t.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID LEFT OUTER JOIN UserStaff AS BaseUserStaff ON UserStaff.iBaseStaffIDid=UserStaff.iStaffID WHERE Staff.iStaffID<>0 AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY Staff.iStaffID OFFSET 7 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_9_id,
                (SELECT CASE WHEN UserStaff.iBaseStaffIDid > 0 THEN BaseUserStaff.UF_2085 ELSE UserStaff.UF_2085 END FROM (SELECT iStaffID FROM ClassResource WHERE ClassResource.iClassID=Class.iClassID UNION SELECT iDefault_StaffID FROM Class as c WHERE c.iClassID=Class.iClassID) as t LEFT OUTER JOIN Staff ON t.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID LEFT OUTER JOIN UserStaff AS BaseUserStaff ON UserStaff.iBaseStaffIDid=UserStaff.iStaffID WHERE Staff.iStaffID<>0 AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY Staff.iStaffID OFFSET 8 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_10_id,
                REPLACE(REPLACE(Class.cName, CHAR(13), ''), CHAR(10), '') AS Name,
                REPLACE(REPLACE(Class.cSection, CHAR(13), ''), CHAR(10), '') AS Section_number,
                CASE
                    WHEN Grades.cName = '0k' THEN 'Kindergarten'
                    WHEN Grades.cName = 'pk' THEN 'PreKindergarten'
                    ELSE CAST(CAST(LTRIM(RTRIM(Grades.cName)) AS INT) AS VARCHAR)
                END AS Grade,
                Course.cName AS Course_name,
                Course.cGovernmentCode AS Course_number,
                '' AS Course_description,
                '' AS Period,
                '' AS Subject,
                '' AS Term_name,
                '' AS Term_start,
                '' AS Term_end
            FROM
                Class
                LEFT OUTER JOIN School ON Class.iSchoolID=School.iSchoolID
                LEFT OUTER JOIN Course ON Class.iCourseID=Course.iCourseID
                LEFT OUTER JOIN Grades ON Class.iLow_GradesID=Grades.iGradesID
				LEFT OUTER JOIN UserStaff ON Class.iDefault_StaffID=UserStaff.iStaffID
            WHERE
                (SELECT COUNT(iEnrollmentID) FROM Enrollment WHERE iClassID=Class.iClassID AND Enrollment.iLV_CompletionStatusID=0) > 0
                AND (SELECT COUNT(*) FROM (SELECT iStaffID FROM ClassResource WHERE ClassResource.iClassID=Class.iClassID UNION SELECT iDefault_StaffID FROM Class as c WHERE c.iClassID=Class.iClassID) x WHERE x.iStaffID <> 0 AND x.iStaffID <> '' AND x.iStaffID IS NOT NULL)  > 0
				AND UserStaff.UF_2085 <> ''
            ORDER BY
                Class.iClassID
"

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
