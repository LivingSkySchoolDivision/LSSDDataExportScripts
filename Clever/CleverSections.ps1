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
                DefaultUserStaff.UF_2085 AS Teacher_id,
                (SELECT UserStaff.UF_2085 FROM ClassResource LEFT OUTER JOIN Staff ON ClassResource.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID WHERE ClassResource.iClassID=Class.iClassID AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY iClassResourceID OFFSET 0 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_2_id,
                (SELECT UserStaff.UF_2085 FROM ClassResource LEFT OUTER JOIN Staff ON ClassResource.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID WHERE ClassResource.iClassID=Class.iClassID AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY iClassResourceID OFFSET 1 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_3_id,
                (SELECT UserStaff.UF_2085 FROM ClassResource LEFT OUTER JOIN Staff ON ClassResource.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID WHERE ClassResource.iClassID=Class.iClassID AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY iClassResourceID OFFSET 2 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_4_id,
                (SELECT UserStaff.UF_2085 FROM ClassResource LEFT OUTER JOIN Staff ON ClassResource.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID WHERE ClassResource.iClassID=Class.iClassID AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY iClassResourceID OFFSET 3 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_5_id,
                (SELECT UserStaff.UF_2085 FROM ClassResource LEFT OUTER JOIN Staff ON ClassResource.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID WHERE ClassResource.iClassID=Class.iClassID AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY iClassResourceID OFFSET 4 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_6_id,
                (SELECT UserStaff.UF_2085 FROM ClassResource LEFT OUTER JOIN Staff ON ClassResource.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID WHERE ClassResource.iClassID=Class.iClassID AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY iClassResourceID OFFSET 5 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_7_id,
                (SELECT UserStaff.UF_2085 FROM ClassResource LEFT OUTER JOIN Staff ON ClassResource.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID WHERE ClassResource.iClassID=Class.iClassID AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY iClassResourceID OFFSET 6 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_8_id,
                (SELECT UserStaff.UF_2085 FROM ClassResource LEFT OUTER JOIN Staff ON ClassResource.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID WHERE ClassResource.iClassID=Class.iClassID AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY iClassResourceID OFFSET 7 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_9_id,
                (SELECT UserStaff.UF_2085 FROM ClassResource LEFT OUTER JOIN Staff ON ClassResource.iStaffID=Staff.iStaffID LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID WHERE ClassResource.iClassID=Class.iClassID AND Staff.iStaffID<>Class.iDefault_StaffID ORDER BY iClassResourceID OFFSET 8 ROWS FETCH NEXT 1 ROWS ONLY) AS Teacher_10_id,
                Class.cName AS Name,
                Class.cSection AS Section_number,
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
                LEFT OUTER JOIN Staff as DefaultStaff ON Class.iDefault_StaffID=DefaultStaff.iStaffID
                LEFT OUTER JOIN UserStaff as DefaultUserStaff ON DefaultStaff.iStaffID=DefaultUserStaff.iStaffID
            WHERE
                Class.iDefault_StaffID > 0
                AND (SELECT COUNT(iEnrollmentID) FROM Enrollment WHERE iClassID=Class.iClassID) > 0
                AND (LEN(DefaultUserStaff.UF_2085) > 0)
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