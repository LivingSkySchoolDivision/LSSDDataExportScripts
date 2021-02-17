param (
    [Parameter(Mandatory=$true)][string]$OutputFileName,
    [switch]$JustPeriodAttendance,
    [switch]$JustDailyAttendance,
    [string]$DateFrom,
    [string]$DateTo,
    [string]$SchoolIDs,
    [string]$SchoolDANs,
    [string]$ConfigFilePath
 )

##############################################
# Default settings
##############################################

# Date for which to get attendance for
$PARAM_DateFrom = [DateTime]::Today;
$PARAM_DateTo = [DateTime]::Today.AddDays(1);

if ($DateFrom -ne $null) {
    try {
    $PARAM_DateFrom = [datetime]::Parse($DateFrom)
    } catch {}
}

if ($DateTo -ne $null) {
    try {
    $PARAM_DateTo = [datetime]::Parse($DateTo)
     } catch {}
}

$SchoolIDList_Split = @();

foreach($item in $SchoolIDs.Split(","))
{
    try {
    $id_int = [int]::Parse($item)
    if ($id_int -gt 0) {
        if ($SchoolIDList_Split.Contains($id_int) -eq $false) {
            $SchoolIDList_Split += ($id_int);
        }        
    }
    } catch {}
}

$PARAM_SchoolIDList = [string]::Join(",", $SchoolIDList_Split)


$SchoolDANList_Split = @();

foreach($item in $SchoolDANs.Split(","))
{
    try {
    $id_int = [int]::Parse($item)
    if ($id_int -gt 0) {
        if ($SchoolDANList_Split.Contains($id_int) -eq $false) {
            $SchoolDANList_Split += ($id_int);
        }        
    }
    } catch {}
}

$PARAM_SchoolDANList= [string]::Join(",", $SchoolDANList_Split)

##############################################
# Script configuration                       #
##############################################

# SQL Query to run
# The output CSV file will use column names from your SQL query.
# Rename them using "as" - example: "SELECT cFirstName as FirstName FROM Students"
$SqlQuery = "SELECT
                AttendanceToday.iSchoolID as SchoolID,
                CONCAT('STUDENT-',AttendanceToday.iStudentID) as StudentID,
                Student.cLastName as StudentLastName,
                Student.cFirstName as StudentFirstName,
                Location.cPhone as HomePhone,
                REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(RTRIM(LTRIM(Grades.cName)),'0K','K'),'01','1'),'02','2'),'03','3'),'04','4'),'05','5'),'06','6'),'07','7'),'08','8'),'09','9') as Grade,
                '02/08/2021' as AbsenceDate,
                (SELECT DISTINCT(CONVERT(VarChar(2), A_i.iBlockNumber))
                        FROM Attendance as A_i
                        WHERE A_i.iStudentID=Student.iStudentID
                        AND A_i.dDate >= '$($PARAM_DateFrom.ToString("yyyy-MM-dd")) 00:00:00'
                        AND A_i.dDate <= '$($PARAM_DateTo.ToString("yyyy-MM-dd")) 00:00:00'
                        FOR XML PATH('')) as PeriodsMissed
                FROM(	SELECT 
                            iSchoolID, iStudentID
                        FROM 
                            Attendance 
                            LEFT OUTER JOIN AttendanceStatus ON Attendance.iAttendanceStatusID=AttendanceStatus.iAttendanceStatusID 
                        WHERE 
                            dDate >= '$($PARAM_DateFrom.ToString("yyyy-MM-dd")) 00:00:00'
                            AND dDate <= '$($PARAM_DateTo.ToString("yyyy-MM-dd")) 00:00:00'
                            AND AttendanceStatus.cName = 'Absent'
                            AND Attendance.iAttendanceReasonsID=0
                        GROUP BY iSchoolID, iStudentID
                    ) as AttendanceToday 
                    LEFT OUTER JOIN Student ON AttendanceToday.iStudentID=Student.iStudentID
                    LEFT OUTER JOIN Grades ON Student.iGradesID=Grades.iGradesID
                    LEFT OUTER JOIN Location ON Student.iLocationID=Location.iLocationID
                    LEFT OUTER JOIN Track ON Student.iTrackID=Track.iTrackID
                    LEFT OUTER JOIN School ON AttendanceToday.iSchoolID=School.iSchoolID
                WHERE 1=1"

if ($PARAM_SchoolIDList.Length -gt 0) {
    $SqlQuery += " AND AttendanceToday.iSchoolID IN ($PARAM_SchoolIDList)"
}

if ($PARAM_SchoolDANList.Length -gt 0) {
    $SqlQuery += " AND School.cCode IN ($PARAM_SchoolDANList)"
}

if ($JustPeriodAttendance -eq $true){
    $SqlQuery += " AND Track.lDaily=0"
} 

if ($JustDailyAttendance -eq $true) {
    $SqlQuery += " AND Track.lDaily=1"
}

write-host $SqlQuery

# CSV Delimeter
# Some systems expect this to be a tab "`t" or a pipe "|".
$Delimeter = ','

# Should all columns be quoted, or just those that contains characters to escape?
# Note: This has no effect on systems with PowerShell versions <7.0 (all fields will be quoted otherwise)
$QuoteAllColumns = $true

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

# This is the "proper" way to do SQL parameters, but it doesn't work.
#$SqlCommand.Parameters.AddWithValue("@START_DATE",$AttendanceDateFrom)
#$SqlCommand.Parameters.AddWithValue("@END_DATE",$AttendanceDateTo)

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