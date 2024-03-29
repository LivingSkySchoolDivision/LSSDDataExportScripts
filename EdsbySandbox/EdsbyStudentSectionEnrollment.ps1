param (
    [Parameter(Mandatory=$true)][string]$OutputFileName,
    [string]$ConfigFilePath
 )

##############################################
# Script configuration                       #
##############################################

# Edsby allows us to sync roughly a dozen schools to the Sandbox
# The ActiveSchools.csv can be used to adjust the schools we want to send
$ActiveSchools = Import-Csv -Path C:\Sync\EdsbySandbox\ActiveSchools.csv | Where-Object {$_.Sync -eq 'T'} 
$iSchoolIDs = Foreach ($ID in $ActiveSchools) {
    if ($ID -ne $ActiveSchools[-1]) { $ID.iSchoolID + ',' } else { $ID.iSchoolID }    
}

# SQL Query to run
# The output CSV file will use column names from your SQL query.
# Rename them using "as" - example: "SELECT cFirstName as FirstName FROM Students"
$SqlQuery = "SELECT 
                HR.iSchoolID as SchoolID,
                CONCAT(HR.iSchoolID,'-',HR.iHomeroomID) as SectionGUID,
                CONCAT('STUDENT-',S.iStudentID) as StudentGUID,
				CASE WHEN SS.dInDate > t.dStartDate THEN FORMAT(SS.dInDate, 'yyyy-MM-dd') ELSE FORMAT(t.dStartDate, 'yyyy-MM-dd') END AS StartDate
            FROM
                Homeroom HR
                INNER JOIN Student S ON HR.iHomeroomID = S.iHomeroomID
                INNER JOIN StudentStatus SS ON S.iStudentID = SS.iStudentID
                INNER JOIN Track T ON S.iTrackID = T.iTrackID
            WHERE 
                T.lDaily = 1 AND
                HR.iSchoolID IN ($iSchoolIDs) AND
                (SS.dInDate <=  getDate() + 1) AND
                ((SS.dOutDate < '1901-01-01') OR (SS.dOutDate >=  { fn CURDATE() }))

            UNION 
                ALL

            SELECT 
                E.iSchoolID as SchoolID,
                CONCAT(E.iSchoolID,'-',E.iClassID) as SectionGUID,
                CONCAT('STUDENT-',E.iStudentID) as StudentGUID,
				FORMAT(E.dInDate, 'yyyy-MM-dd') AS StartDate
            FROM
                Enrollment E
                INNER JOIN CLASS C ON E.iClassID = C.iClassID
				INNER JOIN StudentStatus SS ON E.iStudentID = SS.iStudentID
            WHERE
                E.iSchoolID IN ($iSchoolIDs) AND
                (iLV_CompletionStatusID=0 OR iLV_CompletionStatusID=3568) AND
                (SS.dInDate <=  getDate() + 1) AND
                ((SS.dOutDate < '1901-01-01') OR (SS.dOutDate >=  { fn CURDATE() })) AND
                C.iLV_SessionID != '4720' --Session set to No Edsby;"


# CSV Delimeter
# Some systems expect this to be a tab "`t" or a pipe "|".
$Delimeter = ","

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