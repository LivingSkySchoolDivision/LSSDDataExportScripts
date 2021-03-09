param (
    [Parameter(Mandatory=$true)][string]$OutputFileName,
    [string]$ConfigFilePath
 )

##############################################
# Script configuration                       #
##############################################

# Edsby allows us to sync roughly a dozen schools to the Sandbox
# The ActiveSchools.csv can be used to adjust the schools we want to send
$ActiveSchools = Import-Csv -Path .\ActiveSchools.csv | Where-Object {$_.Sync -eq 'T'} 
$iSchoolIDs = Foreach ($ID in $ActiveSchools) {
    if ($ID -ne $ActiveSchools[-1]) { $ID.iSchoolID + ',' } else { $ID.iSchoolID }    
}

# SQL Query to run
# The output CSV file will use column names from your SQL query.
# Rename them using "as" - example: "SELECT cFirstName as FirstName FROM Students"
$SqlQuery = "SELECT DISTINCT
                HR.iSchoolID AS SchoolID,
                CONCAT(HR.iSchoolID,'-',HR.iHomeroomID) AS SectionGUID,
                CASE WHEN
                    US.iBaseStaffIDid <> 0 THEN CONCAT('STAFF-',US.iBaseStaffIDid) ELSE CONCAT('STAFF-',ST.iStaffID) 
                END AS StaffGUID,
                R.cName AS Role
            FROM 
                Homeroom HR
                LEFT OUTER JOIN Staff ST ON HR.i1_StaffID = ST.iStaffID OR HR.i2_StaffID = ST.iStaffID
                INNER JOIN UserStaff US ON ST.iStaffID = US.iStaffID
                LEFT OUTER JOIN LookupValues R ON US.iEdsbyRoleid = R.iLookupValuesID
                INNER JOIN Student S ON HR.iHomeroomID = S.iHomeroomID
                INNER JOIN StudentStatus SS ON S.iStudentID = SS.iStudentID
                INNER JOIN Track T ON S.iTrackID = T.iTrackID
            WHERE 
                ST.iSchoolID IN ($iSchoolIDs) AND
                (SS.dInDate <=  { fn CURDATE() }) AND
                ((SS.dOutDate < '1901-01-01') OR (SS.dOutDate >=  { fn CURDATE() })) AND
                T.lDaily = 1

            UNION
                ALL

            SELECT
                C.iSchoolID AS SchoolID,
                CONCAT(C.iSchoolID,'-',C.iClassID) AS SectionGUID,
				CASE WHEN 
					US.iBaseStaffIDid <> 0 THEN CONCAT('STAFF-',US.iBaseStaffIDid) ELSE CONCAT('STAFF-',CR.iStaffID) 
				END AS StaffGUID,
                R.cName AS Role
            FROM 
                Class C
                LEFT OUTER JOIN ClassResource CR ON C.iClassID = CR.iClassID
                LEFT OUTER JOIN UserStaff US ON CR.iStaffID = US.iStaffID
                LEFT OUTER JOIN LookupValues R ON US.iEdsbyRoleid = R.iLookupValuesID
                INNER JOIN Staff S ON CR.iStaffID = S.iStaffID
            WHERE 
                C.iSchoolID IN ($iSchoolIDs) AND
                S.iSchoolID = C.iSchoolID;"

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