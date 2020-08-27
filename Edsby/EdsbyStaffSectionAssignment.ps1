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
$SqlQuery = "SELECT
                HR.iSchoolID AS SchoolID,
                CONCAT(HR.iSchoolID,'-',HR.iHomeroomID) AS SectionGUID,
                CONCAT('STAFF-',ST1.iStaffID) AS StaffGUID,
                R.cName AS Role
            FROM 
                Homeroom HR
                LEFT OUTER JOIN Staff ST1 ON HR.i1_StaffID = ST1.iStaffID
                LEFT OUTER JOIN Staff ST2 ON HR.i2_StaffID = ST2.iStaffID
                LEFT OUTER JOIN UserStaff US ON ST1.iStaffID = US.iStaffID	OR ST2.iStaffID = US.iStaffID
                LEFT OUTER JOIN LookupValues R ON US.iEdsbyRoleid = R.iLookupValuesID
                INNER JOIN Track T ON HR.iSchoolID = T.iSchoolID
            WHERE 
                HR.iSchoolID NOT IN (
                    5850953, -- Major School
                    5850963, -- Manacowin School
                    5850964, -- Phoenix School
                    5851066, -- Zinactive
                    5851067 -- Home Based 
                ) AND
                T.lDaily = 1
            UNION
                ALL
            SELECT
                C.iSchoolID AS SchoolID,
                CONCAT(C.iSchoolID,'-',C.iClassID) AS SectionGUID,
                CONCAT('STAFF-',CR.iStaffID) AS StaffGUID,
                R.cName AS Role
            FROM 
                Class C
                LEFT OUTER JOIN ClassResource CR ON C.iClassID = CR.iClassID
                LEFT OUTER JOIN UserStaff US ON CR.iStaffID = US.iStaffID
                LEFT OUTER JOIN LookupValues R ON US.iEdsbyRoleid = R.iLookupValuesID
                INNER JOIN Staff S ON CR.iStaffID = S.iStaffID
            WHERE 
                C.iSchoolID NOT IN (
                    5850953, -- Major School
                    5850963, -- Manacowin School
                    5850964, -- Phoenix School
                    5851066, -- Zinactive
                    5851067 -- Home Based 
                ) AND
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