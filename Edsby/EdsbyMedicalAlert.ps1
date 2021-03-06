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
                S.iSchoolID AS SchoolID, 
                CONCAT('STUDENT-',S.iStudentID) AS GUID,
                CASE WHEN 
					SP.dExpiryDate > GETDATE() OR SP.dExpiryDate = '1900-01-01' THEN '4' ELSE '3' 
				END AS Severity,
                CASE WHEN
                    SP.dExpiryDate > GETDATE() OR SP.dExpiryDate = '1900-01-01' THEN 'Education' ELSE 'Medical' 
                END AS Alert,
                CONCAT('MED-',S.iStudentID) as RecordID,
                CASE WHEN 
                    SP.dExpiryDate > GETDATE() OR SP.dExpiryDate = '1900-01-01' THEN SP.cRestrictionDetails ELSE S.mMedical 
                END AS MedicalAlertString,
                'Active' as Status
            FROM 
                Student S
                LEFT OUTER JOIN StudentStatus SS ON S.iStudentID = SS.iStudentID
                LEFT OUTER JOIN StudentProtection SP ON S.iStudentID = SP.iStudentID
            WHERE 
                (S.mMedical != '' OR
                SP.cRestrictionDetails != '') AND
                (SS.dInDate <= getDate() + 1) AND
                ((SS.dOutDate < '1901-01-01') OR (SS.dOutDate >=  { fn CURDATE() })) AND
                S.iSchoolID NOT IN (5851067) --HomeSchool;"

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