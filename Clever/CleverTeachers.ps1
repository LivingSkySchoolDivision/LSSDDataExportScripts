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
                UserStaff.UF_2085 AS Teacher_id,
                UserStaff.UF_2085 AS Teacher_number,
                '' AS State_teacher_id,
                Staff.mEmail AS Teacher_email,
                LTRIM(RTRIM(Staff.cFirstName)) AS First_name,
                '' AS Middle_name,
                LTRIM(RTRIM(Staff.cLastName)) AS Last_name,
                '' AS Title,
                Staff.mEmail AS Username,
                '' AS Password
            FROM 
                Staff
                LEFT OUTER JOIN UserStaff ON Staff.iStaffID=UserStaff.iStaffID
                LEFT OUTER JOIN School ON Staff.iSchoolID=School.iSchoolID
            WHERE 
                Staff.lInactive = 0
                AND UserStaff.iBaseStaffIDid = ''
                AND Staff.lClassList=1
                AND School.lInactive=0
                AND 
                    ( 
                        ((SELECT COUNT(iClassID) FROM ClassResource WHERE ClassResource.iStaffID=Staff.iStaffID) > 0)
                        OR ((SELECT COUNT(iHomeroomID) FROM Homeroom WHERE (i1_StaffID=Staff.iStaffID) OR (i2_StaffID=Staff.iStaffID)) > 0)
                        OR ((SELECT COUNT(iClassID) FROM Class WHERE Class.iDefault_StaffID=Staff.iStaffID) > 0)
                    )
                AND LEN(LTRIM(RTRIM(UserStaff.UF_2085))) > 0
                AND LEN(LTRIM(RTRIM(Staff.mEmail))) > 0
            ORDER BY
                Staff.cLastName
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