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
                CASE WHEN
                    us.iBaseStaffIDid <> 0 THEN CONCAT('STAFF-', us.iBaseStaffIDid) ELSE CONCAT('STAFF-',s.iStaffID) 
                END AS StaffGUID, 
                cFirstName as FirstName, 
                cLastName as LastName, 
                s.iSchoolID AS SchoolID, 
                '' as Prefix, 
                '' as Title, 
                '' as StaffType, 
                title.cName AS Role, 
                '' as MiddleName, 
                '' as Suffix, 
                '' as PreferredName, 
                '' as StreetAddress, 
                '' as City, 
                '' as StateProvince, 
                '' as Country, 
                '' as PostalCode, 
                s.mEmail as Email, 
                '' as Telephone, 
                '' as Gender, 
                FORMAT(us.dBirthDate, 'yyyy-MM-dd') as Birthday, 
                CASE WHEN
                    us.iBaseStaffIDid <> 0 THEN LEFT(s.mEmail, CHARINDEX('@',s.mEmail)-1) ELSE s.cUserName 
                END AS UserID, 
                '' as Password, 
                'T' as IntegrationAuth 
            FROM Staff s
                INNER JOIN UserStaff us ON s.iStaffID = us.iStaffID
                LEFT OUTER JOIN LookupValues title ON us.iEdsbyRoleid = title.iLookupValuesID
            WHERE 
                s.iSchoolID IN ($iSchoolIDs) 
                AND S.lInactive = 0
                AND cLastName NOT LIKE 'ADMIN%' 
                AND cLastName NOT LIKE '%egov%' 
                AND s.iStaffID NOT IN 
                    (
                    43,--Karleen Graff
                    268, --Carrie Day
                    1457, --Colleen Fouhy
                    1482, --SRB Remote
                    1528, --Shannon Lessard
                    1532, --Douglas Drover
                    1938  --Terry Korpach
                    )
            ORDER BY s.iStaffID;"

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