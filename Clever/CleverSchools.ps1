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
                School.cCode AS School_id,
                LTRIM(RTRIM(School.cName)) AS School_name,
                School.cCode AS School_number,
                '' AS State_id,
                CASE
                    WHEN LGrades.cName = '0k' THEN 'Kindergarten'
                    WHEN LGrades.cName = 'pk' THEN 'PreKindergarten'
                    ELSE CAST(CAST(LTRIM(RTRIM(LGrades.cName)) AS INT) AS VARCHAR)
                END AS Low_grade,
                CASE
                    WHEN HGrades.cName = '0k' THEN 'Kindergarten'
                    WHEN HGrades.cName = 'pk' THEN 'PreKindergarten'
                    ELSE CAST(CAST(LTRIM(RTRIM(HGrades.cName)) AS INT) AS VARCHAR)
                END AS High_grade,
                '' AS Principal,
                LTRIM(RTRIM(EMAIL.mInfo)) AS Principal_email,
                CASE 
                    WHEN Location.cHouseNo !='' 
                    THEN 
                        LTRIM(RTRIM(CONCAT(Location.cHouseNo, ' ', Location.cStreet)))
                    ELSE 
                        LTRIM(RTRIM(Location.cStreet)) 
                END AS School_address,
                LTRIM(RTRIM(CITY.cName)) AS School_city,
                REGION.cCode AS School_state,
                SUBSTRING(Location.cPostalCode,1,3)+' '+SUBSTRING(Location.cPostalCode,4,3) AS School_zip,
                SC.mInfo AS School_phone
            FROM School
                LEFT OUTER JOIN SchoolCommunication SC ON School.iSchoolID = SC.iSchoolID AND SC.iLV_CommunicationTypeID = 3576                
                LEFT OUTER JOIN SchoolCommunication EMAIL ON School.iSchoolID = EMAIL.iSchoolID AND EMAIL.iLV_CommunicationTypeID = 3876
                LEFT OUTER JOIN SchoolCommunication URL ON School.iSchoolID = URL.iSchoolID AND URL.iLV_CommunicationTypeID = 4699
                LEFT OUTER JOIN Location ON School.iLocationID = Location.iLocationID
                LEFT OUTER JOIN LookupValues CITY ON Location.iLV_CityID = CITY.iLookupValuesID
                LEFT OUTER JOIN LookupValues REGION ON Location.iLV_RegionID = REGION.iLookupValuesID
                LEFT OUTER JOIN Country C ON School.iCountryID = C.iCountryID
                LEFT OUTER JOIN LookupValues PROV ON School.iLV_RegionID = PROV.iLookupValuesID 
                LEFT OUTER JOIN Grades AS LGrades ON School.iLow_GRadesID = LGrades.iGradesID
                LEFT OUTER JOIN Grades AS HGrades ON School.iHigh_GradesID = HGrades.iGradesID               
            WHERE                 
                School.lInactive = 0
                AND School.iDistrictID = 1
            ORDER BY School.cName;"

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