param (
    [Parameter(Mandatory=$true)][string]$OutputFileName,
    [string]$ConfigFilePath
 )

##############################################
# Script configuration                       #
##############################################

# Edsby allows us to sync roughly a dozen schools to the Sandbox
# The ActiveSchools.csv can be used to adjust the schools we want to send
$ActiveSchools = Import-Csv -Path .\ActiveSchools.csv
$ActiveSchools = $ActiveSchools | Where-Object {$_.Sync -eq 'T'} 
$iSchoolIDs = Foreach ($ID in $ActiveSchools) {
    if ($ID -ne $ActiveSchools[-1]) { $ID.iSchoolID + ',' } else { $ID.iSchoolID }    
}

# SQL Query to run
# The output CSV file will use column names from your SQL query.
# Rename them using "as" - example: "SELECT cFirstName as FirstName FROM Students"
$SqlQuery = "SELECT
                S.iSchoolID AS SchoolID,
                LTRIM(RTRIM(S.cName)) AS SchoolName,
                REPLACE(REPLACE(LTRIM(RTRIM(ST.cName)),'PS','E'),'HS','S') AS SchoolType,
                'Regular' AS SchoolFocus,
                '' AS DistrictID,
                LTRIM(RTRIM(URL.mInfo)) AS SchoolURL,
                CASE WHEN L.cHouseNo !='' THEN LTRIM(RTRIM(CONCAT(L.cHouseNo, ' ', L.cStreet)))
                ELSE LTRIM(RTRIM(L.cStreet)) END AS StreetAddress,
                LTRIM(RTRIM(CITY.cName)) AS City,
                'SK' AS StateProvince,
                LTRIM(RTRIM(C.cName)) AS Country,
                SUBSTRING(L.cPostalCode,1,3)+' '+SUBSTRING(L.cPostalCode,4,3) AS PostalCode,
                LTRIM(RTRIM(L.GridLocation)) AS GridLocation,
                LTRIM(RTRIM('1-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((SC.mInfo), '(',''),')',''),'-',''),' ',''),1,3)+'-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((SC.mInfo), '(',''),')',''),'-',''),' ',''),4,3)+'-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((SC.mInfo), '(',''),')',''),'-',''),' ',''),7,4))) AS Telephone,
                LTRIM(RTRIM('1-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((FAX.mInfo), '(',''),')',''),'-',''),' ',''),1,3)+'-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((FAX.mInfo), '(',''),')',''),'-',''),' ',''),4,3)+'-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((FAX.mInfo), '(',''),')',''),'-',''),' ',''),7,4))) AS FaxNumber,
                LTRIM(RTRIM(EMAIL.mInfo)) AS Email,
                S.EdsbyGradeLevels AS GradeLevel,
                'America/Regina' AS TimeZone,
                CONCAT('STAFF-',STA.iStaffID) AS PrincipalGUID
            FROM School S
                LEFT OUTER JOIN SchoolCommunication SC ON S.iSchoolID = SC.iSchoolID AND SC.iLV_CommunicationTypeID = 3576
                LEFT OUTER JOIN SchoolCommunication FAX ON S.iSchoolID = FAX.iSchoolID AND FAX.iLV_CommunicationTypeID = 4700	
                LEFT OUTER JOIN SchoolCommunication EMAIL ON S.iSchoolID = EMAIL.iSchoolID AND EMAIL.iLV_CommunicationTypeID = 3876
                LEFT OUTER JOIN SchoolCommunication URL ON S.iSchoolID = URL.iSchoolID AND URL.iLV_CommunicationTypeID = 4699
                LEFT OUTER JOIN Location L ON S.iLocationID = L.iLocationID
                LEFT OUTER JOIN LookupValues CITY ON L.iLV_CityID = CITY.iLookupValuesID
                LEFT OUTER JOIN LookupValues ST ON S.iLV_TypeID = ST.iLookupValuesID
                LEFT OUTER JOIN Country C ON S.iCountryID = C.iCountryID
                LEFT OUTER JOIN LookupValues PROV ON S.iLV_RegionID = PROV.iLookupValuesID
                LEFT OUTER JOIN Staff STA ON LEFT(cPrincipal,CHARINDEX(' ', S.cPrincipal)) = STA.cFirstName AND SUBSTRING(cPrincipal, CHARINDEX(' ', cPrincipal)+1, LEN(cPrincipal)-(CHARINDEX(' ', cPrincipal)-1)) = STA.cLastName AND STA.cUserName NOT LIKE '%.%.%'
            WHERE                 
				S.iSchoolID IN ($iSchoolIDs)
            ORDER BY S.cName;"

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