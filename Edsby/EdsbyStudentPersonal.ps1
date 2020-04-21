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
                S.iStudentID AS StudentGUID, 
                cLegalFirstName AS FirstName, 
                cLegalLAStName AS LaStName, 
                S.iSchoolID AS SchoolID, 
                cStudentNumber AS SID, 
                cGovernmentNumber AS MinistryID, 
                LEFT(LV.cName,1) AS Gender, 
                HR.i1_StaffID AS HomeRoomStaffGUID, 
                '' AS Prefix, 
                cMiddlename AS MiddleName, 
                cLegalSuffix AS Suffix, 
                cFirstName AS PreferredName, 
                replace(replace(concat(l.cHouseNo, ' ', l.cStreet), char(10), ''), char(13), '') AS StreetAddress, 
                city.cName AS City, 
                prov.cName AS StateProvince, 
                country.cName AS Country, 
                l.cPostalCode AS PostalCode, 
                s.mEmail AS Email, 
                CASE WHEN S.mCellPhone >'' THEN '1-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((s.mCellPhone), '(',''),')',''),'-',''),' ',''),1,3)+'-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((s.mCellPhone), '(',''),')',''),'-',''),' ',''),4,3)+'-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((s.mCellPhone), '(',''),')',''),'-',''),' ',''),7,4) ELSE '' END AS Telephone, 
                FORMAT(s.dBirthdate, 'yyyy-MM-dd') AS Birthday, 
                RTRIM(LTRIM(cUserName)) AS UserID, 
                concat(left(cFirstName, 1), left(cLastName, 1), cStudentNumber) AS Password, 
                'T' AS IntegrationAuth,  
                RTRIM(LTRIM(g.cName)) AS Grade
            FROM Student S
                LEFT OUTER JOIN LookupValues LV ON S.iLV_GenderID = LV.iLookupValuesID
                LEFT OUTER JOIN Homeroom hr ON S.iHomeroomID = hr.iHomeroomID
                LEFT OUTER JOIN Location l ON s.iLocationID = l.iLocationID
                LEFT OUTER JOIN LookupValues city ON l.iLV_CityID = city.iLookupValuesID
                LEFT OUTER JOIN LookupValues prov ON l.iLV_RegionID = prov.iLookupValuesID
                LEFT OUTER JOIN Country country ON l.iCountryID = country.iCountryID
                LEFT OUTER JOIN Grades g ON s.iGradesID = g.iGradesID
                LEFT OUTER JOIN StudentStatus SS ON S.iStudentID = SS.iStudentID
            WHERE 
                (SS.dInDate <=  { fn CURDATE() }) AND
                ((SS.dOutDate < '1901-01-01') OR (SS.dOutDate >=  { fn CURDATE() }))  AND 
                (SS.lOutsideStatus = 0)
            ORDER BY S.iStudentID;"

# CSV Delimeter
# Some systems expect this to be a tab "\t" or a pipe "|".
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