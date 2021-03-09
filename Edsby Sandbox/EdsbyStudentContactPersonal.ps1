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
$ActiveSchools = $ActiveSchools | ? {$_.Sync -eq 'T'} 
$iSchoolIDs = Foreach ($ID in $ActiveSchools) {
    if ($ID -ne $ActiveSchools[-1]) { $ID.iSchoolID + ',' } else { $ID.iSchoolID }    
}

# SQL Query to run
# The output CSV file will use column names from your SQL query.
# Rename them using "as" - example: "SELECT cFirstName as FirstName FROM Students"
$SqlQuery = "SELECT 
                CONCAT('CONTACT-', C.iContactID) AS ContactGUID,
                C.cFirstName AS FirstName,
                C.cLastName AS LastName,
                '' AS Relation,
                '' AS AccessToRecords,
                C.iSchoolID AS SchoolID,
                C.mEmail AS Email,
                '' AS ContactSequence,
                PR.cName AS Prefix,
                '' AS MiddleName,
                '' AS Suffix,
                '' AS PreferredName,
                REPLACE(REPLACE(CONCAT(L.cHouseNo, ' ', l.cStreet), char(10), ''), char(13), '') AS StreetAddress,
                CITY.cName AS City,
                PROV.cName AS StateProvince,
                Country.cName AS Country,
                L.cPostalCode AS PostalCode,
                CASE WHEN L.cPhone >'' THEN '1-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((L.cPhone), '(',''),')',''),'-',''),' ',''),1,3)+'-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((L.cPhone), '(',''),')',''),'-',''),' ',''),4,3)+'-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((L.cPhone), '(',''),')',''),'-',''),' ',''),7,4) ELSE '' END AS Telephone,
                CASE WHEN C.mCellPhone >'' THEN '1-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((C.mCellPhone), '(',''),')',''),'-',''),' ',''),1,3)+'-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((C.mCellPhone), '(',''),')',''),'-',''),' ',''),4,3)+'-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((C.mCellPhone), '(',''),')',''),'-',''),' ',''),7,4) ELSE '' END AS MobilePhone, 
                CASE WHEN C.cBusPhone >'' THEN '1-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((C.cBusPhone), '(',''),')',''),'-',''),' ',''),1,3)+'-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((C.cBusPhone), '(',''),')',''),'-',''),' ',''),4,3)+'-'+SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE((C.cBusPhone), '(',''),')',''),'-',''),' ',''),7,4) ELSE '' END AS WorkPhone, 
                '' AS FaxNumber,
                'F' AS IntegrationAuth
            FROM 
                Contact C
                LEFT OUTER JOIN LookupValues PR ON C.iLV_TitleID = PR.iLookupValuesID 
                LEFT OUTER JOIN Location L ON C.iLocationID = L.iLocationID 
                LEFT OUTER JOIN LookupValues CITY ON L.iLV_CityID = CITY.iLookupValuesID
                LEFT OUTER JOIN LookupValues PROV ON L.iLV_RegionID = PROV.iLookupValuesID
                LEFT OUTER JOIN Country ON L.iCountryID = Country.iCountryID
            WHERE 
                C.iSchoolID IN ($iSchoolIDs) AND
                C.iContactID IN (
                    SELECT 
                        DISTINCT(ContactRelation.iContactID) 
                    FROM 
                        StudentStatus 
                        LEFT OUTER JOIN ContactRelation ON StudentStatus.iStudentID=ContactRelation.iStudentID
                        INNER JOIN LookupValues LV ON ContactRelation.iLV_RelationID = LV.iLookupValuesID
                    WHERE 
                        LV.cName NOT IN (
                            'Doctor', 
                            'Mother & Father', 
                            'Sister', 
                            'Self', 
                            'Sibling', 
                            'Wife', 
                            'Cousin', 
                            'Nurse', 
                            'Other', 
                            'Brother', 
                            'Staff Contact', 
                            'Mr', 
                            'Principal', 
                            'Friend', 
                            'Niece', 
                            'Unknown', 
                            'Nurse Practitioner'
                            ) AND
                        C.iSchoolID NOT IN (
                            5850953, -- Major School
                            5850963, -- Manacowin School
                            5850964, -- Phoenix School
                            5851066, -- Zinactive
                            5851067, -- Home Based 
                            5850943 -- Cut Knife Elementary
                            )
                        AND (StudentStatus.dInDate <=  getDate() + 1) 
                        AND (
                            (StudentStatus.dOutDate < '1901-01-01') OR (StudentStatus.dOutDate >=  { fn CURDATE() }))  
                            AND (StudentStatus.lOutsideStatus = 0)
                        ) 
                ORDER BY 
                    C.iContactID;"

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