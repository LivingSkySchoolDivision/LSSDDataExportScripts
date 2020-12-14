param (
    [Parameter(Mandatory=$true)][string]$OutputFileName,
    [string]$ConfigFilePath
 )

# This is based on an import file format provided by BusPlanner/Georef in 2017.
# The format corresponds to an XML "template" file found in the busplanner directory - this file has been included in this repo.
# This template file was originally found in "C:\inetpub\wwwroot\BusPlannerWebSuite\Applications\BusPlannerTasks\Files" on the BusPlanner server.

##############################################
# Script configuration                       #
##############################################

# SQL Query to run
# The output CSV file will use column names from your SQL query.
# Rename them using "as" - example: "SELECT cFirstName as FirstName FROM Students"
$SqlQuery = "SELECT
                School.cCode as SchoolDAN,
                Student.cStudentNumber as StudentNumber,
                Student.cFirstName as FirstName,
                Student.cLastName as LastName,
                CONVERT (varchar,Student.dBirthdate,3) as BirthDate,
                LV_GENDER.cCode as Gender,
                REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(Grades.cName)),'0k','k'),'09','9'),'08','8'),'07','7'),'06','6'),'05','5'),'04','4'),'03','3'),'02','2'),'01','1') as Grade,
                Location.cPhone as Phone,
                Location.cHouseNo as House,
                Location.cStreet as Street,
                LOC_CITY.cName as City,
                LOC_PROV.cName as Province,
                Country.cName as Country,
                Location.cPostalCode as PostalCode,
                UserStudent.cReserveName as ReserveName,
                UserStudent.cReserveHouse as ReserveHouse,
                CAST(
                    CASE
                        WHEN UserStudent.UF_1651 <> 0
                            THEN CONCAT(UserStudent.UF_1651,'-',UserStudent.UF_2098,'-',UserStudent.UF_1653_1,'-',UserStudent.UF_1654_1,'-W',UserStudent.UF_2093)
                    END AS varchar
                ) as LandLocation,
                UserStudent.UF_1651 as Quarter,
                UserStudent.UF_2098 as Section,
                UserStudent.UF_1653_1 as Township,
                UserStudent.UF_1654_1 as Range,
                UserStudent.UF_2093 as Meridian,
                UserStudent.UF_2096 as RiverLot,
                Contact1.iContactID as Contact1ID,
                Contact1.cFirstName as Contact1FirstName,
                Contact1.cLastName as Contact1LastName,
                ContactRelation1LV.cName as Contact1Relation,
                Contact1Location.cPhone as Contact1HomePhone,
                Contact1.cBusPhone as Contact1WorkPhone,
                Contact1.mCellphone as Contact1CellPhone,
                Contact1.mEmail as Contact1Email,
                Contact2.iContactID as Contact2ID,
                Contact2.cFirstName as Contact2FirstName,
                Contact2.cLastName as Contact2LastName,
                ContactRelation2LV.cName as Contact2Relation,
                Contact2Location.cPhone as Contact2HomePhone,
                Contact2.cBusPhone as Contact2WorkPhone,
                Contact2.mCellphone as Contact2CellPhone,
                Contact2.mEmail as Contact2Email
                FROM 
                    StudentStatus
                    LEFT OUTER JOIN Student ON StudentStatus.iStudentID=Student.iStudentID
                    LEFT OUTER JOIN School ON Student.iSchoolID=School.iSchoolID
                    LEFT OUTER JOIN Grades ON Student.iGradesID=Grades.iGradesID
                    LEFT OUTER JOIN Location ON Student.iLocationID=Location.iLocationID
                    LEFT OUTER JOIN LookupValues AS LOC_CITY ON Location.iLV_CityID=LOC_CITY.iLookupValuesID
                    LEFT OUTER JOIN LookupValues AS LOC_PROV ON Location.iLV_RegionID=LOC_PROV.iLookupValuesID
                    LEFT OUTER JOIN Country ON Location.iCountryID=Country.iCountryID
                    LEFT OUTER JOIN LookupValues AS LV_GENDER ON Student.iLV_GenderID=LV_GENDER.iLookupValuesID
                    LEFT OUTER JOIN UserStudent ON Student.iStudentID=UserStudent.iStudentID
                    LEFT OUTER JOIN ContactRelation as ContactRelation1 ON (SELECT TOP 1 iContactRelationID FROM ContactRelation WHERE iStudentID=Student.iStudentID AND lmail=1 ORDER BY iContactPriority)=ContactRelation1.iContactRelationID
                    LEFT OUTER JOIN Contact AS Contact1 ON ContactRelation1.iContactID=Contact1.iContactID
                    LEFT OUTER JOIN LookUpValues as ContactRelation1LV ON ContactRelation1.iLV_RelationID=ContactRelation1LV.iLookupValuesID
                    LEFT OUTER JOIN Location as Contact1Location ON Contact1.iLocationID=Contact1Location.iLocationID
                    LEFT OUTER JOIN ContactRelation as ContactRelation2 ON (SELECT TOP 1 iContactRelationID FROM ContactRelation WHERE iStudentID=Student.iStudentID AND lmail=1 AND iContactID<>Contact1.iContactID ORDER BY iContactPriority)=ContactRelation2.iContactRelationID
                    LEFT OUTER JOIN Contact AS Contact2 ON ContactRelation2.iContactID=Contact2.iContactID
                    LEFT OUTER JOIN LookUpValues as ContactRelation2LV ON ContactRelation2.iLV_RelationID=ContactRelation2LV.iLookupValuesID
                    LEFT OUTER JOIN Location as Contact2Location ON Contact2.iLocationID=Contact2Location.iLocationID
                WHERE
                    (StudentStatus.dInDate <=  { fn CURDATE() }) AND
                    ((StudentStatus.dOutDate < '1901-01-01') OR (StudentStatus.dOutDate >=  { fn CURDATE() }))  AND 
                    (StudentStatus.lOutsideStatus = 0)
                ;"

# CSV Delimeter
# Some systems expect this to be a tab "`t" or a pipe "|".
$Delimeter = "`t"

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