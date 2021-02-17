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
                School.iSchoolID as SchoolID,
                CONCAT('STUDENT-',Student.iStudentID) as StudentID,
                Student.cLastName as StudentLastName,
                Student.cFirstName as StudentFirstName,
                REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(RTRIM(LTRIM(Grades.cName)),'0K','K'),'01','1'),'02','2'),'03','3'),'04','4'),'05','5'),'06','6'),'07','7'),'08','8'),'09','9') as Grade,
                Homeroom.cName as HomeRoom,
                Contact1_RELATIONLV.cName as Contact1Relation,
                Contact1_CONTACT.cLastName as Contact1LastName,
                Contact1_CONTACT.cFirstName as Contact1FirstName,
                Contact1_LOCATION.cPhone as Contact1HomePhone,
                Contact1_CONTACT.mCellphone as Contact1CellPhone,
                Contact1_CONTACT.cBusPhone as Contact1WorkPhone,
                Contact1_CONTACT.mEmail as Contact1Email,
                Contact2_RELATIONLV.cName as Contact2Relation,
                Contact2_CONTACT.cLastName as Contact2LastName,
                Contact2_CONTACT.cFirstName as Contact2FirstName,
                Contact2_LOCATION.cPhone as Contact2HomePhone,
                Contact2_CONTACT.mCellphone as Contact2CellPhone,
                Contact2_CONTACT.cBusPhone as Contact2WorkPhone,
                Contact2_CONTACT.mEmail as Contact2Email,
                Contact3_RELATIONLV.cName as Contact3Relation,
                Contact3_CONTACT.cLastName as Contact3LastName,
                Contact3_CONTACT.cFirstName as Contact3FirstName,
                Contact3_LOCATION.cPhone as Contact3HomePhone,
                Contact3_CONTACT.mCellphone as Contact3CellPhone,
                Contact3_CONTACT.cBusPhone as Contact3WorkPhone,
                Contact3_CONTACT.mEmail as Contact3Email
            FROM
                StudentStatus
                LEFT OUTER JOIN Student ON StudentStatus.iStudentID=Student.iStudentID
                LEFT OUTER JOIN School ON Student.iSchoolID=School.iSchoolID
                LEFT OUTER JOIN Grades ON Student.iGradesID=Grades.iGradesID
                LEFT OUTER JOIN Homeroom ON Student.iHomeroomID=Homeroom.iHomeroomID
                LEFT OUTER JOIN (SELECT CR_o.iStudentID, CR_1.iContactRelationID FROM (SELECT DISTINCT iStudentID FROM ContactRelation) CR_o CROSS APPLY (SELECT iContactID,iContactRelationID FROM ContactRelation CR_i WHERE CR_i.iStudentID=CR_o.iStudentID AND (lMail=1 OR Notify=1) ORDER BY iContactPriority OFFSET 0 ROWS FETCH NEXT 1 ROWS ONLY) CR_1) as ContactRID1 ON ContactRID1.iStudentID=Student.iStudentID
                LEFT OUTER JOIN (SELECT CR_o.iStudentID, CR_1.iContactRelationID FROM (SELECT DISTINCT iStudentID FROM ContactRelation) CR_o CROSS APPLY (SELECT iContactID,iContactRelationID FROM ContactRelation CR_i WHERE CR_i.iStudentID=CR_o.iStudentID AND (lMail=1 OR Notify=1) ORDER BY iContactPriority OFFSET 1 ROWS FETCH NEXT 1 ROWS ONLY) CR_1) as ContactRID2 ON ContactRID2.iStudentID=Student.iStudentID
                LEFT OUTER JOIN (SELECT CR_o.iStudentID, CR_1.iContactRelationID FROM (SELECT DISTINCT iStudentID FROM ContactRelation) CR_o CROSS APPLY (SELECT iContactID,iContactRelationID FROM ContactRelation CR_i WHERE CR_i.iStudentID=CR_o.iStudentID AND (lMail=1 OR Notify=1) ORDER BY iContactPriority OFFSET 2 ROWS FETCH NEXT 1 ROWS ONLY) CR_1) as ContactRID3 ON ContactRID3.iStudentID=Student.iStudentID
                LEFT OUTER JOIN ContactRelation as Contact1_RELATION ON ContactRID1.iContactRelationID=Contact1_RELATION.iContactRelationID
                LEFT OUTER JOIN Contact as Contact1_CONTACT ON Contact1_RELATION.iContactID=Contact1_CONTACT.iContactID
                LEFT OUTER JOIN LookupValues AS Contact1_RELATIONLV ON Contact1_RELATION.iLV_RelationID=Contact1_RELATIONLV.iLookupValuesID
                LEFT OUTER JOIN Location AS Contact1_LOCATION ON Contact1_CONTACT.iLocationID=Contact1_LOCATION.iLocationID
                LEFT OUTER JOIN ContactRelation as Contact2_RELATION ON ContactRID2.iContactRelationID=Contact2_RELATION.iContactRelationID
                LEFT OUTER JOIN Contact as Contact2_CONTACT ON Contact2_RELATION.iContactID=Contact2_CONTACT.iContactID
                LEFT OUTER JOIN LookupValues AS Contact2_RELATIONLV ON Contact2_RELATION.iLV_RelationID=Contact2_RELATIONLV.iLookupValuesID
                LEFT OUTER JOIN Location AS Contact2_LOCATION ON Contact2_CONTACT.iLocationID=Contact2_LOCATION.iLocationID
                LEFT OUTER JOIN ContactRelation as Contact3_RELATION ON ContactRID3.iContactRelationID=Contact3_RELATION.iContactRelationID
                LEFT OUTER JOIN Contact as Contact3_CONTACT ON Contact3_RELATION.iContactID=Contact3_CONTACT.iContactID
                LEFT OUTER JOIN LookupValues AS Contact3_RELATIONLV ON Contact3_RELATION.iLV_RelationID=Contact3_RELATIONLV.iLookupValuesID
                LEFT OUTER JOIN Location AS Contact3_LOCATION ON Contact3_CONTACT.iLocationID=Contact3_LOCATION.iLocationID
            WHERE
                (StudentStatus.dInDate <= getDate()) 
                AND ((StudentStatus.dOutDate <= '1901-01-01') OR (StudentStatus.dOutDate >= getDate()))
                AND (Contact1_CONTACT.iContactID IS NOT NULL)"

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