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
                School.iSchoolID as SchoolID,
                CONCAT('STUDENT-',Student.iStudentID) as StudentID,
                Student.cLastName as StudentLastName,
                Student.cFirstName as StudentFirstName,
                REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(RTRIM(LTRIM(Grades.cName)),'0K','K'),'01','1'),'02','2'),'03','3'),'04','4'),'05','5'),'06','6'),'07','7'),'08','8'),'09','9') as Grade,
                Homeroom.cName as HomeRoom,
                Contact.cLastName as ContactLastName,
                Contact.cFirstName as ContactFirstName,
                ContactRelationLV.cName as ContactRelation,
                ContactRelation.iContactPriority as ContactPriority,
                ContactLocation.cPhone as ContactHomePhone,
                Contact.mCellPhone as ContactMobilePhone,
                Contact.cBusPhone as ContactWorkPhone,
                Contact.mEmail as ContactEmail
            FROM
                StudentStatus
                LEFT OUTER JOIN Student ON StudentStatus.iStudentID=Student.iStudentID
                LEFT OUTER JOIN School ON Student.iSchoolID=School.iSchoolID
                LEFT OUTER JOIN Grades ON Student.iGradesID=Grades.iGradesID
                LEFT OUTER JOIN Homeroom ON Student.iHomeroomID=Homeroom.iHomeroomID
                LEFT OUTER JOIN ContactRelation ON Student.iStudentID=ContactRelation.iStudentID
                LEFT OUTER JOIN Contact ON ContactRelation.iContactID=Contact.iContactID
                LEFT OUTER JOIN LookupValues AS ContactRelationLV ON ContactRelation.iLV_RelationID=ContactRelationLV.iLookupValuesID
                LEFT OUTER JOIN Location AS ContactLocation ON Contact.iLocationID=ContactLocation.iLocationID
            WHERE
                (StudentStatus.dInDate <= getDate()) AND
                ((StudentStatus.dOutDate < '1901-01-01') OR (StudentStatus.dOutDate >= getDate())) AND	
                (ContactRelation.lMail=1 OR ContactRelation.Notify=1)"

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