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
                CONCAT('CONTACT-', ContactRelation.iContactID) as ContactGUID,
                CONCAT('STUDENT-',ContactRelation.iStudentID) as StudentGUID,
                S.iSchoolID as SchoolID,
                LVRelation.cName as Relation,
                CASE WHEN(ContactRelation.lMail=1) THEN 'Yes' ELSE 'No' END AS AccessToRecords,
                'Unknown' as LegalGuardian,
                'Unknown' as PickupRights,
                CASE WHEN(ContactRelation.lLivesWithStudent=1) THEN 'Yes' ELSE 'No' END AS LivesWith,
                'Unknown' as EmergencyContact,
                'Unknown' as HasCustody,
                'Unknown' as DisciplinaryContact,
                'Unknown' as PrimaryCareProvider,
                ContactRelation.iContactPriority as ContactSequence
            FROM
                ContactRelation
                LEFT OUTER JOIN Contact ON ContactRelation.iContactID=Contact.iContactID
                LEFT OUTER JOIN StudentStatus ON ContactRelation.iStudentID=StudentStatus.iStudentID
                LEFT OUTER JOIN LookupValues AS LVRelation ON ContactRelation.iLV_RelationID=LVRelation.iLookupValuesID
                LEFT OUTER JOIN Student S ON ContactRelation.iStudentID = S.iStudentID
            WHERE 
                (StudentStatus.dInDate <=  getDate() + 1) AND
                ((StudentStatus.dOutDate < '1901-01-01') OR (StudentStatus.dOutDate >=  { fn CURDATE() }))  AND 
                (StudentStatus.lOutsideStatus = 0) AND
                S.iSchoolID NOT IN (5851067) --HomeSchool 
                AND LVRelation.cName NOT IN (
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
                );"

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