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
                BaseSchool.cCode AS School_id,
                Student.cStudentNumber AS Student_id,
                Student.cStudentNumber AS Student_number,
                Student.cGovernmentNumber AS State_id,
                Student.cLastName AS Last_name,
                Student.cMiddleName AS Middle_name,
                Student.cFirstName AS First_name,
                CASE
                    WHEN Grades.cName = '0k' THEN 'Kindergarten'
                    WHEN Grades.cName = 'pk' THEN 'PreKindergarten'
                    ELSE CAST(CAST(LTRIM(RTRIM(Grades.cName)) AS INT) AS VARCHAR)
                END AS Grade,
                Gender.cCode AS Gender,
                '' AS Graduation_year,
                FORMAT(Student.dBirthdate, 'MM/dd/yyyy') AS DOB,
                '' AS Race,
                '' AS Hispanic_Latino,
                '' AS Home_language,
                '' AS Ell_status,
                '' AS Frl_status,
                '' AS IEP_status,
                '' AS Student_street,
                '' AS Student_city,
                '' AS Student_state,
                '' AS Student_zip,
                Student.mEmail AS Student_email,
                ContactRelationship.cName AS Contact_relationship,
                CASE
                    WHEN (ContactRelationship.cName IS NOT NULL) THEN 'family'
                    ELSE ''
                END AS Contact_type,
                CONCAT(Contact.cFirstName, ' ', Contact.cLastName) AS Contact_name,
                LTRIM(RTRIM(ContactLocation.cPhone)) AS Contact_phone,
                Contact.mEmail AS Contact_email,
                Contact.iContactID AS Contact_sis_id,
                Student.mEmail AS Username,
                '' AS Password,
                '' AS Unweighted_gpa,
                '' AS Weighted_gpa
            FROM
                Student                
                LEFT OUTER JOIN School as BaseSchool ON Student.iSchoolID=BaseSchool.iSchoolID
                LEFT OUTER JOIN LookupValues as Gender ON Student.iLV_GenderID=Gender.iLookupValuesID
                LEFT OUTER JOIN Grades ON Student.iGradesID=Grades.iGradesID
                LEFT OUTER JOIN (SELECT * FROM ContactRelation WHERE lMail=1 OR Notify=1) AS ImportantContacts ON Student.iStudentID=ImportantContacts.iStudentID
                LEFT OUTER JOIN LookupValues AS ContactRelationship ON ImportantContacts.iLV_RelationID=ContactRelationship.iLookupValuesID
                LEFT OUTER JOIN Contact ON ImportantContacts.iContactID=Contact.iContactID
                LEFT OUTER JOIN Location AS ContactLocation ON Contact.iLocationID=ContactLocation.iLocationID
            WHERE 
                Student.iStudentID in (
                    SELECT 
                        DISTINCT(iStudentID)
                    FROM 
                        StudentStatus
                    Where
                        StudentStatus.dInDate <= GETDATE()
                        AND (
                            StudentStatus.dOutDate < '1901-01-01' OR
                            StudentStatus.dOutDate >= GETDATE()
                        )
                    )
            ORDER BY 
                Student.cStudentNumber
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