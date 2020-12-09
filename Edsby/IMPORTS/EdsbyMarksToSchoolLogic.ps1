param (
    [Parameter(Mandatory=$true)][string]$InputFileName,
    [string]$ConfigFilePath,
    [bool]$DryRun
 )

###########################################################################
# CHANGES YOU NEED TO MAKE TO YOUR DATABASE                               #
###########################################################################

# Add the following fields to your MarksHistory table
#  ImportBatchID - varchar(40)
#  ImportTimestamp - datetime


###########################################################################
# Editable values                                                         #
###########################################################################

# Marks will be entered with this completion status ID from the Lookup values table. This should correspond to your "Completed" or "COMP" value in your lookup values table
$CompletionStatusID = 3569


###########################################################################
# Functions                                                               #
###########################################################################
function Get-CSV {
    param(
        [Parameter(Mandatory=$true)][String] $CSVFile
    )

    return import-csv $CSVFile -header('YearID','ReportingTermNumber','SchoolID','StudentGUID','StudentID','CourseCode','SubSection','TermGrade','FinalGrade','Comment','SectionGUID') | Select -skip 1
}

Function Get-Hash
{
    param
    (
        [String] $String
    )
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($String)
    $hashfunction = [System.Security.Cryptography.HashAlgorithm]::Create('SHA1')
    $StringBuilder = New-Object System.Text.StringBuilder
    $hashfunction.ComputeHash($bytes) |
    ForEach-Object {
        $null = $StringBuilder.Append($_.ToString("x2"))
    }

    return $StringBuilder.ToString()
}

function Convert-SectionID {
    param(
        [Parameter(Mandatory=$true)] $InputString,
        [Parameter(Mandatory=$true)] $SchoolID
    )
    return $InputString.Replace("$SchoolID-","")
}

function Convert-Year {
    param(
        [Parameter(Mandatory=$true)] $InputString
    )

    return [int]$([int]$InputString-1)
}

function Convert-StudentID {
    param(
        [Parameter(Mandatory=$true)] $InputString
    )

    return [int]$InputString.Replace("STUDENT-","")
}


function Get-ByID {
    param(
        [Parameter(Mandatory=$true)] $ID,
        [Parameter(Mandatory=$true)] $Haystack
    )

    foreach($Obj in $Haystack) {
        if ($Obj.ID -eq $ID) {
            return $Obj
        }
    }

    return $null
}

function Get-CreditEarned {
    param(
        [Parameter(Mandatory=$true)] [int]$PossibleCredits,
        [Parameter(Mandatory=$true)] [int]$FinalMark
    )

    if ($FinalMark -gt 49) {
        return $PossibleCredits
    } else {
        return 0
    }
}

function Get-SQLData {
    param(
        [Parameter(Mandatory=$true)] $SQLQuery,
        [Parameter(Mandatory=$true)] $ConnectionString
    )

    # Set up the SQL connection
    $SqlConnection = new-object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = $ConnectionString
    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
    $SqlCommand.CommandText = $SQLQuery
    $SqlCommand.Connection = $SqlConnection
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCommand
    $SqlDataSet = New-Object System.Data.DataSet

    # Run the SQL query
    $SqlConnection.open()
    $SqlAdapter.Fill($SqlDataSet)
    $SqlConnection.close()

    foreach($DSTable in $SqlDataSet.Tables) {
        return $DSTable
    }
    return $null
}

function Write-Log
{
    param(
        [Parameter(Mandatory=$true)] $Message
    )

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss K")> $Message"
}

###########################################################################
# Script initialization                                                   #
###########################################################################

if ($DryRun -eq $true) {
    Write-Log "Performing dry run - will not actually commit changes to the database"
}

Write-Log "Loading config file..."

# Find the config file
$AdjustedConfigFilePath = $ConfigFilePath
if ($AdjustedConfigFilePath.Length -le 0) {
    $AdjustedConfigFilePath = join-path -Path $(Split-Path (Split-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) -Parent) -Parent) -ChildPath "config.xml"
}

# Retreive the connection string from config.xml
if ((test-path -Path $AdjustedConfigFilePath) -eq $false) {
    Throw "Config file not found. Specify using -ConfigFilePath. Defaults to config.xml in the directory above where this script is run from."
}
$configXML = [xml](Get-Content $AdjustedConfigFilePath)
$DBConnectionString = $configXML.Settings.SchoolLogic.ConnectionStringRW

if($DBConnectionString.Length -lt 1) {
    Throw "Connection string was not present in config file. Cannot continue - exiting."
    exit
}

###########################################################################
# Check if the import file exists before going any further                #
###########################################################################
if (Test-Path $InputFileName)
{
} else {
    Write-Log "Couldn't load the input file! Quitting."
    exit
}

###########################################################################
# Collect required info from the SL database                              #
###########################################################################

Write-Log "Loading required data from SchoolLogic DB..."

$SQLQuery_AllCourses = "SELECT iCourseID as ID, cName, cCourseCode, cGovernmentCode, nHighCredit FROM Course ORDER BY iCourseID;"
$SQLQuery_AllSections = "SELECT iClassID as ID, iDefault_StaffID FROM Class ORDER BY iClassID;"
$SQLQuery_AllTeachers = "SELECT iStaffID as ID, CONCAT(cLastName, ', ', cFirstName) as TeacherName FROM Staff ORDER BY iStaffID;"
$SQLQuery_AllClassTerms = "SELECT iClassID as ID, iTermID FROM ClassTerm ORDER BY iClassID;"
$SQLQuery_AllTerms = "SELECT iTermID as ID, dStartDate, dEndDate, cName FROM Term ORDER BY iTermID;"
$SQLQuery_AllStudents = "SELECT iStudentID as ID, iGradesID FROM Student;"

# Convert to hashtables for easier consumption
$AllCourses = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_AllCourses
$AllSections = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_AllSections
$AllTeachers = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_AllTeachers
$AllClassTerms = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_AllClassTerms
$AllTerms = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_AllTerms
$AllStudents = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_AllStudents

Write-Log " Loaded $($AllCourses.Count) courses from SchoolLogic DB."
Write-Log " Loaded $($AllSections.Count) sections from SchoolLogic DB."
Write-Log " Loaded $($AllTeachers.Count) teachers from SchoolLogic DB."
Write-Log " Loaded $($AllClassTerms.Count) classterms from SchoolLogic DB."
Write-Log " Loaded $($AllTerms.Count) terms from SchoolLogic DB."

###########################################################################
# Generate a unique ID for this batch                                     #
###########################################################################

$BatchThumbprint = Get-Hash "BatchThumbprint$(Get-Date)"

###########################################################################
# Process the file                                                        #
###########################################################################

Write-Log "Processing input file..."

$RecordsToInsert = @()
$MarksForInvalidCourses = @()
$MarksForInvalidSections = @()

foreach ($InputRow in Get-CSV -CSVFile $InputFileName)
{
    $CourseCode = [string]$InputRow.CourseCode
    $Course = Get-ByID -ID $CourseCode -Haystack $AllCourses
    
    $TeacherName = ""
    $SectionID = [int]$(Convert-SectionID -InputString $([string]$InputRow.SectionGUID) -SchoolID $InputRow.SchoolID)    
    $Section = Get-ByID -ID $SectionID -Haystack $AllSections
    $StudentID = Convert-StudentID -InputString $([string]$InputRow.StudentGUID)
    $Student = Get-ByID -ID $StudentID -Haystack $AllStudents
        
    $GradesID = 0
    $StartDate = '1900-01-01'
    $EndDate = '1900-01-01'
    $cTerm = ""

    if ($null -ne $Student) {
        $GradesID = $Student.iGradesID
    }

    if ($null -ne $Section) {
        $Teacher = Get-ByID -ID $Section.iDefault_StaffID -Haystack $AllTeachers
        if ($null -ne $Teacher) {
            $TeacherName = $Teacher.TeacherName
        }

        $ClassTerm = Get-ByID -ID $Section.ID -Haystack $AllClassTerms

        
        if ($null -ne $ClassTerm) {
            $Term = Get-ByID -ID $ClassTerm.iTermID -Haystack $AllTerms

            if ($null -ne $Term) {
                $StartDate = $Term.dStartDate
                $EndDate = $Term.dEndDate
                $cTerm = $Term.cName
            }
        }
        
    }

    if ($null -eq $Course) {
        Write-Log "COURSE NOT FOUND: $CourseCode"
        $MarksForInvalidCourses += $InputRow
        continue
    }

    if ($null -eq $Section) {
        Write-Log "SECTION NOT FOUND: $SectionID"
        $MarksForInvalidSections += $InputRow
    }

    $NewMark = [PSCustomObject]@{
        iStudentID = $StudentID
        nFinalMark = [string]$InputRow.FinalGrade        
        nYear = Convert-Year -InputString $([string]$InputRow.YearID)
        iSchoolID = [int]$InputRow.SchoolID
        ImportBatchID = $BatchThumbprint
        iCourseID = [string]$InputRow.CourseCode
        cCourseCode = $Course.cGovernmentCode
        cCourseDesc = $Course.cName
        nCreditPossible = [int]$Course.nHighCredit
        nCreditEarned = 0
        cTeacher = $TeacherName
        cTerm = $cTerm
        dStartDate = $StartDate
        dEndDate = $EndDate
        iGradesID = $GradesID
    }     

    if (($NewMark.nFinalMark -eq "IE") -or ($NewMark.nFinalMark.Length -le 1)) {        
        continue
    } else {        
        $NewMark.nCreditEarned = Get-CreditEarned -FinalMark $([int]$InputRow.FinalGrade) -PossibleCredits $([int]$Course.nHighCredit)
        #write-host $NewMark
        $RecordsToInsert += $NewMark
    }   
}

###########################################################################
# Show marks for invalid courses                                          #
###########################################################################

Write-Log "To insert: $($RecordsToInsert.Count)"
Write-Log "Marks for courses that don't exist: $($MarksForInvalidCourses.Count)"
Write-Log "Marks for sections that don't exist: $($MarksForInvalidSections.Count)"

###########################################################################
# Perform SQL operations                                                  #
###########################################################################

# Set up the SQL connection
$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $DBConnectionString

if ($DryRun -ne $true) {

Write-Log "Inserting $($RecordsToInsert.Count) records..."
    foreach ($NewRecord in $RecordsToInsert) {
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
        $SqlCommand.CommandText = "INSERT INTO MarksHistory(iStudentID,nFinalMark,nYear,iCourseID,ImportBatchID,cCourseCode,cCourseDesc,nCreditPossible,nCreditEarned,cActionCode,iSchoolID,ImportTimestamp,cTeacher,dModified,cTerm,dStartDate,dEndDate,iGradesID,iLV_CompletionStatusID)
                                        VALUES(@STUDENTID, @FINALMARK,@YEAR,@COURSEID,@BATCHID,@COURSECODE,@COURSEDESC,@CREDITPOSSIBLE,@CREDITEARNED,@ACTIONCODE,@SCHOOLID,@IMPORTTIMESTAMP,@TEACHERNAME,@MODIFIEDDATE,@TERMNAME,@STARTDATE,@ENDDATE,@GRADESID,@COMPLETIONSTATUS);"
        $SqlCommand.Parameters.AddWithValue("@STUDENTID",$NewRecord.iStudentID) | Out-Null    
        $SqlCommand.Parameters.AddWithValue("@FINALMARK",$NewRecord.nFinalMark) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@YEAR",$NewRecord.nYear) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@COURSEID",$NewRecord.iCourseID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@COURSECODE",$NewRecord.cCourseCode) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@COURSEDESC",$NewRecord.cCourseDesc) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@BATCHID",$NewRecord.ImportBatchID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@CREDITPOSSIBLE",$NewRecord.nCreditPossible) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@CREDITEARNED",$NewRecord.nCreditEarned) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@ACTIONCODE", 'A') | Out-Null
        $SqlCommand.Parameters.AddWithValue("@SCHOOLID",$NewRecord.iSchoolID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@IMPORTTIMESTAMP",$(Get-Date)) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@MODIFIEDDATE",$(Get-Date)) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@TEACHERNAME",$NewRecord.cTeacher) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@TERMNAME",$NewRecord.cTerm) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@STARTDATE",$NewRecord.dStartDate) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@ENDDATE",$NewRecord.dEndDate) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@GRADESID",$NewRecord.iGradesID) | Out-Null 
        $SqlCommand.Parameters.AddWithValue("@COMPLETIONSTATUS",$CompletionStatusID) | Out-Null
        $SqlCommand.Connection = $SqlConnection

        $SqlConnection.open()
        $Sqlcommand.ExecuteNonQuery() | Out-Null
        $SqlConnection.close()
    }
} else {
    Write-Log "Skipping SQL operation due to -DryRun"
}