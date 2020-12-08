param (
    [Parameter(Mandatory=$true)][string]$InputFileName,
    [string]$ConfigFilePath,
    [bool]$PerformDeletes,
    [bool]$DryDelete,
    [bool]$DryRun,
    [bool]$AllowSyncToEmptyTable
 )

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

function Get-Course {
    param(
        [Parameter(Mandatory=$true)] $iCourseID,
        [Parameter(Mandatory=$true)] $Courses
    )

    # SELECT iCourseID, cName, cCourseCode, cGovernmentCode, nHighCredit FROM Course;
    foreach($Course in $Courses) {
        if ($Course.iCourseID -eq $CourseCode) {
            return $Course
        }
    }

    return $null
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


##

###########################################################################
# Collect required info from the SL database                              #
###########################################################################

Write-Log "Loading required data from SchoolLogic DB..."

$SQLQuery_AllCourses = "SELECT iCourseID, cName, cCourseCode, cGovernmentCode, nHighCredit FROM Course;"

# Convert to hashtables for easier consumption
$AllCourses = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_AllCourses

Write-Log " Loaded $($AllCourses.Count) courses from SchoolLogic DB."

###########################################################################
# Generate a unique ID for this batch                                     #
###########################################################################

$BatchThumbprint = Get-Hash "BatchThumbprint$(Get-Date)"

###########################################################################
# Process the file                                                        #
###########################################################################

Write-Log "Processing input file..."

$AttendanceToImport = @{}

$RecordsToInsert = @()
$MarksForInvalidCourses = @()

foreach ($InputRow in Get-CSV -CSVFile $InputFileName)
{
    $CourseCode = [string]$InputRow.CourseCode
    $Course = Get-Course -iCourseID $CourseCode -Courses $AllCourses

    if ($Course -eq $null) {
        Write-Log "COURSE NOT FOUND: $CourseCode"
        $MarksForInvalidCourses += $InputRow
    }

    $NewMark = [PSCustomObject]@{
        iStudentID = Convert-StudentID -InputString $([string]$InputRow.StudentGUID)
        cCourseCode = [string]$InputRow.CourseCode
        nFinalMark = [string]$InputRow.FinalGrade        
        nYear = Convert-Year -InputString $([string]$InputRow.YearID)
        iSchoolID = [int]$InputRow.SchoolID
        ImportBatchID = $BatchThumbprint
    }     

    if (($NewMark.nFinalMark -eq "IE") -or ($NewMark.nFinalMark.Length -le 1)) {        
        continue
    } else {
        #write-host $NewMark
        $RecordsToInsert += $NewMark
    }   
}

exit
###########################################################################
# Show marks for invalid courses                                          #
###########################################################################



###########################################################################
# Insert into database                                                    #
###########################################################################

Write-Log "To insert: $($RecordsToInsert.Count)"

###########################################################################
# Perform SQL operations                                                  #
###########################################################################

# Set up the SQL connection
$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $DBConnectionString

Write-Log "Inserting $($RecordsToInsert.Count) records..."
foreach ($NewRecord in $RecordsToInsert) {
    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
    $SqlCommand.CommandText = "INSERT INTO MarksHistory(iStudentID,nFinalMark,nYear,cCourseCode,ImportBatchID)
                                    VALUES(@STUDENTID, @FINALMARK,@YEAR,@COURSECODE,@BATCHID);"
    $SqlCommand.Parameters.AddWithValue("@STUDENTID",$NewRecord.iStudentID) | Out-Null    
    $SqlCommand.Parameters.AddWithValue("@FINALMARK",$NewRecord.nFinalMark) | Out-Null
    $SqlCommand.Parameters.AddWithValue("@YEAR",$NewRecord.nYear) | Out-Null
    $SqlCommand.Parameters.AddWithValue("@COURSECODE",$NewRecord.cCourseCode) | Out-Null
    $SqlCommand.Parameters.AddWithValue("@BATCHID",$NewRecord.ImportBatchID) | Out-Null
    $SqlCommand.Connection = $SqlConnection

    $SqlConnection.open()
    if ($DryRun -ne $true) {
        $Sqlcommand.ExecuteNonQuery()
    } else {
        Write-Log " (Skipping SQL query due to -DryRun)"
    }
    $SqlConnection.close()

    break
}
