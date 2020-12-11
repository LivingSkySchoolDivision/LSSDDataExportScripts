<#
.SYNOPSIS
Imports and processes an Edsby Report Card Export CSV into the SchoolLogic Database.

.DESCRIPTION
This script processes the Edsby Report Card Export CSV file and imports its contents into the Marks table of your SchoolLogic/SIRS database.

This script relies on ID numbers set up in a specific way between your SchoolLogic/SIRS database and Edsby. If you are not running this on Living Sky School Division data, you will likely need to edit this script to suit your environment.

Assumptions that this script makes:
 - Your SectionGUIDs in Edsby are iSchoolID-iClassID, corresponding to the iSchoolID and iClassID values in SchoolLogic.
 - Your StudentGUIDs are STUDENT-iStudentID, where STUDENT is literally the word "STUDENT" in capital letters, and iStudentID corresponds to the iStudentID in SchoolLogic.

To adjust how this script parses your Section and Student Ids, edit the Convert-StudentID and/or Convert-SectionID functions below.

You will need to make the following additions to your SchoolLogic database:
In the Marks table, you must add two fields:
 - ImportTimeStamp - datetime
 - ImportBatchID - varchar(40)

These fields help you troubleshoot or roll back any potential errors, and provide a way to delete database changes that this script creates. Each time this script writes to your database, rows it adds will include a unique ID hash as well as a timestamp. To "undo" an import, find the batch ID and remove any entries in your Marks table with that value in the ImportBatchID field.

.PARAMETER ConfigFilePath
Specify the path to the config file. Config file should be an XML file that contains the database connection string. Example config file can be found below.

Example config.xml:
<?xml version="1.0" encoding="utf-8" ?>
<Settings>
  <SchoolLogic>
    <ConnectionStringRW>data source=SERVERNAME;initial catalog=DATABASENAME;user id=USERNAME;password=PASSWORD;Trusted_Connection=false</ConnectionStringRW><!-- Read/Write-->
  </SchoolLogic>  
</Settings>

.PARAMETER EmptyMarksLogPath
If there are any empty marks in your import file, they will be dumped into this CSV file. If this parameter is not specified, empty marks in the import file will be silently ignored. If there are no empty marks, no file will be created.

.PARAMETER OrphanedMarksLogPath
If there are marks in your import file for classes/sections that do not exist in your SchoolLogic/SIRS database, they will be logged to this CSV file. If this parameter is not specified, these "orphaned" marks will be silently ignored. If there are no orphaned marks, no file will be created.

.PARAMETER ErrorLogPath
If there are parsing errors when reading lines from the import file, the lines will be written to this CSV file. If this parameter is not specified, lines causing errors or exceptions will be ignored. If there are no errors, no file will be created.

.PARAMETER Commit
Enable writing to the database. Disabled by default, as a safety mechanism. To enable writing to the database, use "-Commit $true"

.EXAMPLE
PS> .\EdsbyMarksToSchoolLogicMarks.ps1 -ConfigFilePath config.xml -inputfilename import.csv -EmptyMarksLogPath empty.csv -OrphanedMarksLogPath orphans.csv -ErrorLogPath errors.csv -Commit $false

The above example imports a file and will identify any errors either in the console, or in one of three potential CSV files that are created.

.EXAMPLE
PS> .\EdsbyMarksToSchoolLogicMarks.ps1 -ConfigFilePath config.xml -inputfilename import.csv -Commit $true

Having dealt with all of the issues, the above example will import the file into the database.

.LINK
https://github.com/LivingSkySchoolDivision/LSSDDataExportScripts
https://github.com/LivingSkySchoolDivision
https://www.lskysd.ca

#>
param (
    [Parameter(Mandatory=$true)][string]$InputFileName,
    [string]$ConfigFilePath,
    [bool]$Commit,
    [string]$OrphanedMarksLogPath,
    [string]$EmptyMarksLogPath,
    [string]$ErrorLogPath
 )

###########################################################################
# CHANGES YOU NEED TO MAKE TO YOUR DATABASE                               #
###########################################################################

# Add the following fields to your Marks table
#  ImportBatchID - varchar(40)
#  ImportTimestamp - datetime

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

function Convert-StudentID {
    param(
        [Parameter(Mandatory=$true)] $InputString
    )

    return [int]$InputString.Replace("STUDENT-","")
}

function Convert-EarnedCredits {
    param(
        [Parameter(Mandatory=$true)] $InputString,
        [Parameter(Mandatory=$true)][int]$PotentialCredits
    )

    try {
        $parsed = [int]$InputString
        if ($parsed -gt 49) {
            return $PotentialCredits
        } else {
            return 0
        }

    }
    catch {
        # We couldn't parse a number, so assume its a non-mark mark (like IE or NYM) or an error
        return 0
    }
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

if ($Commit -ne $true) {
    Write-Log "Performing dry run - will not actually commit changes to the database. Enable writing to database with -Commit `$true"
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

$SQLQuery_AllSections = "SELECT iClassID as ID, iCourseID FROM Class ORDER BY iClassID;"
$SQLQuery_ReportPeriods = "SELECT iClassID as ID, ClassPeriod.iReportPeriodID FROM ClassPeriod LEFT OUTER JOIN ReportPeriod ON ClassPeriod.iReportPeriodID=ReportPeriod.iReportPeriodID WHERE lMoveToHistory=1 ORDER BY ReportPeriod.dEndDate DESC, iClassID ASC;"
$SQLQuery_Courses = "SELECT iCourseID as ID, nHighCredit FROM Course ORDER BY iCourseID;"

# Convert to hashtables for easier consumption
$AllSections = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_AllSections
$ReportPeriodsToHistory = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_ReportPeriods
$AllCourses = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_Courses

Write-Log " Loaded $($AllSections.Count) sections from SchoolLogic DB."
Write-Log " Loaded $($ReportPeriodsToHistory.Count) report periods (that move to history) from SchoolLogic DB."
Write-Log " Loaded $($AllCourses.Count) courses from SchoolLogic DB."

###########################################################################
# Generate a unique ID for this batch                                     #
###########################################################################

$BatchThumbprint = Get-Hash "BatchThumbprint$(Get-Date)"
Write-Log "ImportBatchID for this script run: $BatchThumbprint"

###########################################################################
# Process the file                                                        #
###########################################################################

Write-Log "Processing input file..."

$RecordsToInsert = @()
$MarksForInvalidSections = @()
$IgnoredEmptyMarks = @()
$ErrorRows = @()

foreach ($InputRow in Get-CSV -CSVFile $InputFileName)
{
    try {
        # Ignore this mark if it's empty
        if ($InputRow.FinalGrade.Length -lt 1) {
            $IgnoredEmptyMarks += $InputRow
            continue
        }

        $SectionID = [int]$(Convert-SectionID -InputString $([string]$InputRow.SectionGUID) -SchoolID $InputRow.SchoolID)
        $Section = Get-ByID -ID $SectionID -Haystack $AllSections
        $iReportPeriodID = 0
        $nCredits = 0

        if ($null -eq $Section) {
            Write-Log "SECTION NOT FOUND: $SectionID"
            $MarksForInvalidSections += $InputRow
        } else {
            $ReportPeriod = Get-ByID -ID $SectionID -Haystack $ReportPeriodsToHistory

            foreach($RP in $ReportPeriod) {
                $iReportPeriodID = $RP.iReportPeriodID
            }

            $Course = Get-ByID -ID $Section.iCourseID -Haystack $AllCourses
            if ($null -ne $Course) {
                $nCredits = Convert-EarnedCredits -PotentialCredits $Course.nHighCredit -InputString $InputRow.FinalGrade
            }
        }

        # Parse the mark into the appropriate field
        $nMark = 0
        $cMark = ""
        try {
            $parsedMark = [int]$InputRow.FinalGrade
            $nMark = $parsedMark            
        }
        catch {
            $cMark = $InputRow.FinalGrade
        }

        # If the mark is empty after we parse it, count it as an ignored empty mark and skip it
        if (($nMark -eq 0) -and ($cMark -eq "")) {
            $IgnoredEmptyMarks += $InputRow
            continue
        }

        $NewMark = [PSCustomObject]@{
            iClassID =  $SectionID
            iStudentID = Convert-StudentID -InputString $([string]$InputRow.StudentGUID)
            cMark = $cMark
            nMark = $nMark
            iSchoolID = [int]$InputRow.SchoolID
            ImportBatchID = $BatchThumbprint
            mComment = [string]$InputRow.Comment
            nCredit = $nCredits
            dDateAssigned = $(Get-Date)
            iReportPeriodID = $iReportPeriodID
        }

        $RecordsToInsert += $NewMark
    }
    catch {
        $ErrorRows += $InputRow
    }
}

###########################################################################
# Show marks for invalid courses                                          #
###########################################################################

Write-Log "To insert: $($RecordsToInsert.Count)"

Write-Log "Ignored empty marks: $($IgnoredEmptyMarks.Count) (use -EmptyMarksLogPath <filename> to dump these to a csv)."
if ($EmptyMarksLogPath.Length -gt 0) {
    Write-Log " Empty marks log written to file: $EmptyMarksLogPath"
    $IgnoredEmptyMarks | Export-CSV $EmptyMarksLogPath
}

Write-Log "Marks for sections that don't exist: $($MarksForInvalidSections.Count)  (use -OrphanedMarksLogPath <filename> to dump these to a csv)."
if ($OrphanedMarksLogPath.Length -gt 0) {
    Write-Log " Log written to file: $OrphanedMarksLogPath"
    $MarksForInvalidSections | Export-CSV $OrphanedMarksLogPath
}

Write-Log "Rows that caused errors: $($ErrorRows.Count)  (use -ErrorLogPath <filename> to dump these to a csv)."
if ($ErrorLogPath.Length -gt 0) {
    Write-Log " Log written to file: $ErrorLogPath"
    $ErrorRows | Export-CSV $ErrorLogPath
}

###########################################################################
# Perform SQL operations                                                  #
###########################################################################

# Set up the SQL connection
$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $DBConnectionString

if ($Commit -eq $true) {
Write-Log "Inserting $($RecordsToInsert.Count) records..."
    foreach ($NewRecord in $RecordsToInsert) {
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
        $SqlCommand.CommandText = "INSERT INTO Marks(ImportBatchID,ImportTimestamp,iStudentID,iClassID,nMark,cMark,dDateAssigned,iSchoolID,mComment,nCredit,iReportPeriodID)
                                        VALUES(@BATCHID,@IMPORTTIMESTAMP,@STUDENTID,@CLASSID,@NMARK,@CMARK,@DATEASSIGN,@SCHOOLID,@MCOMMENT,@NCREDIT,@REPORTPERIODID);"
        $SqlCommand.Parameters.AddWithValue("@STUDENTID",$NewRecord.iStudentID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@CLASSID",$NewRecord.iClassId) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@NMARK",$NewRecord.nMark) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@CMARK",$NewRecord.cMark) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@DATEASSIGN",$NewRecord.dDateAssigned) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@SCHOOLID",$NewRecord.iSchoolID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@MCOMMENT",$NewRecord.mComment) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@NCREDIT",$NewRecord.nCredit) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@BATCHID",$NewRecord.ImportBatchID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@IMPORTTIMESTAMP",$(Get-Date)) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@REPORTPERIODID",$NewRecord.iReportPeriodID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@CREDITS",$NewRecord.nCredit) | Out-Null
        $SqlCommand.Connection = $SqlConnection
        $SqlConnection.open()
        $Sqlcommand.ExecuteNonQuery() | Out-Null
        $SqlConnection.close()
    }
} else {
    Write-Log "Skipping SQL operations. To enable writing to database, add -Commit `$true"
}