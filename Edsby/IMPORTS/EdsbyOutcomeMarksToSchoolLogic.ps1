param (
    [Parameter(Mandatory=$true)][string]$InputFileName,
    [string]$ConfigFilePath,
    [switch]$ImportUnknownOutcomes,
    [switch]$DryRun
 )

###########################################################################
# Functions                                                               #
###########################################################################

function Write-Log
{
    param(
        [Parameter(Mandatory=$true)] $Message
    )

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss K")> $Message"
}

function Validate-CSV {
    param(
        [Parameter(Mandatory=$true)][String] $CSVFile
    )
    # Make sure the CSV has all the required columns for what we need

    $line = Get-Content $CSVFile -first 1

    # Check if the first row contains headings we expect
    if ($line.Contains('"SchoolID"') -eq $false) { throw "Input CSV missing field: SchoolID" }
    if ($line.Contains('"StudentGUID"') -eq $false) { throw "Input CSV missing field: StudentGUID" }
    if ($line.Contains('"CourseCode"') -eq $false) { throw "Input CSV missing field: CourseCode" }
    if ($line.Contains('"CriterionName"') -eq $false) { throw "Input CSV missing field: CriterionName" }
    if ($line.Contains('"Grade"') -eq $false) { throw "Input CSV missing field: Grade" }
    if ($line.Contains('"SectionGUID"') -eq $false) { throw "Input CSV missing field: SectionGUID" }    
    return $true
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

function Get-CSV {
    param(
        [Parameter(Mandatory=$true)][String] $CSVFile
    )

    if ((Validate-CSV $CSVFile) -eq $true) {
        return import-csv $CSVFile  | Select -skip 1
    } else {
        throw "CSV file is not valid - cannot continue"
    }
}

function Convert-ToSLOutcomeMark {
    param(
        [Parameter(Mandatory=$true)] $InputRow
    )
    # Parse cMark vs nMark
    $cMark = ""
    $nMark = ""

    if ([bool]($InputRow.Grade -as [decimal]) -eq $true) {
        $nMark = [decimal]$InputRow.Grade
        if (
            ($nMark -eq 1) -or
            ($nMark -eq 1.5) -or
            ($nMark -eq 2) -or
            ($nMark -eq 2.5) -or
            ($nMark -eq 3) -or
            ($nMark -eq 3.5) -or
            ($nMark -eq 4)
        ) {
            $cMark = [string]$nMark
        }
    } else {
        $cMark = $InputRow.Grade
    }

    return [PSCustomObject]@{
        iStudentID = [int](Convert-StudentID $InputRow.StudentGUID)
        iReportPeriodID = [int]0
        iCourseObjectiveId = [int](Convert-ObjectiveID -OutcomeCode $InputRow.CriterionName -Objectives $SLCourseObjectives -iCourseID $InputRow.CourseCode)
        iCourseID = [int]$InputRow.CourseCode
        iSchoolID = [int]$InputRow.SchoolID
        nMark = [decimal]$nMark
        cMark = [string]$cMark
    }
}

function Convert-StudentID {
    param(
        [Parameter(Mandatory=$true)] $InputString
    )

    return [int]$InputString.Replace("STUDENT-","")
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

function Convert-ObjectiveID {
    param(
        [Parameter(Mandatory=$true)] $OutcomeCode,
        [Parameter(Mandatory=$true)] $iCourseID,
        [Parameter(Mandatory=$true)] $Objectives
    )
    foreach($obj in $Objectives) {
        if (($obj.OutcomeCode -eq $OutcomeCode) -and ($obj.iCourseID -eq $iCourseID))
        {
            return $obj.iCourseObjectiveID
        }
    } 

    return -1
}

###########################################################################
# Script initialization                                                   #
###########################################################################

if ($DryRun -eq $true) {
    Write-Log "Performing dry run - will not actually commit changes to the database"
}

if ($ImportUnknownOutcomes -eq $true) {
    Write-Log "When encountering a mark for an outcome that doesn't exist in SchoolLogic, script will: Import those outcomes into SchoolLogic"
} else {
    Write-Log "When encountering a mark for an outcome that doesn't exist in SchoolLogic, script will: Ignore those marks"
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
# Load the given CSV in, but don't process it yet                         #
###########################################################################

Write-Log "Loading and validating input CSV file..."
$CSVInputFile = Get-CSV -CSVFile $InputFileName

###########################################################################
# Collect required info from the SL database                              #
###########################################################################

Write-Log "Loading required data from SchoolLogic DB..."

$SQLQuery_CourseObjectives = "SELECT iCourseObjectiveID, OutcomeCode, iCourseID, cSubject FROM CourseObjective"

# Convert to hashtables for easier consumption
$SLCourseObjectives_Raw = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_CourseObjectives
Write-Log " Loaded $($SLCourseObjectives_Raw.Length) course objectives."

# Put course objectives in a hashtable for easy lookups
$SLCourseObjectives = @()

foreach($Obj in $SLCourseObjectives_Raw) {
    if (($Obj.OutcomeCode -ne "") -and ($null -ne $Obj.OutcomeCode)) {
            $Outcome = [PSCustomObject]@{
            OutcomeCode = $Obj.OutcomeCode
            OutcomeText = $Obj.OutcomeText
            iCourseObjectiveID = $Obj.iCourseObjectiveID
            iCourseID = $Obj.iCourseID
            cSubject = $Obj.cSubject
        }
        $SLCourseObjectives += $Outcome
        
    }
}

###########################################################################
# Process the file                                                        #
###########################################################################

Write-Log "Processing input file..."
Write-Log " Rows to process: $($CSVInputFile.Length)"
$OutcomeMarksToImport = @()
$OutcomeMarksNeedingOutcomes = @()
$OutcomeNotFound = @{}

foreach ($InputRow in $CSVInputFile)
{
    # If there is no grade, ignore
    if ($InputRow.Grade -eq "") {
        continue;
    }

    # Assemble the final mark object
    $NewOutcomeMark = Convert-ToSLOutcomeMark -InputRow $InputRow

    if ($NewOutcomeMark.iCourseObjectiveId -eq -1) {
        $OutcomeMarksNeedingOutcomes += $InputRow
        $Fingerprint = (Get-Hash -String ("$($InputRow.CourseCode)$($InputRow.CriterionName)"))
        if ($OutcomeNotFound.ContainsKey($Fingerprint) -eq $false) {
            $OutcomeNotFound.Add($Fingerprint,[PSCustomObject]@{
                iCourseID = [int]$InputRow.CourseCode
                OutcomeCode = [string]$InputRow.CriterionName
                OutcomeText = [string]$InputRow.CriterionDesc
                cSubject = "$($InputRow.CriterionName) $($InputRow.CriterionDesc)"
                mNotes = "$($InputRow.CriterionName) $($InputRow.CriterionDesc)"
                iLV_ObjectiveCategoryID = 4146
            })
        }
    } else {
        $OutcomeMarksToImport += ($NewOutcomeMark)
    }
}

Write-Log "Found $($OutcomeMarksToImport.Length) marks to import"
Write-Log "Found $($OutcomeMarksNeedingOutcomes.Length) without matching outcomes in SchoolLogic"
Write-Log "Found $($($OutcomeNotFound.Count)) outcomes that don't exist in our database."

if ($ImportUnknownOutcomes -eq $true) {
# Insert new outcomes that didn't exist in SL before

# Re-import outcomes from SchoolLogic

# Reprocess marks that didn't have matching outcomes before
} else {
    Write-Log "Skipping $($OutcomeMarksNeedingOutcomes.Length) marks due to missing outcomes."
}
