param (
    [Parameter(Mandatory=$true)][string]$InputFileName,
    [string]$ConfigFilePath,
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
# Load the given CSV in, but don't process it yet                         #
###########################################################################

Write-Log "Loading and validating input CSV file..."
$CSVInputFile = Get-CSV -CSVFile $InputFileName

###########################################################################
# Process the file                                                        #
###########################################################################

Write-Log "Processing input file..."
Write-Log " Rows to process: $($CSVInputFile.Length)"
$OutcomeMarksToImport = @()

foreach ($InputRow in $CSVInputFile)
{
    # If there is no grade, ignore
    if ($InputRow.Grade -eq "") {
        continue;
    }

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

    # Assemble the final mark object
    $NewOutcomeMark = [PSCustomObject]@{
        iStudentID = [int](Convert-StudentID $InputRow.StudentGUID)
        iReportPeriodID = [int]0
        iCourseObjectiveId = [int]0 # Look up from objective hashtable
        iCourseID = [int]0 # Lookup from course list
        iSchoolID = [int]$InputRow.SchoolID
        nMark = [decimal]$nMark
        cMark = [string]$cMark
    }

    $OutcomeMarksToImport += ($NewOutcomeMark)
}


Write-Log "[DEBUG] Outputting marks to console..."
foreach($OM in $OutcomeMarksToImport) {
    Write-Log $OM
}
Write-Log "Total marks to import: $($OutcomeMarksToImport.Length)"