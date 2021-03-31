param (
    [Parameter(Mandatory=$true)][string]$InputFileName,
    [string]$ConfigFilePath,
    [switch]$DryRun
 )

###########################################################################
# Functions                                                               #
###########################################################################

import-module ./EdsbyImportModule.psm1 -Scope Local

###########################################################################
# Script initialization                                                   #
###########################################################################

$RequiredCSVColumns = @(
    "ReportingPeriodEndDate",
    "ReportingPeriodName",
    "StudentGUID",
    "SchoolID",
    "OverallMark",
    "SectionGUID"
)

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

try {
    $CSVInputFile = Get-CSV -CSVFile $InputFileName -RequiredColumns $RequiredCSVColumns
}
catch {
    Write-Log("ERROR: $($_.Exception.Message)")
    remove-module edsbyimportmodule
    exit
}

###########################################################################
# Collect required info from the SL database                              #
###########################################################################

Write-Log "Loading required data from SchoolLogic DB..."

Write-Log "Loading and processing class report periods..."
$ClassReportPeriods = Get-ClassReportPeriods -DBConnectionString $DBConnectionString
Write-Log " Loaded report periods for $($ClassReportPeriods.Keys.Count) classes."

Write-Log "Loading class credits..."
$ClassCredits =  Get-AllClassCredits -DBConnectionString $DBConnectionString
Write-Log " Loaded class credit information for $($ClassCredits.Keys.Count) classes."

###########################################################################
# Process the file                                                        #
###########################################################################

Write-Log "Processing input file..."
Write-Log " Rows to process: $($CSVInputFile.Length)"
$MarksToImport = New-Object -TypeName "System.Collections.ArrayList"

$FoundMarkClassesByStudentID = @{}
$OMProcessCounter = 0
foreach ($InputRow in $CSVInputFile)
{
    # If there is no grade, ignore
    if ($InputRow.OverallMark -eq "") {
        continue;
    }

    # Parse the marks
    $NewMark = Convert-ToSLMark -InputRow $InputRow -AllReportPeriods $ClassReportPeriods -AllClassCredits $ClassCredits

    # Check if we already know about this mark
    if ($FoundMarkClassesByStudentID.ContainsKey($NewMark.iStudentID) -eq $false) {
        $FoundMarkClassesByStudentID.Add($NewMark.iStudentID, @{})
    }

    if ($FoundMarkClassesByStudentID[$NewMark.iStudentID].Contains($NewMark.iReportPeriodID) -eq $false) {
        $FoundMarkClassesByStudentID[$NewMark.iStudentID].Add($NewMark.iReportPeriodID, (New-Object -TypeName "System.Collections.ArrayList"))
    }

    if ($FoundMarkClassesByStudentID[$NewMark.iStudentID][$NewMark.iReportPeriodID].Contains($NewMark.iClassID) -eq $false) {
        $MarksToImport += $NewMark
        $FoundMarkClassesByStudentID[$NewMark.iStudentID][$NewMark.iReportPeriodID] += $NewMark.iClassID
    }
     
    
    $OMProcessCounter++
    $PercentComplete = [int]([decimal]($OMProcessCounter/$CSVInputFile.Length) * 100)
    if ($PercentComplete % 5 -eq 0) {
        Write-Progress -Activity "Processing input file" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
    }
}

Write-Log "Found $($MarksToImport.Length) marks to import"

$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $DBConnectionString

###########################################################################
# Import the marks                                                        #
###########################################################################

Write-Log "Inserting class marks into SchoolLogic..."
if ($DryRun -ne $true) {
    $OMInsertCounter = 0
    foreach($M in $MarksToImport) {
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
        $SqlCommand.CommandText = " UPDATE Marks 
                                    SET 
                                        nMark=@NMARK, 
                                        cMark=@CMARK,
                                        nCredit=@NCREDIT
                                    WHERE 
                                        iStudentID=@STUDENTID 
                                        AND iClassID=@CLASSID 
                                        AND iReportPeriodID=@REPID 
                                        AND NOT (nMark=0 AND cMark='')
                                    IF @@ROWCOUNT = 0 
                                    INSERT INTO 
                                        Marks(iStudentID, iReportPeriodID, iClassID, nMark, cMark, nCredit, dDateAssigned, iSchoolID, ImportTimestamp, ImportBatchID)
                                        VALUES(@STUDENTID, @REPID, @CLASSID, @NMARK, @CMARK, @NCREDIT, @DDATEASS, @SCHOOLID, @DDATEASS, @DDATEASS);"
        
        $SqlCommand.Parameters.AddWithValue("@STUDENTID",$M.iStudentID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@REPID",$M.iReportPeriodID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@CLASSID",$M.iClassID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@NMARK",$M.nMark) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@CMARK",$M.cMark) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@NCREDIT",$M.nCredit) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@SCHOOLID",$M.iSchoolID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@DDATEASS",$(Get-Date)) | Out-Null
        $SqlCommand.Connection = $SqlConnection

        $SqlConnection.open()
        if ($DryRun -ne $true) {
            $Sqlcommand.ExecuteNonQuery() | Out-Null
        } 
        $SqlConnection.close()

        $OMInsertCounter++
        $PercentComplete = [int]([decimal]($OMInsertCounter/$MarksToImport.Count) * 100)
        if ($PercentComplete % 5 -eq 0) {
            Write-Progress -Activity "Inserting marks" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
        }
    }
} else {
    Write-Log "Skipping database write due to -DryRun"
}
Write-Log "Done!"
remove-module EdsbyImportModule