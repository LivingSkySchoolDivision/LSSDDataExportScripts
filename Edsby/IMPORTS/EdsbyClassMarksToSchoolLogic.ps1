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
    "ReportingTermNumber",
    "StudentGUID",
    "SchoolID",
    "OverallMark",
    "SectionGUID"
)

if ($DryRun -eq $true) {
    Write-Log "Performing dry run - will not actually commit changes to the database"
}

$SQLQuery_ClassCredits =    "SELECT
                                Class.iClassID,
                                Course.nHighCredit
                            FROM 
                                Class
                                LEFT OUTER JOIN Course ON Class.iCourseID=Course.iCourseID
                            WHERE
                                Course.nHighCredit > 0"

$SQLQuery_ClassReportPeriods = "SELECT 
                                    Class.iClassID,
                                    Track.iTrackID,
                                    ReportPeriod.iReportPeriodID,
                                    ReportPeriod.cName,
                                    ReportPEriod.dStartDate,
                                    ReportPEriod.dEndDate
                                FROM
                                    Class
                                    LEFT OUTER JOIN Track ON Class.iTrackID=Track.iTrackID
                                    LEFT OUTER JOIN Term ON Track.iTrackID=Term.iTrackID
                                    LEFT OUTER JOIN ReportPeriod ON Term.iTermID=ReportPeriod.iTermID
                                WHERE
                                    ReportPeriod.iReportPeriodID IS NOT NULL
                                ORDER BY
                                    Track.iTrackID,
                                    ReportPEriod.dEndDate"


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
$ClassReportPeriods = Convert-ClassReportPeriodsToHashtable -AllClassReportPeriods $(Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_ClassReportPeriods)
Write-Log " Loaded report periods for $($ClassReportPeriods.Keys.Count) classes."

$ClassCredits = Convert-CourseCreditsToHashtable -CourseCreditsDataTable $(Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_ClassCredits)
Write-Log " Loaded class credit information for $($ClassCredits.Keys.Count) classes."

###########################################################################
# Process the file                                                        #
###########################################################################

Write-Log "Processing input file..."
Write-Log " Rows to process: $($CSVInputFile.Length)"
$MarksToImport_Raw = New-Object -TypeName "System.Collections.ArrayList"

$OMProcessCounter = 0
foreach ($InputRow in $CSVInputFile)
{
    # If there is no grade, ignore
    if ($InputRow.OverallMark -eq "") {
        continue;
    }

    # Assemble the final mark object
    $NewMark = Convert-ToSLMark -InputRow $InputRow -AllReportPeriods $ClassReportPeriods -AllClassCredits $ClassCredits

    if (($NewMark.nMark -gt 1) -or ($NewMark.cMark -ne "")) {
        $MarksToImport_Raw += $NewMark
    }
    
    $OMProcessCounter++
    $PercentComplete = [int]([decimal]($OMProcessCounter/$CSVInputFile.Length) * 100)
    if ($PercentComplete % 5 -eq 0) {
        Write-Progress -Activity "Processing input file" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
    }
}

Write-Log "Found $($MarksToImport_Raw.Length) marks to import"

Write-Log "De-duping marks..."
$MarksToImport = New-Object -TypeName "System.Collections.ArrayList"
$FoundMarkClassesByStudentID = @{}

$DeDupeCounter = 0
foreach($Mark in $MarksToImport_Raw) {

    try {
        if ($FoundMarkClassesByStudentID.ContainsKey($Mark.iStudentID) -eq $false) {
            $FoundMarkClassesByStudentID.Add($Mark.iStudentID, @{})
        }

        if ($FoundMarkClassesByStudentID[$Mark.iStudentID].Contains($Mark.iReportPeriodID) -eq $false) {
            $FoundMarkClassesByStudentID[$Mark.iStudentID].Add($Mark.iReportPeriodID, (New-Object -TypeName "System.Collections.ArrayList"))
        }

        if ($FoundMarkClassesByStudentID[$Mark.iStudentID][$Mark.iReportPeriodID].Contains($Mark.iClassID) -eq $false) {
            
            $MarksToImport += $Mark
            $FoundMarkClassesByStudentID[$Mark.iStudentID][$Mark.iReportPeriodID] += $Mark.iClassID
        }

        $DeDupeCounter++
        $PercentComplete = [int]([decimal]($DeDupeCounter/$CSVInputFile.Length) * 100)
        if ($PercentComplete % 5 -eq 0) {
            Write-Progress -Activity "De-Duping" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
        }
    }
    catch {
        write-log "EXCEPTION: $($Mark)"
    }
}

Write-Log "De-duped to $($MarksToImport.Length) marks to import."

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