param (
    [Parameter(Mandatory=$true)][string]$InputFileName,
    [string]$ConfigFilePath,
    [switch]$ImportUnknownOutcomes,
    [switch]$DryRun
 )

###########################################################################
# Functions                                                               #
###########################################################################

import-module $PSScriptRoot/EdsbyImportModule.psm1 -Scope Local

###########################################################################
# Script initialization                                                   #
###########################################################################

$RequiredCSVColumns = @(
    "ReportingPeriodEndDate",
    "ReportingPeriodName",
    "StudentGUID",
    "SchoolID",
    "Citizenship",
    "Collaboration",
    "Engagement",
    "SelfDirected",
    "SectionGUID"
)

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

Write-Log "Loading and processing course objectives..."
$SLCourseObjectives = Get-CourseObjectives -DBConnectionString $DBConnectionString
Write-Log " Loaded and Processed $($SLCourseObjectives.Length) course objectives."

Write-Log "Loading and processing class report periods..."
$ClassReportPeriods = Get-ClassReportPeriods -DBConnectionString $DBConnectionString
Write-Log " Loaded report periods for $($ClassReportPeriods.Keys.Count) classes."

Write-Log "Loading and processing course grades..."
$CourseGrades = Get-CourseGrades -DBConnectionString $DBConnectionString
Write-Log " Loaded $($CourseGrades.Keys.Count) course grades."

###########################################################################
# Process the file                                                        #
###########################################################################

Write-Log "Processing input file..."
Write-Log " Rows to process: $($CSVInputFile.Length)"
$OutcomeMarksToImport = New-Object -TypeName "System.Collections.ArrayList"
$OutcomeMarksNeedingOutcomes = @{}
$OutcomeNotFound = @{}

$OMProcessCounter = 0
foreach ($InputRow in $CSVInputFile)
{
    $NewOutcomeMarks = Convert-ToSLB -InputRow $InputRow -AllOutcomes $SLCourseObjectives -AllReportPeriods $ClassReportPeriods

    foreach($NewOutcomeMark in $NewOutcomeMarks) {
        if ($NewOutcomeMark.iCourseObjectiveID -eq -1) {

            # When we collect the ones that didn't match, we need to only collect 1 copy of the input row
            # but we're getting up to 4 marks from the Convert-ToSLB function. So, fingerprint the row, and
            # only store one of them.
            $fingerprint = Get-Hash "$($InputRow.StudentGUID)-$($InputRow.SchoolID)-$($InputRow.CourseCode)-$($InputRow.SectionGUID)-$($InputRow.ReportingTermNumber)"
            if ($OutcomeMarksNeedingOutcomes.ContainsKey($fingerprint) -eq $false) {
                $OutcomeMarksNeedingOutcomes.Add($fingerprint, $InputRow)
            }

            if ($OutcomeNotFound.ContainsKey($NewOutcomeMark.OutcomeCode) -eq $false) {
                $OutcomeText = Get-SLBText -CourseID $($NewOutcomeMark.iCourseID) -OutcomeName $($NewOutcomeMark.OutcomeName) -AllCourseGrades $CourseGrades
                $OutcomeNotFound.Add($NewOutcomeMark.OutcomeCode, [PSCustomObject]@{
                    iCourseID = [int]$($NewOutcomeMark.iCourseID)
                    OutcomeCode = $($NewOutcomeMark.OutcomeCode)
                    OutcomeText = $OutcomeText
                    cSubject = "$($NewOutcomeMark.OutcomeName): $OutcomeText"
                    mNotes = "$($NewOutcomeMark.OutcomeName): $OutcomeText"
                    iLV_ObjectiveCategoryID = 4150 # This number is used to distinguish normal outcomes from SLB outcomes
                })
            }
        } else {
            $OutcomeMarksToImport += $NewOutcomeMark
        }
    }

    $OMProcessCounter++
    $PercentComplete = [int]([decimal]($OMProcessCounter/$CSVInputFile.Length) * 100)
    if ($PercentComplete % 5 -eq 0) {
        Write-Progress -Activity "Processing input file" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
    }
}

Write-Log "Found $($OutcomeMarksToImport.Length) marks to import"
if ($OutcomeMarksNeedingOutcomes.Keys.Count -gt 0) {
    Write-Log "Found $($OutcomeMarksNeedingOutcomes.Keys.Count) without matching outcomes in SchoolLogic"
}
if ($OutcomeNotFound.Count -gt 0) {
    Write-Log "Found $($OutcomeNotFound.Count) outcomes that don't exist in our database."
}

$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $DBConnectionString

if (($ImportUnknownOutcomes -eq $true) -and ($OutcomeNotFound.Count -gt 0)) {

    $OutcomeMarksNeedingOutcomes_Two = @{}
    $OutcomeNotFound_Two = @{}

    if ($DryRun -ne $true) {
        # Insert new outcomes that didn't exist in SL before
        $OInsertCounter = 0
        foreach ($NewOutcome in $OutcomeNotFound.Values) {
            $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
            $SqlCommand.CommandText = "INSERT INTO CourseObjective(lImportedFromEdsby,OutcomeCode,OutcomeText,iCourseID,cSubject,mNotes,iLV_ObjectiveCategoryID)
                                            VALUES(1,@OUTCOMECODE,@OUTCOMETEXT,@ICOURSEID,@CSUBJECT,@MNOTES,@CATEGORYID);"

            $SqlCommand.Parameters.AddWithValue("@OUTCOMECODE",$NewOutcome.OutcomeCode) | Out-Null
            $SqlCommand.Parameters.AddWithValue("@OUTCOMETEXT",$NewOutcome.OutcomeText) | Out-Null
            $SqlCommand.Parameters.AddWithValue("@ICOURSEID",$NewOutcome.iCourseID) | Out-Null
            $SqlCommand.Parameters.AddWithValue("@CSUBJECT",$NewOutcome.cSubject) | Out-Null
            $SqlCommand.Parameters.AddWithValue("@MNOTES",$NewOutcome.mNotes) | Out-Null
            $SqlCommand.Parameters.AddWithValue("@CATEGORYID",$NewOutcome.iLV_ObjectiveCategoryID) | Out-Null
            $SqlCommand.Connection = $SqlConnection

            $SqlConnection.open()
            if ($DryRun -ne $true) {
                $Sqlcommand.ExecuteNonQuery() | Out-Null
            }
            $SqlConnection.close()

            $OInsertCounter++
            $PercentComplete = [int]([decimal](($OInsertCounter)/[decimal]($OutcomeNotFound.Values.Count)) * 100)
            if ($PercentComplete % 5 -eq 0) {
                Write-Progress -Activity "Inserting outcomes" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
            }
        }
    } else {
        Write-Log "Skipping database write due to -DryRun"
    }

    # Re-import outcomes from SchoolLogic
    Write-Log "Reloading outcomes from SchoolLogic..."
    $SLCourseObjectives = Get-CourseObjectives -DBConnectionString $DBConnectionString
    Write-Log " Processed $($SLCourseObjectives.Length) course objectives."

    # Reprocess marks that didn't have matching outcomes before
    foreach ($InputRow in $OutcomeMarksNeedingOutcomes.Values)
    {
        $NewOutcomeMarks = Convert-ToSLB -InputRow $InputRow -AllOutcomes $SLCourseObjectives -AllReportPeriods $ClassReportPeriods

        foreach($NewOutcomeMark in $NewOutcomeMarks) {
            if ($NewOutcomeMark.iCourseObjectiveId -eq -1) {

                $fingerprint = Get-Hash "$($InputRow.StudentGUID)-$($InputRow.SchoolID)-$($InputRow.CourseCode)-$($InputRow.SectionGUID)-$($InputRow.ReportingTermNumber)"
                if ($OutcomeMarksNeedingOutcomes_Two.ContainsKey($fingerprint) -eq $false) {
                    $OutcomeMarksNeedingOutcomes_Two.Add($fingerprint, $InputRow)
                }

                if ($OutcomeNotFound_Two.ContainsKey($NewOutcomeMark.OutcomeCode) -eq $false) {
                    $OutcomeNotFound_Two.Add($NewOutcomeMark.OutcomeCode,[PSCustomObject]@{
                        iCourseID = [int]$($NewOutcomeMark.iCourseID)
                        OutcomeCode = $($NewOutcomeMark.OutcomeCode)
                        OutcomeText = $OutcomeText
                        cSubject = "$($NewOutcomeMark.OutcomeName): $OutcomeText"
                        mNotes = "$($NewOutcomeMark.OutcomeName): $OutcomeText"
                        iLV_ObjectiveCategoryID = 4150 # This number is used to distinguish normal outcomes from SLB outcomes
                    })
                }
            } else {
                $OutcomeMarksToImport += ($NewOutcomeMark)
            }
        }
    }

    Write-Log "Now $($OutcomeMarksToImport.Length) marks to import"
    Write-Log "Now $([int]$($OutcomeMarksNeedingOutcomes_Two.Keys.Count)) without matching outcomes in SchoolLogic"
    Write-Log "Now $($($OutcomeNotFound_Two.Count)) outcomes that don't exist in our database."

} else {
    if ($($OutcomeMarksNeedingOutcomes.Values.Count) -gt 0) {
        Write-Log "Skipping $($OutcomeMarksNeedingOutcomes.Values.Count) marks due to missing outcomes."
    }
}

###########################################################################
# Import the outcome marks                                                #
###########################################################################

Write-Log "Inserting outcome marks into SchoolLogic..."
if ($DryRun -ne $true) {
    $OMInsertCounter = 0
    foreach($M in $OutcomeMarksToImport) {
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
        $SqlCommand.CommandText = " UPDATE StudentCourseObjective
                                    SET
                                        nMark=@NMARK,
                                        cMark=@CMARK
                                    WHERE
                                        iStudentID=@STUDENTID
                                        AND iCourseObjectiveID=@OBJECTIVEID
                                        AND iReportPeriodID=@REPID
                                        AND iCourseID=@COURSEID
                                    IF @@ROWCOUNT = 0
                                    INSERT INTO
                                        StudentCourseObjective(iStudentID, iReportPeriodID, iCourseObjectiveID, iCourseID, iSchoolID, nMark, cMark)
                                        VALUES(@STUDENTID, @REPID, @OBJECTIVEID, @COURSEID, @SCHOOLID, @NMARK, @CMARK);"

        $SqlCommand.Parameters.AddWithValue("@STUDENTID",$M.iStudentID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@REPID",$M.iReportPeriodID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@OBJECTIVEID",$M.iCourseObjectiveId) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@COURSEID",$M.iCourseID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@SCHOOLID",$M.iSchoolID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@NMARK",$M.nMark) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@CMARK",$M.cMark) | Out-Null
        $SqlCommand.Connection = $SqlConnection

        $SqlConnection.open()
        if ($DryRun -ne $true) {
            $Sqlcommand.ExecuteNonQuery() | Out-Null
        }
        $SqlConnection.close()

        $OMInsertCounter++
        $PercentComplete = [int]([decimal]($OMInsertCounter/$OutcomeMarksToImport.Count) * 100)
        if ($PercentComplete % 5 -eq 0) {
            Write-Progress -Activity "Inserting outcome marks" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
        }
    }
} else {
    Write-Log "Skipping database write due to -DryRun"
}

Write-Log "Done!"
remove-module EdsbyImportModule