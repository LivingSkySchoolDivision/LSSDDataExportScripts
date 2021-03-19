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
    if ($line.Contains('"ReportingTermNumber"') -eq $false) { throw "Input CSV missing field: ReportingTermNumber" }  
    if ($line.Contains('"StudentGUID"') -eq $false) { throw "Input CSV missing field: StudentGUID" }
    if ($line.Contains('"SchoolID"') -eq $false) { throw "Input CSV missing field: SchoolID" }
    if ($line.Contains('"Citizenship"') -eq $false) { throw "Input CSV missing field: Citizenship" }
    if ($line.Contains('"Collaboration"') -eq $false) { throw "Input CSV missing field: Collaboration" }
    if ($line.Contains('"Engagement"') -eq $false) { throw "Input CSV missing field: Engagement" }
    if ($line.Contains('"Discipline"') -eq $false) { throw "Input CSV missing field: Discipline" }
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

function Convert-SectionID {
    param(
        [Parameter(Mandatory=$true)] $InputString,
        [Parameter(Mandatory=$true)] $SchoolID
    )
    return $InputString.Replace("$SchoolID-","")
}

function Convert-IndividualSLBMark {
    param(
        [Parameter(Mandatory=$true)] $InputRow,
        [Parameter(Mandatory=$true)] $OutcomeName,
        [Parameter(Mandatory=$true)] $MarkFieldName
    )
    
    # Parse cMark vs nMark
    $cMark = ""
    $nMark = [decimal]0.0

    if ([bool]($InputRow.OverallMark -as [decimal]) -eq $true) {
        $nMark = [decimal]$InputRow.OverallMark
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
        $cMark = $InputRow.OverallMark
    }

    $iClassID = (Convert-SectionID -SchoolID $InputRow.SchoolID -InputString $InputRow.SectionGUID)
    $Number = [int]($InputRow.ReportingTermNumber)
    $iReportPeriodID = [int]((Get-ReportPeriodID -iClassID $iClassID -AllClassReportPeriods $AllReportPeriods -Number $Number))

    $NewMark = [PSCustomObject]@{
        iReportPeriodID = [int]$iReportPeriodID
        iStudentID = [int](Convert-StudentID $InputRow.StudentGUID)
        iClassID = $iClassID
        iSchoolID = $InputRow.SchoolID
        cMark = $cMark
        nMark = $nMark
    }

    if ($NewMark.iReportPeriodID -eq -1) {
        Write-Log "Invalid classid and report period number combination: $($iClassID) / $($InputRow.ReportingTermNumber)"
    }

    return $NewMark

}

function Convert-ToSLB {
    param(
        [Parameter(Mandatory=$true)] $InputRow,
        [Parameter(Mandatory=$true)] $AllOutcomes,
        [Parameter(Mandatory=$true)] $AllReportPeriods
    )

    # This needs to output _4_ marks, one for each outcome. The consumer of this function will need to handle getting an array back.
    # "Citizenship","Collaboration","Engagement","Discipline"

    $Output = New-Object -TypeName "System.Collections.ArrayList"

    # Parse Citizenship mark
    $Mark_Citizenship = Convert-IndividualSLBMark -InputRow $InputRow -OutcomeName "Citizenship" -MarkFieldName "Citizenship"
    if ($null -ne $Mark_Citizenship) {
        $Output += $Mark_Citizenship
    }

    # Parse Collaboration mark
    $Mark_Collaboration = Convert-IndividualSLBMark -InputRow $InputRow -OutcomeName "Collaboration" -MarkFieldName "Collaboration"
    if ($null -ne $Mark_Collaboration) {
        $Output += $Mark_Collaboration
    }
    
    # Parse Engagement mark
    $Mark_Engagement = Convert-IndividualSLBMark -InputRow $InputRow -OutcomeName "Engagement" -MarkFieldName "Engagement"
    if ($null -ne $Mark_Engagement) {
        $Output += $Mark_Engagement
    }

    # Parse Self-Directed mark (which is erroneously named "Discipline" in the export file)
    $Mark_SelfDirected = Convert-IndividualSLBMark -InputRow $InputRow -OutcomeName "SelfDirected" -MarkFieldName "Discipline"
    if ($null -ne $Mark_SelfDirected) {
        $Output += $Mark_SelfDirected
    }

    return $Output
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
function Convert-ObjectivesToHashtable {
    param(
        [Parameter(Mandatory=$true)] $Objectives
    )
    $Output = New-Object -TypeName "System.Collections.ArrayList"

    foreach($Obj in $Objectives) {
        if (($Obj.OutcomeCode -ne "") -and ($null -ne $Obj.OutcomeCode)) {
                $Outcome = [PSCustomObject]@{
                OutcomeCode = $Obj.OutcomeCode
                OutcomeText = $Obj.OutcomeText
                iCourseObjectiveID = $Obj.iCourseObjectiveID
                iCourseID = $Obj.iCourseID
                cSubject = $Obj.cSubject
            }
            $Output += $Outcome
            
        }
    }

    return $Output

}


function Convert-ClassReportPeriodsToHashtable {
    param(
        [Parameter(Mandatory=$true)] $AllClassReportPeriods
    )

    $Output = @{}

    foreach($RP in $AllClassReportPeriods) {
        if ($null -ne $RP) {
            if ($RP.iClassID -gt 0) {
                if ($Output.ContainsKey($RP.iClassID) -eq $false) {
                    $OutPut.Add($RP.iClassID, (New-Object -TypeName "System.Collections.ArrayList"))
                }
                $NewRP = [PSCustomObject]@{
                    iClassID = $RP.iClassID; 
                    iTrackID = $RP.iTrackID;
                    iReportPeriodID = $RP.iReportPeriodID;
                    cName = $RP.cName;
                    dStartDate = $RP.dStartDate;
                    dEndDate = $RP.dEndDate;
                }                
                $Output[$RP.iClassID] += $NewRP;
            }
        }
    }

    return $Output
}

function Get-ReportPeriodID {
    param(
        [Parameter(Mandatory=$true)] [int]$iClassID,
        [Parameter(Mandatory=$true)] [int]$Number,
        [Parameter(Mandatory=$true)] $AllClassReportPeriods
    ) 

    if ($Number -gt 0) {
        if ($AllClassReportPeriods.ContainsKey($iClassID)) {  
            if ($AllClassReportPeriods[$iClassID].Length -ge ($Number)) {
                return $($AllClassReportPeriods[$iClassID][$Number-1]).iReportPeriodID
            }
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

$SQLQuery_CourseObjectives = "SELECT iCourseObjectiveID, OutcomeCode, iCourseID, cSubject FROM CourseObjective"
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
$CSVInputFile = Get-CSV -CSVFile $InputFileName

###########################################################################
# Collect required info from the SL database                              #
###########################################################################

Write-Log "Loading required data from SchoolLogic DB..."

Write-Log "Loading and processing course objectives..."
$SLCourseObjectives_Raw = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_CourseObjectives
Write-Log " Loaded $($SLCourseObjectives_Raw.Length) course objectives."
$SLCourseObjectives = Convert-ObjectivesToHashtable -Objectives $SLCourseObjectives_Raw
Write-Log " Processed $($SLCourseObjectives.Length) course objectives."
Write-Log "Loading and processing class report periods..."
$ClassReportPeriods = Convert-ClassReportPeriodsToHashtable -AllClassReportPeriods $(Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_ClassReportPeriods)
Write-Log " Loaded report periods for $($ClassReportPeriods.Keys.Count) classes."

###########################################################################
# Process the file                                                        #
###########################################################################

Write-Log "Processing input file..."
Write-Log " Rows to process: $($CSVInputFile.Length)"
$OutcomeMarksToImport = New-Object -TypeName "System.Collections.ArrayList"
$OutcomeMarksNeedingOutcomes = New-Object -TypeName "System.Collections.ArrayList"
$OutcomeNotFound = @{}

$OMProcessCounter = 0
foreach ($InputRow in $CSVInputFile)
{
    # If there is no grade, ignore
    if ($InputRow.Grade -eq "") {
        continue;
    }

    # Assemble the final mark object
    $NewOutcomeMark = Convert-ToSLOutcomeMark -InputRow $InputRow -AllOutcomes $SLCourseObjectives -AllReportPeriods $ClassReportPeriods 

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
                iLV_ObjectiveCategoryID = 4146 # This number is used to distinguish normal outcomes from SLB outcomes
            })
        }
    } else {
        $OutcomeMarksToImport += ($NewOutcomeMark)
    }

    $OMProcessCounter++
    $PercentComplete = [int]([decimal]($OMProcessCounter/$CSVInputFile.Length) * 100)
    if ($PercentComplete % 5 -eq 0) {
        Write-Progress -Activity "Processing input file" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
    }
}

Write-Log "Found $($OutcomeMarksToImport.Length) marks to import"
if ($OutcomeMarksNeedingOutcomes.Length -gt 0) {
    Write-Log "Found $($OutcomeMarksNeedingOutcomes.Length) without matching outcomes in SchoolLogic"
}
if ($OutcomeNotFound.Count -gt 0) {
    Write-Log "Found $($OutcomeNotFound.Count) outcomes that don't exist in our database."
}

$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $DBConnectionString


if (($ImportUnknownOutcomes -eq $true) -and ($OutcomeNotFound.Count -gt 0)) {
        
    $OutcomeMarksNeedingOutcomes_Two = New-Object -TypeName "System.Collections.ArrayList"
    $OutcomeNotFound_Two = @{}

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
        } else {
            Write-Log " (Skipping SQL query due to -DryRun)"
        }
        $SqlConnection.close()

        $OInsertCounter++        
        $PercentComplete = [int]([decimal](($OInsertCounter)/[decimal]($OutcomeNotFound.Values.Count)) * 100)
        if ($PercentComplete % 5 -eq 0) {
            Write-Progress -Activity "Inserting outcomes" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
        }
    }

    # Re-import outcomes from SchoolLogic
    Write-Log "Reloading outcomes from SchoolLogic..."
    $SLCourseObjectives_Raw = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_CourseObjectives
    Write-Log " Loaded $($SLCourseObjectives_Raw.Length) course objectives."
    Write-Log "Processing course objectives from SchoolLogic DB..."
    $SLCourseObjectives = Convert-ObjectivesToHashtable -Objectives $SLCourseObjectives_Raw
    Write-Log " Processed $($SLCourseObjectives.Length) course objectives."

    # Reprocess marks that didn't have matching outcomes before
    foreach ($InputRow in $OutcomeMarksNeedingOutcomes)
    {    
        # Assemble the final mark object
        $NewOutcomeMark = Convert-ToSLOutcomeMark -InputRow $InputRow -AllOutcomes $SLCourseObjectives -AllReportPeriods $ClassReportPeriods 
    
        if ($NewOutcomeMark.iCourseObjectiveId -eq -1) {
            $OutcomeMarksNeedingOutcomes_Two += $InputRow

            $Fingerprint = (Get-Hash -String ("$($InputRow.CourseCode)$($InputRow.CriterionName)"))
            if ($OutcomeNotFound_Two.ContainsKey($Fingerprint) -eq $false) {
                $OutcomeNotFound_Two.Add($Fingerprint,[PSCustomObject]@{
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

    Write-Log "Now $($OutcomeMarksToImport.Length) marks to import"
    Write-Log "Now $($OutcomeMarksNeedingOutcomes_Two.Length) without matching outcomes in SchoolLogic"
    Write-Log "Now $($($OutcomeNotFound_Two.Count)) outcomes that don't exist in our database."

} else {
    if ($($OutcomeMarksNeedingOutcomes.Length) -gt 0) {
        Write-Log "Skipping $($OutcomeMarksNeedingOutcomes.Length) marks due to missing outcomes."
    }
}

exit

###########################################################################
# Import the outcome marks                                                #
###########################################################################

Write-Log "Inserting outcome marks into SchoolLogic..."
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
    } else {
        Write-Log " (Skipping SQL query due to -DryRun)"
    }
    $SqlConnection.close()

    $OMInsertCounter++
    $PercentComplete = [int]([decimal]($OMInsertCounter/$OutcomeMarksToImport.Count) * 100)
    if ($PercentComplete % 5 -eq 0) {
        Write-Progress -Activity "Inserting outcome marks" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
    }
}
Write-Log "Done!"