param (
    [Parameter(Mandatory=$true)][string]$InputFileName,
    [string]$ConfigFilePath,
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
    "SchoolID",
    "IncidentComment",
    "MeetingDate",
    "IncidentTags",
    "IncidentID",
    "IncidentUpdateDateTime",
    "MeetingID",
    "StudentGUID",
    "MeetingTeacherGUIDs",
    "SectionGUID",
    "SectionName",
    "MeetingPeriodIDs",
    "IncidentCode",
    "IncidentReasonCode"
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

# Convert to hashtables for easier consumption
Write-Log " Loading homeroom blocks..."
$HomeroomBlocks = Get-AllHomeroomBlocks -DBConnectionString $DBConnectionString
Write-Log " Loading period blocks..."
$PeriodBlocks = Get-AllPeriodBlocks -DBConnectionString $DBConnectionString

###########################################################################
# Process the file                                                        #
###########################################################################

Write-Log "Processing input file..."
Write-Log " Rows to process: $($CSVInputFile.Length)"
$AttendanceToImport = @()

$Counter_ProcessedRows = 0
foreach ($InputRow in $CSVInputFile)
{
    # We don't care about "presents" so ignore those
    # We may have to come back and have a "present" cancel out an absence or something... but for now, just ignore them
    if ($InputRow.IncidentCode.ToLower() -eq "present") {
        continue;
    }

    $NewRecord = [PSCustomObject]@{
        # Copy over the easy fields
        iSchoolID = [int]$InputRow.SchoolID
        dDate = [datetime]$InputRow.MeetingDate
        mComment = [string]$InputRow.IncidentComment.Trim()
        mTags = [string]$InputRow.IncidentTags.Trim()
        cIncidentID = [string]$InputRow.IncidentID
        dEdsbyLastUpdate = [datetime]$InputRow.IncidentUpdateDateTime
        iMeetingID = [int]$InputRow.MeetingID

        # Convert fields that we can convert using just this file
        iStudentID = Convert-StudentID -InputString $([string]$InputRow.StudentGUID)
        iStaffID = Convert-StaffID -InputString $([string]$InputRow.MeetingTeacherGUIDs)
        iClassID = Convert-SectionID -InputString $([string]$InputRow.SectionGUID) -SchoolID $([string]$InputRow.SchoolID) -ClassName $([string]$InputRow.SectionName)

        # Convert fields using data from SchoolLogic
        # iBlockNumber
        iBlockNumber = Convert-BlockID -EdsbyPeriodsID $([int]$InputRow.MeetingPeriodIDs) -ClassName $([string]$InputRow.SectionName) -PeriodBlockDataTable $PeriodBlocks -DailyBlockDataTable $HomeroomBlocks
        iAttendanceStatusID = Convert-AttendanceStatus -AttendanceCode $([string]$InputRow.IncidentCode) -AttendanceReasonCode $([string]$InputRow.IncidentReasonCode)
        iAttendanceReasonsID = Convert-AttendanceReason -InputString $([string]$InputRow.IncidentReasonCode)
    }

    if ($null -ne $NewRecord) {
        $AttendanceToImport += $NewRecord
    }
    
    $Counter_ProcessedRows++
    $PercentComplete = [int]([decimal]($Counter_ProcessedRows/$CSVInputFile.Length) * 100)
    if ($PercentComplete % 5 -eq 0) {
        Write-Progress -Activity "Processing input file" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
    }
}
Write-Log "Finished processing $($AttendanceToImport.Count) entries"

###########################################################################
# Perform SQL operations                                                  #
###########################################################################

Exit-PSHostProcess
# Set up the SQL connection
$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $DBConnectionString


Write-Log "Inserting/Updating $($AttendanceToImport.Count) records..."
$Counter_ProcessedRows = 0
if ($DryRun -ne $true) {
    foreach ($NewRecord in $AttendanceToImport) {
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
        $SqlCommand.CommandText = "UPDATE Attendance
                                    SET 
                                        iAttendanceStatusID=@STATUSID,
                                        iAttendanceReasonsID=@REASONID
                                    WHERE
                                        iStudentID=@STUDENTID
                                        AND iClassID=@CLASSID
                                        AND iBlockNumber=@BLOCKNUM
                                        AND dDate=@DDATE
                                        AND iSchoolID=@ISCHOOLID
                                    IF @@ROWCOUNT = 0
                                    INSERT INTO Attendance(iBlockNumber, iStudentID, iAttendanceStatusID, iAttendanceReasonsID, dDate, iClassID, iMinutes, mComment, iStaffID, iSchoolID, cEdsbyIncidentID, mEdsbyTags, dEdsbyLastUpdated,iMeetingID)
                                        VALUES(@BLOCKNUM,@STUDENTID,@STATUSID,@REASONID,@DDATE,@CLASSID,@MINUTES,@MCOMMENT,@ISTAFFID,@ISCHOOLID,@EDSBYINCIDENTID,@EDSBYTAGS,@EDSBYLASTUPDATED,@MEETINGID);"
        $SqlCommand.Parameters.AddWithValue("@BLOCKNUM",$NewRecord.iBlockNumber) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@STUDENTID",$NewRecord.iStudentID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@STATUSID",$NewRecord.iAttendanceStatusID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@REASONID",$NewRecord.iAttendanceReasonsID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@DDATE",$NewRecord.dDate) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@CLASSID",$NewRecord.iClassID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@MINUTES",0) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@MCOMMENT",$NewRecord.mComment) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@ISTAFFID",$NewRecord.iStaffID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@ISCHOOLID",$NewRecord.iSchoolID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@EDSBYINCIDENTID",$NewRecord.cIncidentID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@EDSBYTAGS",$NewRecord.mTags) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@EDSBYLASTUPDATED",$NewRecord.dEdsbyLastUpdate) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@MEETINGID",$NewRecord.iMeetingID) | Out-Null
        $SqlCommand.Connection = $SqlConnection

        $SqlConnection.open()
        if ($DryRun -ne $true) {
            $Sqlcommand.ExecuteNonQuery() | Out-Null
        } else {
            Write-Log " (Skipping SQL query due to -DryRun)"
        }
        $SqlConnection.close()

        $Counter_ProcessedRows++
        $PercentComplete = [int]([decimal]($Counter_ProcessedRows/$AttendanceToImport.Count) * 100)
        if ($PercentComplete % 5 -eq 0) {
            Write-Progress -Activity "Inserting records..." -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
        }
    }
} else {
    Write-Log "Skipping database write due to -DryRun"
}
Write-Log "Done!"
remove-module EdsbyImportModule
