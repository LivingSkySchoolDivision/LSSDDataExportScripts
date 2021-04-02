param (
    [Parameter(Mandatory=$true)][string]$InputFileName,
    [string]$ConfigFilePath,
    [switch]$PerformDeletes,
    [switch]$DryDelete,
    [switch]$DryRun,
    [switch]$AllowSyncToEmptyTable,
    [switch]$DisableSafeties
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
$HomeroomBlocks = Get-AllHomeroomBlocks -DBConnectionString $DBConnectionString
$PeriodBlocks = Get-AllPeriodBlocks -DBConnectionString $DBConnectionString

###########################################################################
# Process the file                                                        #
###########################################################################

Write-Log "Processing input file..."

$AttendanceToImport = @{}

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

        # We'll fill these in below, because it's easier to refer to all the fields we need
        Thumbprint = ""
        ValueHash = ""
    }

    # Calculate hashes
    $NewRecord.Thumbprint = Get-Hash "$($NewRecord.iSchoolID)-$($NewRecord.iStudentID)-$($NewRecord.dDate.ToString("yyyyMMdd"))-$($NewRecord.iClassID)-$($NewRecord.iMeetingID)-$($NewRecord.iAttendanceStatusID)"
    $NewRecord.ValueHash = Get-Hash "$($NewRecord.dEdsbyLastUpdate.ToString("yyyyMMdd"))-$($NewRecord.mComment)-$($NewRecord.iAttendanceReasonsID)-$($NewRecord.mTags)"

    # Add to hashtable
    if ($AttendanceToImport.ContainsKey($($NewRecord.Thumbprint)) -eq $false) {
        $AttendanceToImport.Add($($NewRecord.Thumbprint), $NewRecord)
    } else {
        # A thumbprint already exists, so now we need to determine which we keep
        # Is the value hash the same? If so, discard
        # If the value hash is different, use the one updated most recently

        $ExistingRecord = $AttendanceToImport[$($NewRecord.Thumbprint)]

        if ($ExistingRecord.ValueHash -eq $NewRecord.ValueHash) {
            # Do nothing - the values are the same, so just ignore this duplicate.
            Write-Log "Ignoring duplicate record for $($NewRecord.Thumbprint) - Values identical"
        } else {
            # Trust the one updated the most recently
            if ($NewRecord.dEdsbyLastUpdate -gt $ExistingRecord.dEdsbyLastUpdate) {
                # Replace the existing record with a newer one
                $AttendanceToImport[$($NewRecord.Thumbprint)] = $NewRecord
                Write-Log "Overriding older record with newer one for $($NewRecord.Thumbprint)"
            } else {
                # Do nothing - Ignore the older record
                Write-Log "Ignoring duplicate record for $($NewRecord.Thumbprint) - Record is older"
            }
        }
    }
}

###########################################################################
# Compare to database                                                     #
###########################################################################

# Get a list of thumbprints and value hashes from the database, to compare

# Which entries do we need to add?
# Which entries do we need to update?
# Which entries do we need to remove entirely?

Write-Log "Caching existing attendance to compare to..."

$SQLQuery_ExistingAttendance = "SELECT cThumbprint, cValueHash FROM Attendance WHERE lEdsbySyncDoNotTouch=0 ORDER BY cThumbprint;"  
$ExistingAttendance_Raw = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_ExistingAttendance

$ExistingAttendance = @{}
foreach($ExistingAttendanceRow in $ExistingAttendance_Raw)
{
    if ($null -ne $ExistingAttendanceRow.cThumbprint) {
        #$ExistingAttendance += @{ $ExistingAttendanceRow.cThumbprint = $ExistingAttendanceRow.cValueHash }

        if ($ExistingAttendance.ContainsKey($ExistingAttendanceRow.cThumbprint) -eq $false) {
            $ExistingAttendance.Add($ExistingAttendanceRow.cThumbprint, $ExistingAttendanceRow.cValueHash)
        } else {
            $ExistingAttendance[$ExistingAttendanceRow.cThumbprint] = $ExistingAttendanceRow.cValueHash
        }
    }
}

# Fail if we were unable to get any existing attendance
if ($AllowSyncToEmptyTable -ne $true) {
    if ($ExistingAttendance.Count -lt 1) {
        Write-Log "No existing attendance found. Stopping for safety. To disable this safety check, use -AllowSyncToEmptyTable."
        exit
    }
}

# Now go through attendance we're importing, and see what bucket it should be in
$RecordsToInsert = @()
$RecordsToUpdate = @()
$ThumbprintsToDelete = @()

Write-Log "Finding records to insert and update..."
foreach($ImportedRecord in $AttendanceToImport.Values)
{
    # Does this thumbprint exist in our table?
    # If not, insert it
    # If yes, check it's value hash - does it match?
    #  If not, update it
    #  If yes, no work needs to be done


    if ($ExistingAttendance.ContainsKey($($ImportedRecord.Thumbprint)))
    {
        if ($ExistingAttendance[$ImportedRecord.Thumbprint] -ne $($ImportedRecord.ValueHash))
        {
            # Flag this one for an update
            $RecordsToUpdate += $ImportedRecord
        } else {
            # Do nothing - value hashes match, ignore this one
        }
    } else {
        # Thumbprint doesn't exist - add new
        $RecordsToInsert += $ImportedRecord
    }
}

if (($PerformDeletes -eq $true) -or ($DryDelete -eq $true)) {
    Write-Log "Finding records to delete..."
    # Find attendance records that have been deleted
    foreach($ExistingThumbprint in $ExistingAttendance.Keys)
    {
        # Does this thumbprint exist in the data we're importing?
        #  If yes, do nothing
        #  If no, flag for removal

        if ($AttendanceToImport.ContainsKey($ExistingThumbprint) -eq $false) {
            $ThumbprintsToDelete += $ExistingThumbprint
        }
    }
}
Write-Log "To insert: $($RecordsToInsert.Count)"
Write-Log "To update: $($RecordsToUpdate.Count)"
Write-Log "To delete: $($ThumbprintsToDelete.Count)"

###########################################################################
# Perform some safety checks                                              #
###########################################################################

if ($DisableSafeties -ne $true) {

    # Stop if the script would delete more than 10% of the existing database entries
    if ($ThumbprintsToDelete.Count -gt ($($ExistingAttendance.Count) * 0.1)) {
        Write-Log "WARNING: Script would delete more than 10% of existing database entries. Stopping script for safety. To disable this safety, use -DisableSafeties"
        exit
    }
}
###########################################################################
# Perform SQL operations                                                  #
###########################################################################

# Set up the SQL connection
$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $DBConnectionString


if ($PerformDeletes -eq $true) {
    # Implement a limit on the number of records that this script will willingly delete
    if ($DryRun -ne $true) {
        Write-Log "Deleting $($ThumbprintsToDelete.Count) records..."
        foreach ($ThumbToDelete in $ThumbprintsToDelete) {
            if ($ThumbToDelete.Length -gt 1) {
                Write-Log " > Deleting $ThumbToDelete"
                $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
                $SqlCommand.CommandText = "DELETE FROM Attendance WHERE cThumbprint=@THUMB;"
                $SqlCommand.Parameters.AddWithValue("@THUMB",$ThumbToDelete) | Out-Null
                $SqlCommand.Connection = $SqlConnection

                $SqlConnection.open()
                if ($DryRun -ne $true) {
                    $Sqlcommand.ExecuteNonQuery()
                }
                $SqlConnection.close()
            }
        }
    } else {
        Write-Log "Skipping database write due to -DryRun"
    }
}

Write-Log "Inserting $($RecordsToInsert.Count) records..."
if ($DryRun -ne $true) {
    foreach ($NewRecord in $RecordsToInsert) {
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
        $SqlCommand.CommandText = "INSERT INTO Attendance(iBlockNumber, iStudentID, iAttendanceStatusID, iAttendanceReasonsID, dDate, iClassID, iMinutes, mComment, iStaffID, iSchoolID, cEdsbyIncidentID, mEdsbyTags, dEdsbyLastUpdated,iMeetingID,cThumbprint,cValueHash)
                                        VALUES(@BLOCKNUM,@STUDENTID,@STATUSID,@REASONID,@DDATE,@CLASSID,@MINUTES,@MCOMMENT,@ISTAFFID,@ISCHOOLID,@EDSBYINCIDENTID,@EDSBYTAGS,@EDSBYLASTUPDATED,@MEETINGID,@THUMB,@VALHASH);"
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
        $SqlCommand.Parameters.AddWithValue("@THUMB",$NewRecord.Thumbprint) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@VALHASH",$NewRecord.ValueHash) | Out-Null
        $SqlCommand.Connection = $SqlConnection

        $SqlConnection.open()
        if ($DryRun -ne $true) {
            $Sqlcommand.ExecuteNonQuery()
        } else {
            Write-Log " (Skipping SQL query due to -DryRun)"
        }
        $SqlConnection.close()
    }
} else {
    Write-Log "Skipping database write due to -DryRun"
}
Write-Log "Updating $($RecordsToUpdate.Count) records..."
if ($DryRun -ne $true) {
    foreach ($UpdatedRecord in $RecordsToUpdate) {
        $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
        $SqlCommand.CommandText = "UPDATE Attendance SET iAttendanceReasonsID=@REASONID, mComment=@MCOMMENT, mEdsbyTags=@EDSBYTAGS, dEdsbyLastUpdated=@EDSBYLASTUPDATED, cValueHash=@VALHASH WHERE cThumbprint=@THUMB"
        $SqlCommand.Parameters.AddWithValue("@REASONID",$UpdatedRecord.iAttendanceReasonsID) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@MCOMMENT",$UpdatedRecord.mComment) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@EDSBYTAGS",$UpdatedRecord.mTags) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@EDSBYLASTUPDATED",$UpdatedRecord.dEdsbyLastUpdate) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@VALHASH",$UpdatedRecord.ValueHash) | Out-Null
        $SqlCommand.Parameters.AddWithValue("@THUMB",$UpdatedRecord.Thumbprint) | Out-Null
        $SqlCommand.Connection = $SqlConnection

        $SqlConnection.open()
        if ($DryRun -ne $true) {
            $Sqlcommand.ExecuteNonQuery()
        } else {
            Write-Log " (Skipping SQL query due to -DryRun)"
        }
        $SqlConnection.close()
    }
} else {
    Write-Log "Skipping database write due to -DryRun"
}

Write-Log "Done!"
remove-module EdsbyImportModule

