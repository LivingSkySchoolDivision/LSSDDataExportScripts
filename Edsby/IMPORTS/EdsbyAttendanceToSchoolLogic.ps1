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
    if ($line.Contains('"IncidentComment"') -eq $false) { throw "Input CSV missing field: IncidentComment" }
    if ($line.Contains('"MeetingDate"') -eq $false) { throw "Input CSV missing field: MeetingDate" }
    if ($line.Contains('"IncidentTags"') -eq $false) { throw "Input CSV missing field: IncidentTags" }
    if ($line.Contains('"IncidentID"') -eq $false) { throw "Input CSV missing field: IncidentID" }
    if ($line.Contains('"IncidentUpdateDateTime"') -eq $false) { throw "Input CSV missing field: IncidentUpdateDateTime" }
    if ($line.Contains('"MeetingID"') -eq $false) { throw "Input CSV missing field: MeetingID" }
    if ($line.Contains('"StudentGUID"') -eq $false) { throw "Input CSV missing field: StudentGUID" }
    if ($line.Contains('"MeetingTeacherGUIDs"') -eq $false) { throw "Input CSV missing field: MeetingTeacherGUIDs" }
    if ($line.Contains('"SectionGUID"') -eq $false) { throw "Input CSV missing field: SectionGUID" }
    if ($line.Contains('"SectionName"') -eq $false) { throw "Input CSV missing field: SectionName" }
    if ($line.Contains('"MeetingPeriodIDs"') -eq $false) { throw "Input CSV missing field: MeetingPeriodIDs" }
    if ($line.Contains('"IncidentCode"') -eq $false) { throw "Input CSV missing field: IncidentCode" }
    if ($line.Contains('"IncidentReasonCode"') -eq $false) { throw "Input CSV missing field: IncidentReasonCode" }
    
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

function Convert-AttendanceReason {
    param(
        [Parameter(Mandatory=$true)] $InputString
    )

    # We'll return -1 as something we should ignore, and then ignore those rows later in the program

    # Could pull from the database, but im in a crunch, so this is getting manually correlated for now.
    # Attendance Reasons in SchoolLogic:
    # 98    Known Reason
    # 100   Medical
    # 101   Extra-Curr
    # 103   Curricular
    # 104   Engaged

    # Reason codes from Edsby
    #   S-Curr
    #   S-XCurr
    #   S-Med
    #   S-Exp
    #   S-Eng
    #   A-Med
    #   A-Exp
    #   A-UnExp
    #   LA-Curr
    #   LA-XCurr
    #   LA-Med
    #   LA-Exp
    #   LA-UnExp
    #   LE-Curr
    #   LE-XCurr
    #   LE-Med
    #   LE-Exp
    #   LE-UnExp

    if ($InputString -like '*-Exp') { return 98 }
    if ($InputString -like '*-Med') { return 100 }
    if ($InputString -like '*-Curr') { return 103 }
    if ($InputString -like '*-XCurr') { return 101 }
    if ($InputString -like '*-Eng') { return 104 }

    return 0
}

function Convert-AttendanceStatus {
    param(
        [Parameter(Mandatory=$true)] $AttendanceCode,
        [Parameter(Mandatory=$true)] $AttendanceReasonCode
    )

    # We'll return -1 as something we should ignore, and then ignore those rows later in the program

    # Could pull from the database, but im in a crunch, so this is getting manually correlated for now.
    # Attendance Statuses in SchoolLogic:
    # 1     Present
    # 2     Absent
    # 3     Late
    # 4     School (Absent from class, but still at a school fuction)
    # 5     No Change
    # 6     Leave Early
    # 7     Division (School closures)

    if ($AttendanceCode -like 'absent*') { return 2 }  # Unexplained absence
    if ($AttendanceCode -like 'sanctioned*') { return 4 }
    if ($AttendanceCode -like 'late*') { return 3 }

    # Excused might mean absent or leave early
    if ($AttendanceCode -like 'excused*') {
        if ($null -ne $AttendanceReasonCode) {
            if ($AttendanceReasonCode -like 'le-*') {
                return 6
            }
            if ($AttendanceReasonCode -like 'la-*') {
                return 3
            }
            if ($AttendanceReasonCode -like 's-*') {
                return 4
            }
            return 2
        }
    }


    return -1
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
        [Parameter(Mandatory=$true)] $SchoolID,
        [Parameter(Mandatory=$true)] $ClassName
    )

    if ($ClassName -like 'homeroom*') {
        return 0
    } else {
        return $InputString.Replace("$SchoolID-","")
    }
}

function Convert-StudentID {
    param(
        [Parameter(Mandatory=$true)] $InputString
    )

    return [int]$InputString.Replace("STUDENT-","")
}

function Convert-StaffID {
    param(
        [Parameter(Mandatory=$true)] $InputString
    )

    # We might get a list of staff, in which case we should parse it and just return the first one
    $StaffList = $InputString.Split(',')

    return [int]$StaffList[0].Replace("STAFF-","")
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


function Convert-BlockID {
    param(
        [Parameter(Mandatory=$true)][int] $EdsbyPeriodsID,
        [Parameter(Mandatory=$true)][string] $ClassName,
        [Parameter(Mandatory=$true)] $PeriodBlockDataTable,
        [Parameter(Mandatory=$true)] $DailyBlockDataTable
    )

    # Determine if this is a homeroom or a period class
    # The only way we can do that with the data we have, is that homeroom class names
    # always start with the word "Homeroom".
    # Homerooms use the DailyBlockDataTable, scheduled classes use the PeriodBlockDataTable

    $Block = $null

    if ($ClassName -like 'homeroom*') {
        $Block = $DailyBlockDataTable.Where({ $_.ID -eq $EdsbyPeriodsID })
    } else {
        $Block = $PeriodBlockDataTable.Where({ $_.ID -eq $EdsbyPeriodsID })
    }

    if ($null -ne $Block) {
        return [int]$($Block.iBlockNumber)
    }

    return -1
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
# Collect required info from the SL database                              #
###########################################################################

Write-Log "Loading required data from SchoolLogic DB..."

$SQLQuery_HomeroomBlocks = "SELECT iAttendanceBlocksID as ID, iBlockNumber, cName FROM AttendanceBlocks;"
$SQLQuery_PeriodBlocks = "SELECT iBlocksID as ID, iBlockNumber, cName FROM Blocks"

# Convert to hashtables for easier consumption
$HomeroomBlocks = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_HomeroomBlocks
$PeriodBlocks = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_PeriodBlocks

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
            } else {
                Write-Log " (Skipping SQL query due to -DryRun)"
            }
            $SqlConnection.close()
        }
    }
}

Write-Log "Inserting $($RecordsToInsert.Count) records..."
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

Write-Log "Updating $($RecordsToUpdate.Count) records..."
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

