param (
    [Parameter(Mandatory=$true)][string]$InputFileName,
    [string]$ConfigFilePath
 )

##############################################
# How this script works                      #
##############################################

# Import
# Processing
# Referencing SchoolLogic data
# Find changes between our list and the absences table in SchoolLogic
# Insert new attendance entries
# Delete old attendance entries

##############################################
# Data we need to convert                    #
##############################################

# iBlockNumber
# iEnrollmentID (maybe)
# iClassID, iStudentID, iStaffID (easily parsed from the input file)
# Attendance statuses and reasons

##############################################
# Data we need from SchoolLogic              #
##############################################

# iAttendanceBlockIds, for period attendance block numbers
# iBlockIDs for daily attendance block numbers

##############################################
# Functions                                  #
##############################################
function Get-CSV {
    param(
        [Parameter(Mandatory=$true)][String] $CSVFile
    )

    return import-csv $CSVFile -header('SchoolID','IncidentID','IncidentDate','UpdateDate','StudentFirstName','StudentLastName','StudentGUID','StudentID','StudentMinistryID','StudentGrade','PeriodIDs','MeetingID','MeetingStartTime','MeetingEndTime','Class','ClassGUID','TeacherNames','TeacherGUIDs','Room','Code','ReasonCode','Reason','Comment','Tags') | Select -skip 1
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

##############################################
# Script initialization                      #
##############################################

Write-Output "Loading config file..."

# Find the config file
$AdjustedConfigFilePath = $ConfigFilePath
if ($AdjustedConfigFilePath.Length -le 0) {
    $AdjustedConfigFilePath = join-path -Path $(Split-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) -Parent) -ChildPath "config.xml"
}

# Retreive the connection string from config.xml
if ((test-path -Path $AdjustedConfigFilePath) -eq $false) {
    Throw "Config file not found. Specify using -ConfigFilePath. Defaults to config.xml in the directory above where this script is run from."
}
$configXML = [xml](Get-Content $AdjustedConfigFilePath)
$DBConnectionString = $configXML.Settings.SchoolLogic.ConnectionStringRW

# Check if the import file exists before going any further
if (Test-Path $InputFileName) 
{    
} else {
    write-output "Couldn't load the input file! Quitting."
    exit
}

##############################################
# Collect required info from the SL database #
##############################################

Write-Output "Loading required data from SchoolLogic DB..."

$SQLQuery_HomeroomBlocks = "SELECT iAttendanceBlocksID as ID, iBlockNumber, cName FROM AttendanceBlocks;"
$SQLQuery_PeriodBlocks = "SELECT iBlocksID as ID, iBlockNumber, cName FROM Blocks"

# Convert to hashtables for easier consumption
$HomeroomBlocks = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_HomeroomBlocks
$PeriodBlocks = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_PeriodBlocks

##############################################
# Process the file                           #
##############################################

Write-Output "Processing input file..."

$ConvertedRows = @()

foreach ($InputRow in Get-CSV -CSVFile $InputFileName)
{
    $ConvertedObj = @{}

    # Copy over the easy fields    
    $ConvertedObj.iSchoolID = [int]$InputRow.SchoolID
    $ConvertedObj.dDate = [datetime]$InputRow.IncidentDate    
    $ConvertedObj.mComment = [string]$InputRow.Comment.Trim()
    $ConvertedObj.mTags = [string]$InputRow.Tags.Trim()
    $ConvertedObj.cIncidentID = [string]$InputRow.IncidentID
    $ConvertedObj.dEdsbyLastUpdate = [datetime]$InputRow.UpdateDate
    $ConvertedObj.iMeetingID = [int]$InputRow.MeetingID
    
    # Convert fields that we can convert using just this file
    $ConvertedObj.iStudentID = Convert-StudentID -InputString $([string]$InputRow.StudentGUID)
    $ConvertedObj.iStaffID = Convert-StaffID -InputString $([string]$InputRow.TeacherGUIDs)
    $ConvertedObj.iClassID = Convert-SectionID -InputString $([string]$InputRow.ClassGUID) -SchoolID $([string]$InputRow.SchoolID) -ClassName $([string]$InputRow.Class)

    # Convert fields using data from SchoolLogic
    # iBlockNumber
    $ConvertedObj.iBlockNumber = Convert-BlockID -EdsbyPeriodsID $([int]$InputRow.PeriodIDs) -ClassName $([string]$InputRow.Class) -PeriodBlockDataTable $PeriodBlocks -DailyBlockDataTable $HomeroomBlocks
    $ConvertedObj.iAttendanceStatusID = Convert-AttendanceStatus -AttendanceCode $([string]$InputRow.Code) -AttendanceReasonCode $([string]$InputRow.ReasonCode)
    $ConvertedObj.iAttendanceReasonsID = Convert-AttendanceReason -InputString $([string]$InputRow.ReasonCode)

    $ConvertedObj.Thumbprint = Get-Hash "$($ConvertedObj.iSchoolID)-$($ConvertedObj.iStudentID)-$($ConvertedObj.dDate.ToString("yyyyMMdd"))-$($ConvertedObj.iMeetingID)"
    $ConvertedObj.ValueHash = Get-Hash "$($ConvertedObj.dEdsbyLastUpdate.ToString("yyyyMMdd"))-$($ConvertedObj.mComment)-$($ConvertedObj.iAttendanceStatusID)-$($ConvertedObj.iAttendanceReasonsID)"

    # Ignore rows that converted badly
    if ($ConvertedObj.iAttendanceStatusID -ne -1) {
        $ConvertedRows += $ConvertedObj
    }
}

#foreach ($NewRecord in $ConvertedRows) {
#    write-host "$($NewRecord.Thumbprint)"
#}

#$ConvertedRows | Foreach-Object {[PSCustomObject]$_}  | Format-Table

##############################################
# Compare to database                        #
##############################################

# Lets assume that we can identify an absence in the db by:
#  Date
#  Student ID number
#  Class or homeroom ID
#  Edsby incident ID (can we trust this?)

# Which entries do we need to add?
# Which entries do we need to update?
# Which entries do we need to remove entirely?


##############################################
# Import into SQL                            #
##############################################

Write-Output "Inserting into database..."

# Set up the SQL connection
$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $DBConnectionString

foreach ($NewRecord in $ConvertedRows) {
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
    $Sqlcommand.ExecuteNonQuery() | Out-File -Append log.log
    $SqlConnection.close()
}

