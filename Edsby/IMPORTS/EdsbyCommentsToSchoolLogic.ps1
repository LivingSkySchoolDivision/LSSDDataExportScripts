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
    if ($line.Contains('"ReportingTermNumber"') -eq $false) { throw "Input CSV missing field: ReportingTermNumber" }  
    if ($line.Contains('"StudentGUID"') -eq $false) { throw "Input CSV missing field: StudentGUID" }
    if ($line.Contains('"SchoolID"') -eq $false) { throw "Input CSV missing field: SchoolID" }
    if ($line.Contains('"Comment"') -eq $false) { throw "Input CSV missing field: Comment" }
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

function Convert-ToComment {
    param(
        [Parameter(Mandatory=$true)] $InputRow,
        [Parameter(Mandatory=$true)] $AllReportPeriods
    )

    $iClassID = (Convert-SectionID -SchoolID $InputRow.SchoolID -InputString $InputRow.SectionGUID)
    $Number = [int]($InputRow.ReportingTermNumber)
    $iReportPeriodID = [int]((Get-ReportPeriodID -iClassID $iClassID -AllClassReportPeriods $AllReportPeriods -Number $Number))

    $NewMark = [PSCustomObject]@{
        iReportPeriodID = [int]$iReportPeriodID
        iStudentID = [int](Convert-StudentID $InputRow.StudentGUID)
        iClassID = $iClassID
        iSchoolID = $InputRow.SchoolID
        mComment = $InputRow.Comment
    }

    if ($NewMark.iReportPeriodID -eq -1) {
        Write-Log "Invalid classid and report period number combination: $($iClassID) / $($InputRow.ReportingTermNumber)"
    }

    return $NewMark
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

function Convert-ClassReportPeriodsToHashtable {
    param(
        [Parameter(Mandatory=$true)] $AllClassReportPeriods
    )

    $Output = @{}

    foreach($RP in $AllClassReportPeriods) {
        if ($null -ne $RP) {
            if ($RP.iClassID -gt 0) {
                if ($Output.ContainsKey($RP.iClassID) -eq $false) {
                    $OutPut.Add($RP.iClassID, @())
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

Write-Log "Loading and processing class report periods..."
$ClassReportPeriods = Convert-ClassReportPeriodsToHashtable -AllClassReportPeriods $(Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_ClassReportPeriods)
Write-Log " Loaded report periods for $($ClassReportPeriods.Keys.Count) classes."

###########################################################################
# Process the file                                                        #
###########################################################################

Write-Log "Processing input file..."
Write-Log " Rows to process: $($CSVInputFile.Length)"
$MarksToImport = @()

$OMProcessCounter = 0
foreach ($InputRow in $CSVInputFile)
{
    # If there is no grade, ignore
    if ($InputRow.OverallMark -eq "") {
        continue;
    }

    # Assemble the final mark object
    $NewComment = Convert-ToComment -InputRow $InputRow -AllReportPeriods $ClassReportPeriods

    if ($NewComment.mComment.Length -gt 0) {
        $MarksToImport += $NewComment
    }
    
    $OMProcessCounter++
    $PercentComplete = [int]([decimal]($OMProcessCounter/$CSVInputFile.Length) * 100)
    if ($PercentComplete % 5 -eq 0) {
        Write-Progress -Activity "Processing input file" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
    }
}

Write-Log "Found $($MarksToImport.Length) comments to import"

$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $DBConnectionString

###########################################################################
# Import the comments                                                     #
###########################################################################

Write-Log "Inserting commentsinto SchoolLogic..."
$OMInsertCounter = 0
foreach($M in $MarksToImport) {
    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
    $SqlCommand.CommandText = " UPDATE Marks 
                                SET 
                                    mComment=@MCOMMENT
                                WHERE 
                                    iStudentID=@STUDENTID 
                                    AND iClassID=@CLASSID 
                                    AND iReportPeriodID=@REPID 
                                    AND nMark=0 
                                    AND cMark=''
                                IF @@ROWCOUNT = 0 
                                INSERT INTO 
                                    Marks(iStudentID, iReportPeriodID, iClassID, dDateAssigned, iSchoolID, ImportTimestamp, ImportBatchID, mComment)
                                    VALUES(@STUDENTID, @REPID, @CLASSID, @DDATEASS, @SCHOOLID, @DDATEASS, @DDATEASS, @MCOMMENT);"
    
    $SqlCommand.Parameters.AddWithValue("@STUDENTID",$M.iStudentID) | Out-Null
    $SqlCommand.Parameters.AddWithValue("@REPID",$M.iReportPeriodID) | Out-Null
    $SqlCommand.Parameters.AddWithValue("@CLASSID",$M.iClassID) | Out-Null
    $SqlCommand.Parameters.AddWithValue("@SCHOOLID",$M.iSchoolID) | Out-Null
    $SqlCommand.Parameters.AddWithValue("@DDATEASS",$(Get-Date)) | Out-Null
    $SqlCommand.Parameters.AddWithValue("@MCOMMENT",$M.mComment) | Out-Null
    $SqlCommand.Connection = $SqlConnection

    $SqlConnection.open()
    if ($DryRun -ne $true) {
        $Sqlcommand.ExecuteNonQuery() | Out-Null
    } 
    $SqlConnection.close()

    $OMInsertCounter++
    $PercentComplete = [int]([decimal]($OMInsertCounter/$MarksToImport.Count) * 100)
    if ($PercentComplete % 5 -eq 0) {
        Write-Progress -Activity "Inserting comments" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete;
    }
}
Write-Log "Done!"