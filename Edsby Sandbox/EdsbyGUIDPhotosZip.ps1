param (
    [Parameter(Mandatory=$true)][string]$OutputFileName,
    [string]$ConfigFilePath,
    [int]$BatchSize = 50
 )

##############################################
# Script configuration                       #
##############################################

# Edsby allows us to sync roughly a dozen schools to the Sandbox
# The ActiveSchools.csv can be used to adjust the schools we want to send
$ActiveSchools = Import-Csv -Path .\ActiveSchools.csv
$ActiveSchools = $ActiveSchools | ? {$_.Sync -eq 'T'} 
$iSchoolIDs = Foreach ($ID in $ActiveSchools) {
    if ($ID -ne $ActiveSchools[-1]) { $ID.iSchoolID + ',' } else { $ID.iSchoolID }    
}

# SQL Query to get the number of records we'll need

$SqlQuery_Count = "SELECT
                        Student.cStudentNumber
                    FROM 
                        StudentPhoto
                        LEFT OUTER JOIN StudentStatus ON StudentPhoto.iStudentID=StudentStatus.iStudentID
                        LEFT OUTER JOIN Student ON StudentPhoto.iStudentId=Student.iStudentId
                    WHERE 
                        (StudentStatus.dInDate <=  { fn CURDATE() }) 
                        AND ((StudentStatus.dOutDate < '1901-01-01') OR (StudentStatus.dOutDate >=  { fn CURDATE() }))  
                        AND (StudentStatus.lOutsideStatus = 0) 
                        AND Student.iSchoolID IN ($iSchoolIDs)
                        "

# SQL Query to get image data
# The output CSV file will use column names from your SQL query.
# Rename them using "as" - example: "SELECT cFirstName as FirstName FROM Students"

$SqlQuery_Photos = "SELECT
                            CONCAT('STUDENT-',Student.iStudentID), 
                            StudentPhoto.bImage, 
                            StudentPhoto.cImageType 
                        FROM 
                            StudentPhoto
                            LEFT OUTER JOIN StudentStatus ON StudentPhoto.iStudentID=StudentStatus.iStudentID
                            LEFT OUTER JOIN Student ON StudentPhoto.iStudentId=Student.iStudentId
                        WHERE 
                            (StudentStatus.dInDate <=  { fn CURDATE() }) 
                            AND ((StudentStatus.dOutDate < '1901-01-01') OR (StudentStatus.dOutDate >=  { fn CURDATE() }))  
                            AND (StudentStatus.lOutsideStatus = 0)
                            AND Student.iSchoolID IN ($iSchoolIDs)
                        ORDER BY StudentPhoto.iStudentId 
                        "

##############################################
# No configurable settings beyond this point #
##############################################

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
$ConnectionString = $configXML.Settings.SchoolLogic.ConnectionString

# Set up the SQL connection
$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $ConnectionString
$SqlCommand = New-Object System.Data.SqlClient.SqlCommand
$SqlCommand.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCommand

# Create a temporary scratch folder to store all the images in
$ScratchFolderPath = "$OutputFileName.tmp"
write-host "PATH IS $ScratchFolderPath"

if ((Test-Path $OutputFileName) -eq $true) {
    Remove-Item $OutputFileName -Recurse -Force
}

if ((Test-Path $ScratchFolderPath) -eq $true) {
    Remove-Item $ScratchFolderPath -Recurse -Force
}

New-Item -Path $ScratchFolderPath -ItemType Directory

# Get the path for the output file
# You wouldn't think you'd have to do this
# But you're using powershell
$OutputLocation = (get-item $ScratchFolderPath)

# Get a count of how many photos there are
$SqlCommand.CommandText = $SqlQuery_Count
$SqlConnection.open()
$countDataSet = New-Object System.Data.DataSet
$Count = $SqlAdapter.Fill($countDataSet)
$SqlConnection.close()

# Run the SQL query for the photos in batches
$batchNumber = 0
for ($x=0;$x -le $Count;$x+=$BatchSize) {
    $Offset = $batchNumber * $BatchSize
    $BatchSQL = "$SqlQuery_Photos OFFSET $Offset ROWS FETCH NEXT $BatchSize ROWS ONLY"

    write-host "$Offset /  $Count"

    # Get this batch from SQL
    
    $SqlCommand.CommandText = $BatchSQL
    $SqlConnection.open()
    $PhotoDataSet = New-Object System.Data.DataSet
    $throwaway123 = $SqlAdapter.Fill($PhotoDataSet)
    $SqlConnection.close()
    
    # Write these to file
    
    foreach($DSTable in $PhotoDataSet.Tables) {
        foreach($DataRow in $DSTable){
            # WriteAllBytes wants to write in the directory above where it should for some reason
            # I have no idea why.
            [byte[]]$photoBytes = $DataRow[1]
            $StudentId = $DataRow[0]
            $FilePath = join-path $OutputLocation "$FullFileName$StudentId.jpg"
            [IO.File]::WriteAllBytes("$FilePath", $photoBytes)
        }
    }

    $batchNumber++
}

# Zip the whole scratch folder up
Compress-Archive -Path $ScratchFolderPath/*.* -DestinationPath $OutputFileName

# Delete the scratch folder, leaving only the zip file
Remove-Item $ScratchFolderPath -Recurse -Force