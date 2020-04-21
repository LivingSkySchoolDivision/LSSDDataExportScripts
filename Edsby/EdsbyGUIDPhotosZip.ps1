param (
    [Parameter(Mandatory=$true)][string]$OutputFileName,
    [string]$ConfigFilePath
 )

##############################################
# Script configuration                       #
##############################################

# SQL Query to run
# The output CSV file will use column names from your SQL query.
# Rename them using "as" - example: "SELECT cFirstName as FirstName FROM Students"
$SqlQuery = "SELECT TOP 20
                StudentPhoto.iStudentId, 
                StudentPhoto.bImage, 
                StudentPhoto.cImageType 
            FROM 
                StudentStatus
                LEFT OUTER JOIN StudentPhoto ON StudentStatus.iStudentID=StudentPhoto.iStudentID
            WHERE 
                (StudentStatus.dInDate <=  { fn CURDATE() }) 
                AND ((StudentStatus.dOutDate < '1901-01-01') OR (StudentStatus.dOutDate >=  { fn CURDATE() }))  
                AND (StudentStatus.lOutsideStatus = 0)
            ;"

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
$SqlCommand.CommandText = $SqlQuery
$SqlCommand.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCommand
$SqlDataSet = New-Object System.Data.DataSet

# Run the SQL query
$SqlConnection.open()
$SqlAdapter.Fill($SqlDataSet)
$SqlConnection.close()

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

# Get all the student photo images
foreach($DSTable in $SqlDataSet.Tables) {
    foreach($DataRow in $DSTable){
        # WriteAllBytes wants to write in the directory above where it should for some reason
        # I have no idea why.
        [byte[]]$photoBytes = $DataRow[1]
        $StudentId = $DataRow[0]
        $FilePath = join-path $OutputLocation "$FullFileName$StudentId.jpg"
        [IO.File]::WriteAllBytes("$FilePath", $photoBytes)
    }
}

# Zip the whole scratch folder up
Compress-Archive -Path $ScratchFolderPath/*.* -DestinationPath $OutputFileName

# Delete the scratch folder, leaving only the zip file
Remove-Item $ScratchFolderPath -Recurse -Force