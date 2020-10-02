param (
    [Parameter(Mandatory=$true)][string]$EdsbyLinkOutDirectory,
    [Parameter(Mandatory=$true)][string]$OutputFileName,
    [string]$ConfigFilePath
 )

function Get-CSV {
    param(
        [Parameter(Mandatory=$true)][String] $CSVFile
    )

    return import-csv $CSVFile | Select -skip 2
}

# CSV Delimeter
# Some systems expect this to be a tab "`t" or a pipe "|".
$Delimeter = ','

# Check if the output file exists, and if it does, delete it
if (Test-Path $OutputFileName) 
{
    Remove-Item $OutputFileName
}

$CurrentLocation = Get-Location
$fullOutputFilePath = Join-Path $CurrentLocation $OutputFileName

write-host "Reading all csv files in `"$EdsbyLinkOutDirectory`""
Set-Location $EdsbyLinkOutDirectory

Get-ChildItem .\*.csv | ForEach-Object {           
    $fileOutputRows = @()

    Write-Output "Processing $($_.Name)..."     
    $inputrecords = @(Get-CSV -CSVFile $_.Name)    
    foreach($record in $inputrecords)
    {    
        $fileOutputRows += $record                
    }            

    $fileOutputRows | ForEach-Object {[PSCustomObject]$_} | export-csv $fullOutputFilePath -Append -notypeinformation -Delimiter $Delimeter
}
Set-Location $CurrentLocation

