param (
    [Parameter(Mandatory=$true)][string]$InputDirectory,
    [Parameter(Mandatory=$true)][string]$OutputFileName,
    [Parameter(Mandatory=$true)][string]$FileFilter,
    [Parameter][int]$Skip
 )

function Get-CSV {
    param(
        [Parameter(Mandatory=$true)][String] $CSVFile
    )

    return import-csv $CSVFile | Select-Object -skip $Skip
}

# CSV Delimeter
# Some systems expect this to be a tab "`t" or a pipe "|".
$Delimeter = ','

# Check if the output file exists, and if it does, delete it
if (Test-Path $OutputFileName) 
{
    Remove-Item $OutputFileName
}

write-host "Reading all csv files in `"$InputDirectory`" with filter `"$FileFilter`""

Get-ChildItem -Path $InputDirectory -Filter $FileFilter | ForEach-Object {           
    $fileOutputRows = @()

    Write-Output "Processing $($_.FullName)..."     
    $inputrecords = @(Get-CSV -CSVFile $_.FullName)    
    foreach($record in $inputrecords)
    {    
        $fileOutputRows += $record                
    }            

    $fileOutputRows | ForEach-Object {[PSCustomObject]$_} | export-csv -Path ($OutputFileName) -Append -notypeinformation -Delimiter $Delimeter
}
