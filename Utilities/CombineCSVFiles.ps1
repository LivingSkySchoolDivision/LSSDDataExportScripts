param (
    [Parameter(Mandatory=$true)][string]$InputDirectory,
    [Parameter(Mandatory=$true)][string]$OutputFileName,
    [Parameter(Mandatory=$true)][string]$FileFilter,
    [string]$HeaderLines = 1
 )

# Check if the output file exists, and if it does, delete it
if (Test-Path $OutputFileName) 
{
    Remove-Item $OutputFileName
}

write-host "Reading all csv files in `"$InputDirectory`" with filter `"$FileFilter`""

# Get the file headings of each file, and make sure they all match
# Save the file headings for later use

if ($HeaderLines -gt 0) {
    $FileHeadingsRow = ""

    Get-ChildItem -Path $InputDirectory -Filter $FileFilter | ForEach-Object {           
        $firstline = Get-Content $_.FullName -first 1

        Write-Output "Inspecting $($_.FullName)..."    
        if ($FileHeadingsRow -eq "") { 
            $FileHeadingsRow = $firstline
        } else {
            if ($firstline.Equals($FileHeadingsRow) -eq $false) {
                throw "Headings of file `"$_.FullName`" does not match. Aborting."
            }
        }
    }
}


$file = [system.io.file]::OpenWrite($OutputFileName)
$writer = New-Object System.IO.StreamWriter($file)

# Write the file headings
if ($HeaderLines -gt 0) {
    $writer.WriteLine($FileHeadingsRow)
}

Get-ChildItem -Path $InputDirectory -Filter $FileFilter | ForEach-Object {           
    $fileOutputRows = @()

    Write-Output "Processing $($_.FullName)..."     
    (cat $($_.FullName) | Select-Object -Skip $HeaderLines) | ForEach-Object { $writer.WriteLine($_) }
}

$writer.Close()
$file.Close()

Write-Output "Done!"