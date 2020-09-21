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
$SqlQuery = "SELECT 
                LTRIM(RTRIM(Assignment.[Facility Code])) as School_id,
                LTRIM(RTRIM(Employee.No_)) as Staff_id,
                LTRIM(RTRIM(Employee.[Company E-Mail])) as Staff_email,
                LTRIM(RTRIM(Employee.[First Name])) as First_name,
                LTRIM(RTRIM(Employee.[Last Name])) as Last_name,
                '' as Department,
                LTRIM(RTRIM(Employee.[Job Title])) as Title,	
                LTRIM(RTRIM(Employee.[Company E-Mail])) as Username,
                '' as Password,
                '' as Role
            FROM 
                [Living Sky Live co_`$Employee] as Employee
                LEFT OUTER JOIN [Living Sky Live co_`$Employee Assignment] as Assignment ON Employee.No_=Assignment.[Employee No_]
            WHERE
                Assignment.[Contract Start Date] <= GETDATE()
                AND (Assignment.[End Date] >= GETDATE()
                    OR Assignment.[End Date] < '1900-01-01')
                AND Assignment.[Employer No_] <> 'CASUAL'
                AND Assignment.[Employer No_] <> 'SUB'
                AND Employee.[Company E-Mail] IS NOT NULL
                AND Employee.[Company E-Mail] <> '';"

# DANs for facilities, since these aren't in Navision
# These are the provincial assigned codes, retreived from
# the schoollogic database (mostly)
$SchoolDANs = @{
    #'0001'='DEFAULT_DISTRICT_OFFICE'; # Division Based
    '0002'='DEFAULT_DISTRICT_OFFICE'; # Central Office
    '0004'='DEFAULT_DISTRICT_OFFICE'; # Bus Ops
    '0006'='DEFAULT_DISTRICT_OFFICE'; # Facilities Maintenance
    '0011'='DEFAULT_DISTRICT_OFFICE'; # All Schools
    '0012'='5810211'; # BCS
    '0014'='5850201'; # Bready
    '0016'='5850401'; # Connaught
    '0018'='5910111'; # CKES (Closed)
    '0020'='6410721'; # HCES
    '0022'='5850501'; # Lawrence
    '0024'='5850601'; # McKitrick
    '0028'='5810221'; # STVital
    '0030'='5910711'; # UPS
    '0032'='5910911'; # NCES
    '0034'='5910123'; # CKCS
    '0036'='6410713'; # SHS
    '0038'='5910813'; # UCHS
    '0040'='5910923'; # McLurg
    '0042'='5850904'; # NBCHS
    '0044'='5010213'; # Cando
    '0048'='5710213'; # Hafford
    '0050'='4410223'; # Kerrobert
    '0052'='6410313'; # Leoville
    '0054'='4410323'; # Luseland
    '0056'='4410413'; # Macklin
    '0060'='5810713'; # Maymont
    '0062'='6410513'; # Medstead
    '0066'='5894003'; # Heritage
    '0067'='6694003'; # MLCA 
    '0072'='5910313'; # Hillsvale
    '0074'='5911011'; # Lakeview
    '0076'='5911113'; # Scott
    '0516'='5850401'; # Connaught
    '0616'='5850401'; # Connaught
    '0624'='5850601'; # McKitrick
    '0090'='5850801'; # LSSD Virtual
}

# CSV Delimeter
# Some systems expect this to be a tab "`t" or a pipe "|".
$Delimeter = ','

# Should all columns be quoted, or just those that contains characters to escape?
# Note: This has no effect on systems with PowerShell versions <7.0 (all fields will be quoted otherwise)
$QuoteAllColumns = $false

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
$ConnectionString = $configXML.Settings.Navision.ConnectionString

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

# Post processing
# Convert the NAV facility codes to school DAN numbers
# and toss out any that don't match the hashtable above

foreach($DSTable in $SqlDataSet.Tables) {    
    $RowsToRemove = @()
    foreach($DataRow in $DSTable) {
        if ($SchoolDANs.ContainsKey($DataRow["School_id"]) -eq $false) {
            $RowsToRemove += $DataRow
        } else {
            $DataRow["School_id"] = $SchoolDANs[$DataRow["School_id"]]
        }
    }

    foreach($Row in $RowsToRemove) {
        $DSTable.Rows.Remove($Row)
    }
}

# Output to a CSV file
foreach($DSTable in $SqlDataSet.Tables) {
    if (($PSVersionTable.PSVersion.Major -ge 7) -and ($QuoteAllColumns -eq $false)) {
        $DSTable | export-csv $OutputFileName -notypeinformation -Delimiter $Delimeter -UseQuotes AsNeeded
    } else {        
        $DSTable | export-csv $OutputFileName -notypeinformation -Delimiter $Delimeter
    }
}