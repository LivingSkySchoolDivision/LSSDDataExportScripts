###########################################################################
# General Utility                                                         #
###########################################################################

function Write-Log {
    param(
        [Parameter(Mandatory=$true)] $Message
    )

    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss K")> $Message"
}

Function Get-Hash {
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
        [Parameter(Mandatory=$true)][String] $CSVFile,
        [Parameter(Mandatory=$true)]$RequiredColumns
    )

    if ((Validate-CSV $CSVFile -RequiredColumns $RequiredColumns) -eq $true) {
        return import-csv $CSVFile  | Select-Object -skip 1
    } else {
        remove-module EdsbyImportModule
        throw "CSV file is not valid - cannot continue"
    }
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

function Validate-CSV {
    param(
        [Parameter(Mandatory=$true)][String] $CSVFile,
        [Parameter(Mandatory=$true)] $RequiredColumns
    )
   
    # Make sure the CSV has all the required columns for what we need

    $line = Get-Content $CSVFile -first 1

    # Check if the first row contains headings we expect
    foreach($Column in $RequiredColumns) {        
        if ($line.Contains("`"$Column`"") -eq $false) { throw "Input CSV missing field: $($Column)" }
    }
   
    return $true
}

###########################################################################
# Parsers and Converters                                                  #
###########################################################################

function Convert-StudentID {
    param(
        [Parameter(Mandatory=$true)] $InputString
    )

    return [int]$InputString.Replace("STUDENT-","")
}

function Convert-ObjectiveID {
    param(
        [Parameter(Mandatory=$true)] $OutcomeCode,
        [Parameter(Mandatory=$true)] $iCourseID,
        [Parameter(Mandatory=$true)] $Objectives
    )
    foreach($obj in $Objectives) {
        if (($obj.OutcomeCode -eq $OutcomeCode) -and ($obj.iCourseID -eq $iCourseID))
        {
            return $obj.iCourseObjectiveID
        }
    }

    return -1
}

function Convert-StaffID {
    param(
        [Parameter(Mandatory=$true)] $InputString
    )

    # We might get a list of staff, in which case we should parse it and just return the first one
    $StaffList = $InputString.Split(',')

    return [int]$StaffList[0].Replace("STAFF-","")
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

function Convert-SectionID {
    param(
        [Parameter(Mandatory=$true)] $InputString,
        [Parameter(Mandatory=$true)] $SchoolID,
        $ClassName
    )

    if ($ClassName -like 'homeroom*') {
        return 0
    } else {
        return $InputString.Replace("$SchoolID-","")
    }
}



###########################################################################
# SQL Data                                                                #
###########################################################################

$SQLQuery_CourseObjectives = "SELECT iCourseObjectiveID, OutcomeCode, iCourseID, cSubject FROM CourseObjective"
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
$SQLQuery_CourseGrades = "SELECT iCourseID, LTRIM(RTRIM(Grades.cName)) as cGrade FROM COURSE LEFT OUTER JOIN Grades ON Course.iLow_GradesID=Grades.iGradesID"
$SQLQuery_ClassCredits =    "SELECT
                                    Class.iClassID,
                                    Course.nHighCredit
                                FROM 
                                    Class
                                    LEFT OUTER JOIN Course ON Class.iCourseID=Course.iCourseID
                                WHERE
                                    Course.nHighCredit > 0"
$SQLQuery_HomeroomBlocks = "SELECT iAttendanceBlocksID as ID, iBlockNumber, cName FROM AttendanceBlocks;"
$SQLQuery_PeriodBlocks = "SELECT iBlocksID as ID, iBlockNumber, cName FROM Blocks"

function Get-CourseObjectives {
    param(
        [Parameter(Mandatory=$true)] [string] $DBConnectionString
    )

    $Raw = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_CourseObjectives
    $Processed = Convert-ObjectivesToHashtable -Objectives $Raw

    return $Processed    
}

function Get-ClassReportPeriods {
    param(
        [Parameter(Mandatory=$true)] [string] $DBConnectionString
    )

    $Raw = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_ClassReportPeriods
    $Processed = Convert-ClassReportPeriodsToHashtable -AllClassReportPeriods $Raw

    return $Processed    
}

function Get-CourseGrades {
    param(
        [Parameter(Mandatory=$true)] [string] $DBConnectionString
    )

    $Raw = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_CourseGrades
    $Processed = Convert-CourseGradesToHashTable -CourseGrades $Raw

    return $Processed    
}

function Get-AllClassCredits {
    param(
        [Parameter(Mandatory=$true)] [string] $DBConnectionString
    )

    $Raw = Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_ClassCredits
    $Processed = Convert-CourseCreditsToHashtable -CourseCreditsDataTable $Raw

    return $Processed    
}

function Get-AllHomeroomBlocks {
    param(
        [Parameter(Mandatory=$true)] [string] $DBConnectionString
    )

    return Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_HomeroomBlocks   
}

function Get-AllPeriodBlocks {
    param(
        [Parameter(Mandatory=$true)] [string] $DBConnectionString
    )

    return Get-SQLData -ConnectionString $DBConnectionString -SQLQuery $SQLQuery_PeriodBlocks
}


###########################################################################
# Data Lookups                                                            #
###########################################################################

function Get-ReportPeriodID {
    param(
        [Parameter(Mandatory=$true)] [int]$iClassID,
        [Parameter(Mandatory=$true)] [string]$RPName,
        [Parameter(Mandatory=$true)] [DateTime]$RPEndDate,
        [Parameter(Mandatory=$true)] $AllClassReportPeriods
    )

    # We only care of the date, not the time, so remove the time from the inputted date

    if ($AllClassReportPeriods.ContainsKey($iClassID)) {
        # Check end dates
        foreach($RP in $AllClassReportPeriods[$iClassID]) {
            if ($RPEndDate -eq $RP.dEndDate) {
                return [int]$($RP.iReportPeriodID)
            }
        }

        # If that didn't work, try a fuzzier search on end dates
        foreach($RP in $AllClassReportPeriods[$iClassID]) {
            if (([datetime]$RPEndDate -lt [datetime]($RP.dEndDate.AddDays(5))) -and ([datetime]($RPEndDate -gt $RP.dEndDate.AddDays(-5)))) {
                return [int]$($RP.iReportPeriodID)
            }
        }

        # If that didn't work, check names
        foreach($RP in $AllClassReportPeriods[$iClassID]) {
            if ($RPName -eq $RP.cName) {
                return [int]$($RP.iReportPeriodID)
            }
        }
    }
    
    return [int]-1
}

function Convert-EarnedCredits {
    param(
        [Parameter(Mandatory=$true)] $InputString,
        [Parameter(Mandatory=$true)][decimal]$PotentialCredits
    )

    try {
        $parsed = [decimal]$InputString
        if ($parsed -gt 49) {
            return $PotentialCredits
        } else {
            return 0
        }

    }
    catch {
        # We couldn't parse a number, so assume its a non-mark mark (like IE or NYM) or an error
        return 0
    }
}

function Get-ClassCredits {
    param(
        [Parameter(Mandatory=$true)] $iClassID,
        [Parameter(Mandatory=$true)] $AllClassCredits
    )

    if ($AllClassCredits.ContainsKey($iClassID) -eq $true) {
        return $AllClassCredits[$iClassID]
    }

    return 0
}

function Convert-ObjectivesToHashtable {
    param(
        [Parameter(Mandatory=$true)] $Objectives
    )
    $Output = New-Object -TypeName "System.Collections.ArrayList"

    foreach($Obj in $Objectives) {
        if (($Obj.OutcomeCode -ne "") -and ($null -ne $Obj.OutcomeCode)) {
                $Outcome = [PSCustomObject]@{
                OutcomeCode = $Obj.OutcomeCode
                OutcomeText = $Obj.OutcomeText
                iCourseObjectiveID = $Obj.iCourseObjectiveID
                iCourseID = $Obj.iCourseID
                cSubject = $Obj.cSubject
            }
            $Output += $Outcome

        }
    }

    return $Output
}

function Convert-CourseGradesToHashTable {
    param(
        [Parameter(Mandatory=$true)] $CourseGrades
    )
    $Output = @{}

    foreach($Obj in $CourseGrades) {
        if ($null -ne $Obj) {
            if ($null -ne $Obj.iCourseID) {
                if ($Output.ContainsKey($Obj.iCourseID) -eq $false) {
                    $Output.Add($Obj.iCourseID, $Obj.cGrade)
                }
            }
        }
    }

    return $Output
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
                    $OutPut.Add($RP.iClassID, (New-Object -TypeName "System.Collections.ArrayList"))
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

function Convert-CourseCreditsToHashtable {
    param(
        [Parameter(Mandatory=$true)] $CourseCreditsDataTable
    )

    $Output = @{}

    foreach($Obj in $CourseCreditsDataTable) {
        if ($null -ne $Obj) {
            if (($null -ne $Obj.iClassID) -and ($null -ne $Obj.nHighCredit)) {
                $Output.Add([string]$Obj.iClassID, $Obj.nHighCredit)
            }
        }
    }

    return $Output
}

###########################################################################
# Class Marks                                                             #
###########################################################################

function Convert-ToSLMark {
    param(
        [Parameter(Mandatory=$true)] $InputRow,
        [Parameter(Mandatory=$true)] $AllReportPeriods,
        [Parameter(Mandatory=$true)] $AllClassCredits
    )

    # Parse cMark vs nMark
    $cMark = ""
    $nMark = [decimal]0.0

    if ([bool]($InputRow.OverallMark -as [decimal]) -eq $true) {
        $nMark = [decimal]$InputRow.OverallMark
        if (
            ($nMark -eq 1) -or
            ($nMark -eq 1.5) -or
            ($nMark -eq 2) -or
            ($nMark -eq 2.5) -or
            ($nMark -eq 3) -or
            ($nMark -eq 3.5) -or
            ($nMark -eq 4)
        ) {
            $cMark = [string]$nMark
        }
    } else {
        $cMark = $InputRow.OverallMark
    }

    $iClassID = Convert-SectionID -SchoolID $InputRow.SchoolID -InputString $InputRow.SectionGUID
    if ($iClassID -eq -1) {
        Write-Host "ERROR PARSING CLASSID FOR $($InputRow.SectionGUID)"
        return $null
    }
    
    $iReportPeriodID = [int]((Get-ReportPeriodID -iClassID $iClassID -AllClassReportPeriods $AllReportPeriods -RPEndDate $InputRow.ReportingPeriodEndDate -RPName $InputRow.ReportingPeriodName))
    
    $NewMark = [PSCustomObject]@{
        iReportPeriodID = [int]$iReportPeriodID
        iStudentID = [int](Convert-StudentID $InputRow.StudentGUID)
        iClassID = $iClassID
        iSchoolID = $InputRow.SchoolID
        nMark = [decimal]$nMark
        cMark = [string]$cMark     
        nCredit = Convert-EarnedCredits -InputString $nMark -PotentialCredits $(Get-ClassCredits -AllClassCredits $AllClassCredits -iClassID $iClassID) 
    }

    return $NewMark
}

###########################################################################
# Outcome Marks                                                           #
###########################################################################

function Convert-ToSLOutcomeMark {
    param(
        [Parameter(Mandatory=$true)] $InputRow,
        [Parameter(Mandatory=$true)] $AllOutcomes,
        [Parameter(Mandatory=$true)] $AllReportPeriods
    )
    # Parse cMark vs nMark
    $cMark = ""
    $nMark = [decimal]0.0

    if ([bool]($InputRow.Grade -as [decimal]) -eq $true) {
        $nMark = [decimal]$InputRow.Grade
        if (
            ($nMark -eq 1) -or
            ($nMark -eq 1.5) -or
            ($nMark -eq 2) -or
            ($nMark -eq 2.5) -or
            ($nMark -eq 3) -or
            ($nMark -eq 3.5) -or
            ($nMark -eq 4)
        ) {
            $cMark = [string]$nMark
        }
    } else {
        $cMark = $InputRow.Grade
    }

    $iClassID = (Convert-SectionID -SchoolID $InputRow.SchoolID -InputString $InputRow.SectionGUID)

    $iReportPeriodID = [int]((Get-ReportPeriodID -iClassID $iClassID -AllClassReportPeriods $AllReportPeriods -RPEndDate $InputRow.ReportingPeriodEndDate -RPName $InputRow.ReportingPeriodName))

    $NewOutcomeMark = [PSCustomObject]@{
        iStudentID = [int](Convert-StudentID $InputRow.StudentGUID)
        iReportPeriodID = [int]$iReportPeriodID
        iCourseObjectiveId = [int](Convert-ObjectiveID -OutcomeCode $InputRow.CriterionName -Objectives $AllOutcomes -iCourseID $InputRow.CourseCode)
        iCourseID = [int]$InputRow.CourseCode
        iSchoolID = [int]$InputRow.SchoolID
        nMark = [decimal]$nMark
        cMark = [string]$cMark
    }

    return $NewOutcomeMark
}

###########################################################################
# Attendance                                                              #
###########################################################################
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


###########################################################################
# Comments                                                                #
###########################################################################

function Convert-ToComment {
    param(
        [Parameter(Mandatory=$true)] $InputRow,
        [Parameter(Mandatory=$true)] $AllReportPeriods
    )

    $iClassID = (Convert-SectionID -SchoolID $InputRow.SchoolID -InputString $InputRow.SectionGUID)
    $iReportPeriodID = [int]((Get-ReportPeriodID -iClassID $iClassID -AllClassReportPeriods $AllReportPeriods -RPEndDate $InputRow.ReportingPeriodEndDate -RPName $InputRow.ReportingPeriodName))

    $NewMark = [PSCustomObject]@{
        iReportPeriodID = [int]$iReportPeriodID
        iStudentID = [int](Convert-StudentID $InputRow.StudentGUID)
        iClassID = $iClassID
        iSchoolID = $InputRow.SchoolID
        mComment = $InputRow.Comment.Trim()
    }

    return $NewMark
}


###########################################################################
# SLBs                                                                    #
###########################################################################

function Convert-IndividualSLBMark {
    param(
        [Parameter(Mandatory=$true)] $InputRow,
        [Parameter(Mandatory=$true)] $OutcomeName,
        [Parameter(Mandatory=$true)] $MarkFieldName
    )

    # Parse cMark vs nMark
    $cMark = ""
    $nMark = [decimal]0.0

    $ThisOutcomeMark = $($InputRow.$MarkFieldName)
    if (($ThisOutcomeMark.Length -eq 0) -or ($null -eq $ThisOutcomeMark) -or ($ThisOutcomeMark -eq "")) {
        return $null
    }

    if ([bool]($ThisOutcomeMark -as [decimal]) -eq $true) {
        $nMark = [decimal]$ThisOutcomeMark
        if (
            ($nMark -eq 1) -or
            ($nMark -eq 1.5) -or
            ($nMark -eq 2) -or
            ($nMark -eq 2.5) -or
            ($nMark -eq 3) -or
            ($nMark -eq 3.5) -or
            ($nMark -eq 4)
        ) {
            $cMark = [string]$nMark
        }
    } else {
        $cMark = $ThisOutcomeMark
    }


    $iClassID = (Convert-SectionID -SchoolID $InputRow.SchoolID -InputString $InputRow.SectionGUID)
    $iReportPeriodID = [int]((Get-ReportPeriodID -iClassID $iClassID -AllClassReportPeriods $AllReportPeriods -RPEndDate $InputRow.ReportingPeriodEndDate -RPName $InputRow.ReportingPeriodName))

    $OutcomeCode = "$($OutcomeName.ToUpper())-$($InputRow.CourseCode)";

    $NewMark = [PSCustomObject]@{
        iStudentID = [int](Convert-StudentID $InputRow.StudentGUID)
        iReportPeriodID = [int]$iReportPeriodID
        iCourseObjectiveID = [int](Convert-ObjectiveID -OutcomeCode $OutcomeCode -Objectives $AllOutcomes -iCourseID $InputRow.CourseCode)
        iCourseID = [int]$($InputRow.CourseCode)
        iSchoolID = $InputRow.SchoolID
        cMark = [string]$cMark
        nMark = [decimal]$nMark
        OutcomeName = $OutcomeName
        OutcomeCode = $OutcomeCode
    }

    return $NewMark

}

function Get-SLBText {
    param(
        [Parameter(Mandatory=$true)][int] $CourseID,
        [Parameter(Mandatory=$true)][String] $OutcomeName,
        [Parameter(Mandatory=$true)] $AllCourseGrades

    )

    # Get grade level of the given course
    $GradeLevel = -1;
    if ($AllCourseGrades.ContainsKey($CourseID)) {
        $GradeLevel = $AllCourseGrades[$CourseID]
    }

    # Try to process what the outcome text would be for that grade.
    if ($GradeLevel -ne -1) {
        if (
            ($GradeLevel -eq "pk") -or
            ($GradeLevel -eq "0k") -or
            ($GradeLevel -eq "01") -or
            ($GradeLevel -eq "02") -or
            ($GradeLevel -eq "03") -or
            ($GradeLevel -eq "04") -or
            ($GradeLevel -eq "05") -or
            ($GradeLevel -eq "06")
        )
        {
            if ($OutcomeName -like "CITIZENSHIP") {
                return "Respectful, shows caring, takes responsibility for actions."
            }
            if ($OutcomeName -like "COLLABORATIVE") {
                return "Willing to work with all classmates, encourages and includes others."
            }
            if ($OutcomeName -like "ENGAGEMENT") {
                return "Wants to learn and keeps trying when the work gets hard."
            }
            if ($OutcomeName -like "SELF-DIRECTED") {
                return "Takes initiative, completes tasks, strong work habits."
            }
        }
        if (
            ($GradeLevel -eq "07") -or
            ($GradeLevel -eq "08") -or
            ($GradeLevel -eq "09")
        )
        {
            if ($OutcomeName -like "CITIZENSHIP") {
                return "Respectful to others and property, takes responsibility for actions and decisions."
            }
            if ($OutcomeName -like "COLLABORATIVE") {
                return "Willing to work with all classmates, encourages and includes others."
            }
            if ($OutcomeName -like "ENGAGEMENT") {
                return "Involved in the learning tasks."
            }
            if ($OutcomeName -like "SELF-DIRECTED") {
                return "Takes initiative, completes tasks, strong work habits."
            }

        }
        if (
            ($GradeLevel -eq "10") -or
            ($GradeLevel -eq "11") -or
            ($GradeLevel -eq "12")
        )
        {
            if ($OutcomeName -like "CITIZENSHIP") {
                return "Respectful, responsible, academically honest."
            }
            if ($OutcomeName -like "COLLABORATIVE") {
                return "Offers and receives ideas while working with others."
            }
            if ($OutcomeName -like "ENGAGEMENT") {
                return "Involved in the learning tasks."
            }
            if ($OutcomeName -like "SELF-DIRECTED") {
                return "Takes initiative, completes tasks, strong work habits."
            }
        }

        return "$OutcomeName for grade $GradeLevel"
    }

    return "UNKNOWN OUTCOME FOR COURSE $($CourseID) AND GRADE $($GradeLevel) PLEASE CONTACT SIS SUPPORT DESK"
}

function Convert-ToSLB {
    param(
        [Parameter(Mandatory=$true)] $InputRow,
        [Parameter(Mandatory=$true)] $AllOutcomes,
        [Parameter(Mandatory=$true)] $AllReportPeriods
    )

    # This needs to output _4_ marks, one for each outcome. The consumer of this function will need to handle getting an array back.
    # "Citizenship","Collaboration","Engagement","SelfDirected"

    $Output = New-Object -TypeName "System.Collections.ArrayList"

    # Parse Citizenship mark
    $Mark_Citizenship = Convert-IndividualSLBMark -InputRow $InputRow -OutcomeName "CITIZENSHIP" -MarkFieldName "Citizenship"
    if ($null -ne $Mark_Citizenship) {
        $Output += $Mark_Citizenship
    }

    # Parse Collaboration mark
    $Mark_Collaboration = Convert-IndividualSLBMark -InputRow $InputRow -OutcomeName "COLLABORATIVE" -MarkFieldName "Collaboration"
    if ($null -ne $Mark_Collaboration) {
        $Output += $Mark_Collaboration
    }

    # Parse Engagement mark
    $Mark_Engagement = Convert-IndividualSLBMark -InputRow $InputRow -OutcomeName "ENGAGEMENT" -MarkFieldName "Engagement"
    if ($null -ne $Mark_Engagement) {
        $Output += $Mark_Engagement
    }

    # Parse Self-Directed mark
    $Mark_SelfDirected = Convert-IndividualSLBMark -InputRow $InputRow -OutcomeName "SELF-DIRECTED" -MarkFieldName "SelfDirected"
    if ($null -ne $Mark_SelfDirected) {
        $Output += $Mark_SelfDirected
    }

    return $Output
}