
<#
.DESCRIPTION
    Combine-MDATPExposedMachines provides a quick and dirty way to combine multiple TVM exposed machine reports
    into a single file, ingest AD information for machines in MDATP, and provide quick metrics on vulnerability
    or machine counts in an exported excel file with pivottables.
    
.DEPENDENCIES
    You must have the ImportExcel module installed
    You must have access to MDATP Threat and Vulnerability Management to download exposed machines reports
    You must have access to Active Directory to fill in the canonical name and OS info
    
.NOTES
    Version       : 1.0
    Author        : Blake Bourgeois
    Creation Date : 6/18/2020
#>

## set this to wherever you're going to store the exposed machines exports! ##
$pwd = "$ENV:UserProfile\Downloads"

# gets all the CSV exports from the folder
$reports = get-childitem $pwd -Include *.csv -Depth 0

#ATP is inserts a malformed header if you're trying to pull this into PS, this deletes the header...but we want to make sure it only happens if the "Assets Export" header is present
foreach($report in $reports){
    $line = get-content $report.FullName -First 1
    $line = $line.Split(" ")[0]
    if($line -eq 'Assets'){
        (get-content $report.FullName | Select-Object -Skip 1) | Set-Content $report.FullName
        }
    }


$date = get-date -Format yyyyMMdd

# use the CSV filename (and remove the 'exposed machines' bit) to populate a vulnerability column in the combined master file
foreach($report in $reports){
    $base = ($report.BaseName -split "\s-\s")[0]
    Import-Csv $report.FullName | Select-Object *,@{Name='Finding';Expression={$base}} | Export-Csv $pwd\$date.csv -Append -NoTypeInformation
    }

# pull the results back in so we can tie into AD
$output = import-csv $pwd\$date.csv
# update each line with AD information
# for example I like sorting by canonical name to group machines together by department/location however your AD is configured
$output | foreach-object {
    # reinitialize as to not fill the spreadsheet with junk duplicates
    $cn = ""
    $os = ""
    $osv = ""
    $computer = ""
    $computer = get-adcomputer $_.Name -Properties OperatingSystem,OperatingSystemVersion,CanonicalName 
    $cn = $computer.CanonicalName
    $os = $computer.OperatingSystem
    $osv = $computer.OperatingSystemVersion

    if($cn){
        $_ | Add-Member -MemberType NoteProperty -Name "AD Path" -Value $cn}
    if($os){
        $_ | Add-Member -MemberType NoteProperty -Name "OS" -Value $os}
    if($osv){
        $_ | Add-Member -MemberType NoteProperty -Name "OS Version" -Value $osv}
    }

# save back over the intial CSV with the extra data...just in case
$output | export-csv $pwd\$date.csv -NoTypeInformation


# creates a master export of all the vulns and pivots
# with credit to this great post https://jamesone111.wordpress.com/2017/12/12/using-the-import-excel-module-part-3-pivots-and-charts-data-and-calculations/
$xl = ""
$xl = $output | Export-Excel -Path $pwd\$date-MDATP_Full.xlsx -WorksheetName "Vulnerabilities" -TableName "Vulnerabilities" -PassThru
$Pt1 = New-PivotTableDefinition -PivotTableName "Machines by Vulnerability Count" -PivotData @{"Name" = "Count"} -SourceWorkSheet "Vulnerabilities" -PivotRows Name
$Pt2 = New-PivotTableDefinition -PivotTableName "Vulnerability by Machine Count" -PivotData @{"Finding" = "Count"} -SourceWorkSheet "Vulnerabilities" -PivotRows Finding
$xl = Export-Excel -ExcelPackage $xl -WorksheetName "Vulnerabilities" -PivotTableDefinition $Pt1 -PassThru
Export-Excel -ExcelPackage $xl -WorksheetName "Vulnerabilities" -PivotTableDefinition $Pt2 #-Show
