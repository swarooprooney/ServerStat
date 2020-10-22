function get-customServerResult{
$excel = New-Object -Com Excel.Application
$target = $null
while ($target -eq $null){
$target = read-host "Enter source/working directory name"
if (-not(test-path $target)){
    Write-host "Invalid directory path, re-enter."
    $target = $null
    }
elseif (-not (get-item $target).psiscontainer){
    Write-host "Target must be a directory, re-enter."
    $target = $null
    }
}
$wb = $excel.Workbooks.Open("$target\Server.xlsx")
$worksheet = $wb.sheets.item("Sheet1")
$intRowMax =  ($worksheet.UsedRange.Rows).count
$Columnnumber = 1
$outPathPing = "$target\Ping"
$outPathTrace = "$target\Trace"
$ipaddress = $(ipconfig | where {$_ -match 'IPv4.+\s(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})' } | out-null; $Matches[1])
for($intRow = 2 ; $intRow -le $intRowMax ; $intRow++)
{
 $dateTime = get-date -Format ddMMyyyyhhmmss
 $name = $worksheet.cells.item($intRow,$ColumnNumber).value2
 "pinging $name ..."
 $result = ping $name 
 $result + "The source IP :$ipaddress" | Out-File "$outPathping$name$dateTime.txt"
 "tracert $name..."
 $result = tracert $name
 $result +  "The source IP :$ipaddress"  | Out-File "$outPathTrace$name$dateTime.txt"
}
$excel.Quit()
}
function get-defaultResult{
$serverArray = "icehtml.cop.eme.uk","retailservices.direct.services.e-ssi.net","retaillegacyservices.direct.services.e-ssi.net","icews.corp.pg.eon.net"
$myDesktop = [Environment]::GetFolderPath("Desktop")
"The results will be stored at $myDesktop" -ForegroundColor Green
$outPathPing = "$myDesktop\Ping"
$outPathTrace = "$myDesktop\Trace"
$ipaddress = $(ipconfig | where {$_ -match 'IPv4.+\s(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})' } | out-null; $Matches[1])
for($intRow = 0 ; $intRow -le 3 ; $intRow++)
{
 $dateTime = get-date -Format ddMMyyyyhhmmss
 $name = $serverArray[$intRow]
 "pinging $name ..."
 $result = ping $name 
 $result + "The source IP :$ipaddress" | Out-File "$outPathping$name$dateTime.txt"
 "tracert $name..."
 $result = tracert $name
 $result +  "The source IP :$ipaddress"  | Out-File "$outPathTrace$name$dateTime.txt"
}

 

}
Write-host "Would you like to specify the servers(Default is No). Use No to check default ICE servers" -ForegroundColor Yellow
$Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) 
     { 
       Y 
       {
        Write-host "Yes, Custom Server"; get-customServerResult
       } 
       N {
       Write-Host "No, use servers in the script."; get-defaultResult
       } 
       Default {
       Write-Host "Default, use servers in the script"; get-defaultResult
       } 
     } 