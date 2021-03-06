Import-Module ActiveDirectory

$reportFile = "c:\temp\rep.html"

$act_step1 = "Retrieving computer information from ActiveDirectory" 
$act_step2 = "Testing connection to servers"
$act_step3 = "Preparing output"
$act_step4 = "Saving report"



write-progress -Activity $act_step1 -Status "Processing..."

$serverList = Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, Enabled, LastLogonDate, IPv4Address, IPv6Address, DNSHostName, DistinguishedName, cn, CanonicalName | `
    Where-Object{ ( $_.operatingSystem ) -and ( $_.operatingSystem.toLower() ).contains( "server" ) } 

write-progress -Activity $act_step1 -PercentComplete 100 -Status "Processing..."

$outputList = @()
$count = 0
foreach( $server in $serverList )
{
    $count++
    Write-Progress -Activity $act_step2 -Status "Contacting $($server.cn)..." -PercentComplete ( $count * 100 / $serverList.Count )
    $test = Test-Connection -ComputerName $server.cn -Count 3 -ea SilentlyContinue
    if ( $test )
    {
       $online = $true
       try
       {
           $os = Get-WmiObject Win32_OperatingSystem -ComputerName $server.cn -ea Stop
           $LastBootUpTime = $os.ConvertToDateTime($os.LastBootUpTime)
           $LocalDateTime = $os.ConvertToDateTime($os.LocalDateTime)
           $up = $LocalDateTime - $LastBootUpTime
           $uptime = "$($up.Days) days $($up.Hours) hours"
       }
       catch
       {
           $uptime = "not available"
       }
       $lastLogon = "" 
       $ip = $test[2].ipv4address.IPAddressToString
    }
    else
    {
       $online = $false
       $ipaddress = ""
       $uptime = ""
       $lastLogon = "$(( [datetime]::Now - $server.LastLogonDate ).Days) days ago"
    }
    $osver = "{0} {1}" -F $server.OperatingSystem, $server.OperatingSystemServicePack
    $item = new-object psobject
    Add-Member -InputObject $item -MemberType NoteProperty -Name "Online" -Value $online
    Add-Member -InputObject $item -MemberType NoteProperty -Name "Uptime" -Value $uptime
    Add-Member -InputObject $item -MemberType NoteProperty -Name "Last logon" -Value $lastLogon
    Add-Member -InputObject $item -MemberType NoteProperty -Name "Name" -Value $server.cn
    Add-Member -InputObject $item -MemberType NoteProperty -Name "OS Version" -Value $osver
    Add-Member -InputObject $item -MemberType NoteProperty -Name "Path" -Value $server.CanonicalName
    Add-Member -InputObject $item -MemberType NoteProperty -Name "IP Address" -Value $ip
    Add-Member -InputObject $item -MemberType NoteProperty -Name "Enabled" -Value $server.Enabled
    $outputList += $item
}

Write-Progress -Activity $act_step3  -Status "Processing..."

$table_online = $outputList | ? { $_.Online } | Sort-Object Name | ConvertTo-Html Name, "OS Version", "IP Address", Uptime, Path -Fragment

Write-Progress -Activity $act_step3  -PercentComplete 33 -Status "Processing..."

$table_offline = $outputList | ? { -not $_.Online } | Sort-Object Name | ConvertTo-Html Name, "OS Version", "Last logon", Enabled, Path -Fragment

Write-Progress -Activity $act_step3 -PercentComplete 66 -Status "Processing..."
 
$report = " 
<!DOCTYPE html>
<html>
<head>
<style>
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;white-space:nowrap;} 
TH{border-width: 1px;padding: 4px;border-style: solid;border-color: black} 
TD{border-width: 1px;padding: 2px 10px;border-style: solid;border-color: black} 
</style>
</head>
<body> 
<H2>Server report</H2> 
<H3>Online</H3> 
$table_online
<H3>Offline</H3> 
$table_offline
</body>
</html>"  

Write-Progress -Activity $act_step4 -Status "Processing..."


$report  | Set-Content $reportFile -Force 

Write-Progress -Activity $act_step4 -Status "Processing..." -Completed


Invoke-Expression $reportFile 

