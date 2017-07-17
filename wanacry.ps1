

$OffComputers = @()
$CheckFail = @()
$Patched = @()
$Unpatched = @()

$log = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath "WannaCry patch state for $($ENV:USERDOMAIN).log"

$Patches = @('KB4012219', 'KB4012213', 'KB4012214', 'KB4012215', 'KB4012216', 'KB4012217', 'KB4012598', 'KB4013429', 'KB4015217', 'KB4015438', 'KB4015549', 'KB4015550', 'KB4015551', 'KB4015552', 'KB4015553', 'KB4016635', 'KB4019215', 'KB4019216', 'KB4019264', 'KB4019472')

$WindowsComputers = (Get-ADComputer -Filter {
    (OperatingSystem  -Like 'Windows*') -and (OperatingSystem -notlike '*Windows 10*')
}).Name|
Sort-Object

"WannaCry patch status $(Get-Date -Format 'yyyy-MM-dd HH:mm')" |Out-File -FilePath $log

$ComputerCount = $WindowsComputers.count
"There are $ComputerCount computers to check"
$loop = 0
foreach($Computer in $WindowsComputers)
{
  $ThisComputerPatches = @()
  $loop ++
  "$loop of $ComputerCount `t$Computer"
  try
  {
    $null = Test-Connection -ComputerName $Computer -Count 1 -ErrorAction Stop
    try
    {
      $Hotfixes = Get-HotFix -ComputerName $Computer -ErrorAction Stop

      $Patches | ForEach-Object -Process {
        if($Hotfixes.HotFixID -contains $_)
        {
          $ThisComputerPatches += $_
        }
      }
    }
    catch
    {
      $CheckFail += $Computer
      "***`t$Computer `tUnable to gather hotfix information" |Out-File -FilePath $log -Append
      continue
    }
    If($ThisComputerPatches)
    {
      "$Computer is patched with $($ThisComputerPatches -join (','))" |Out-File -FilePath $log -Append
      $Patched += $Computer
    }
    Else
    {
      $Unpatched += $Computer
      "*****`t$Computer IS UNPATCHED! *****" |Out-File -FilePath $log -Append
    }
  }
  catch
  {
    $OffComputers += $Computer
    "****`t$Computer `tUnable to connect." |Out-File -FilePath $log -Append
  }
}
' '
"Summary for domain: $ENV:USERDNSDOMAIN"
"Unpatched ($($Unpatched.count)):" |Out-File -FilePath $log -Append
$Unpatched -join (', ')  |Out-File -FilePath $log -Append
'' |Out-File -FilePath $log -Append
"Patched ($($Patched.count)):" |Out-File -FilePath $log -Append
$Patched -join (', ') |Out-File -FilePath $log -Append
'' |Out-File -FilePath $log -Append
"Off/Untested($(($OffComputers + $CheckFail).count)):"|Out-File -FilePath $log -Append
($OffComputers + $CheckFail | Sort-Object)-join (', ')|Out-File -FilePath $log -Append

"Of the $($WindowsComputers.count) windows computers in active directory, $($OffComputers.count) were off, $($CheckFail.count) couldn't be checked, $($Unpatched.count) were unpatched and $($Patched.count) were successfully patched."
'Full details in the log file.'

try
{
  Start-Process -FilePath notepad++ -ArgumentList $log
}
catch
{
  Start-Process -FilePath notepad.exe -ArgumentList $log
}[reflection.assembly]::LoadWithPartialName("System.Version")
$os = Get-WmiObject -class Win32_OperatingSystem
$osName = $os.Caption
$s = "%systemroot%\system32\drivers\srv.sys"
$v = [System.Environment]::ExpandEnvironmentVariables($s)
If (Test-Path "$v")
    {
    Try
        {
        $versionInfo = (Get-Item $v).VersionInfo
        $versionString = "$($versionInfo.FileMajorPart).$($versionInfo.FileMinorPart).$($versionInfo.FileBuildPart).$($versionInfo.FilePrivatePart)"
        $fileVersion = New-Object System.Version($versionString)
        }
    Catch
        {
        Write-Host "Unable to retrieve file version info, please verify vulnerability state manually." -ForegroundColor Yellow
        Return
        }
    }
Else
    {
    Write-Host "Srv.sys does not exist, please verify vulnerability state manually." -ForegroundColor Yellow
    Return
    }
if ($osName.Contains("Vista") -or ($osName.Contains("2008") -and -not $osName.Contains("R2")))
    {
    if ($versionString.Split('.')[3][0] -eq "1")
        {
        $currentOS = "$osName GDR"
        $expectedVersion = New-Object System.Version("6.0.6002.19743")
        } 
    elseif ($versionString.Split('.')[3][0] -eq "2")
        {
        $currentOS = "$osName LDR"
        $expectedVersion = New-Object System.Version("6.0.6002.24067")
        }
    else
        {
        $currentOS = "$osName"
        $expectedVersion = New-Object System.Version("9.9.9999.99999")
        }
    }
elseif ($osName.Contains("Windows 7") -or ($osName.Contains("2008 R2")))
    {
    $currentOS = "$osName LDR"
    $expectedVersion = New-Object System.Version("6.1.7601.23689")
    }
elseif ($osName.Contains("Windows 8.1") -or $osName.Contains("2012 R2"))
    {
    $currentOS = "$osName LDR"
    $expectedVersion = New-Object System.Version("6.3.9600.18604")
    }
elseif ($osName.Contains("Windows 8") -or $osName.Contains("2012"))
    {
    $currentOS = "$osName LDR"
    $expectedVersion = New-Object System.Version("6.2.9200.22099")
    }
elseif ($osName.Contains("Windows 10"))
    {
    if ($os.BuildNumber -eq "10240")
        {
        $currentOS = "$osName TH1"
        $expectedVersion = New-Object System.Version("10.0.10240.17319")
        }
    elseif ($os.BuildNumber -eq "10586")
        {
        $currentOS = "$osName TH2"
        $expectedVersion = New-Object System.Version("10.0.10586.839")
        }
    elseif ($os.BuildNumber -eq "14393")
        {
        $currentOS = "$($osName) RS1"
        $expectedVersion = New-Object System.Version("10.0.14393.953")
        }
    elseif ($os.BuildNumber -eq "15063")
        {
        $currentOS = "$osName RS2"
        "No need to Patch. RS2 is released as patched. "
        return
        }
    }
elseif ($osName.Contains("2016"))
    {
    $currentOS = "$osName"
    $expectedVersion = New-Object System.Version("10.0.14393.953")
    }
elseif ($osName.Contains("Windows XP"))
    {
    $currentOS = "$osName"
    $expectedVersion = New-Object System.Version("5.1.2600.7208")
    }
elseif ($osName.Contains("Server 2003"))
    {
    $currentOS = "$osName"
    $expectedVersion = New-Object System.Version("5.2.3790.6021")
    }
else
    {
    Write-Host "Unable to determine OS applicability, please verify vulnerability state manually." -ForegroundColor Yellow
    $currentOS = "$osName"
    $expectedVersion = New-Object System.Version("9.9.9999.99999")
    }
Write-Host "`n`nCurrent OS: $currentOS (Build Number $($os.BuildNumber))" -ForegroundColor Cyan
Write-Host "`nExpected Version of srv.sys: $($expectedVersion.ToString())" -ForegroundColor Cyan
Write-Host "`nActual Version of srv.sys: $($fileVersion.ToString())" -ForegroundColor Cyan
If ($($fileVersion.CompareTo($expectedVersion)) -lt 0)
    {
    Write-Host "`n`n"
    Write-Host "System is NOT Patched" -ForegroundColor Red
    }
Else
    {
    Write-Host "`n`n"
    Write-Host "System is Patched" -ForegroundColor Green
    }
#
