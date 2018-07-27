#--------------------------------------------------------------------
# Parameters
param (
   [Parameter(mandatory=$false)] [String]$path

)
$patch = "C:\Users\thart\Documents\"
#--------------------------------------------------------------------
# User Defined Variables

# Open csv file after creation
$autoLaunch = $false

#--------------------------------------------------------------------
# Static Variables

$scriptName = "VeeamJobDetail"
$scriptVer = "1.0"
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$starttime = Get-Date -uformat "%m-%d-%Y %I:%M:%S"
$allDetails = @()

#--------------------------------------------------------------------
# Load Snap-ins

# Add Veeam snap-in if required
#If ((Get-PSSnapin -Name VeeamPSSnapin -ErrorAction SilentlyContinue) -eq $null) {add-pssnapin VeeamPSSnapin}

#--------------------------------------------------------------------
# Functions

#--------------------------------------------------------------------
# Main Procedures

#Clear-Host
Write-Host "********************************************************************************"
Write-Host "$scriptName`tVer:$scriptVer`t`t`tStart Time:`t$starttime"
Write-Host "********************************************************************************`n"

# Get Backup Jobs
$jobs = Get-VBRJob | ?{$_.JobType -eq "Backup"}

# Loop through each job adding details to array
foreach ($job in $jobs) {
   $jobOptions = New-Object PSObject
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Name" -value $job.name
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Enabled" -value $job.isscheduleenabled
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Backup Mode" -value $job.backuptargetoptions.algorithm
   $repo = (Get-VBRBackupRepository | ?{$_.HostId -eq $job.TargetHostId -and $_.Path -eq $job.TargetDir}).name
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Repository" -value $repo
   $proxies = $null
   foreach ($prox in ($job | get-vbrjobproxy)) {
      $pName = $prox.Name
      $proxies = $proxies + $pName
   }
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Proxy" -value $proxies
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Auto Proxy" -Value $job.sourceproxyautodetect
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Next Run" -Value $job.scheduleoptions.nextrun
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Restore Points" -Value $job.backupstorageoptions.retaincycles
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Deduplication" -Value $job.backupstorageoptions.enablededuplication
   $comp = $job.backupstorageoptions.compressionlevel
   If ($comp -eq 0) {$comp = "None"}
   If ($comp -eq 4) {$comp = "Dedupe Friendly"}
   If ($comp -eq 5) {$comp = "Optimal"}
   If ($comp -eq 6) {$comp = "High"}
   If ($comp -eq 9) {$comp = "Extreme"}
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Compression" -Value $comp
   $opti = $job.backupstorageoptions.stgblocksize
   If ($opti -eq "KbBlockSize8192") {$opti = "Local Target(16TB+ Files)"}
   If ($opti -eq "KbBlockSize1024") {$opti = "Local Target"}
   If ($opti -eq "KbBlockSize512") {$opti = "LAN Target"}
   If ($opti -eq "KbBlockSize256") {$opti = "WAN Target"}
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Optimized" -Value $opti
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Integrity Checks" -Value $job.backupstorageoptions.enableintegritychecks
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Exclude Swap" -Value $job.visourceoptions.excludeswapfile
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Remove Deleted VMs" -Value $job.backupstorageoptions.enabledeletedvmdataretention
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Retain Deleted VMs" -Value $job.backupstorageoptions.retaindays
   $jobOptions | Add-Member -MemberType NoteProperty -Name "CBT Enabled" -Value $job.visourceoptions.usechangetracking
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Auto Enable CBT" -Value $job.visourceoptions.enablechangetracking
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Set VM Note" -Value $job.visourceoptions.setresultstovmnotes
   $jobOptions | Add-Member -MemberType NoteProperty -Name "VM Attribute Name" -Value $job.visourceoptions.vmattributename
   $jobOptions | Add-Member -MemberType NoteProperty -Name "VMTools Quiesce" -Value $job.visourceoptions.vmtoolsquiesce
   $jobOptions | Add-Member -MemberType NoteProperty -Name "VSS Enabled" -Value $job.vssoptions.enabled
   $igfs = $job.vssoptions.guestfsindexingtype
   If ($igfs -eq "None") {$igfs = "Disabled"}
   ElseIf ($igfs -eq "EveryFolders") {$igfs = "Enabled"}
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Index Guest FS" -Value $igfs
   $jobOptions | Add-Member -MemberType NoteProperty -Name "VSS Username" -Value $($job | get-vbrjobvssoptions).credentials.username
   $jobOptions | Add-Member -MemberType NoteProperty -Name "Description" -Value $job.Description
   $allDetails += $jobOptions
}

#--------------------------------------------------------------------
# Outputs

# Display results summary
$allDetails | select Name, Enabled | Sort Name | ft -AutoSize

If (!$path -or !$path.EndsWith(".csv")) {
   Write-Host "`n`nUsing Default Path"
   $path = $scriptDir + "\" + $scriptName + "_" + (Get-Date -uformat %m-%d-%Y_%I-%M-%S) + ".csv"
   $path
} Else {
   Write-Host "`n`nUsing Supplied Path"
   $path
}

# Export results
$allDetails | Sort Name | Export-Csv C:\Users\thart\Documents\veeam.csv -NoTypeInformation -Force

# Open csv
If ($autoLaunch) {
   Invoke-Item $path
}

$finishtime = Get-Date -uformat "%m-%d-%Y %I:%M:%S"
Write-Host "`n`n"
Write-Host "********************************************************************************"
Write-Host "$scriptName`t`t`t`tFinish Time:`t$finishtime"
Write-Host "********************************************************************************"

<# Prompt to exit script - This leaves PS window open when run via right-click
Write-Host "`n`n"
Write-Host "Press any key to continue ..." -foregroundcolor Gray
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
#>