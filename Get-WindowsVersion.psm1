function Get-WindowsVersion {
    <#
.SYNOPSIS
    List Windows Version from computer. Compatible with PSVersion 3 or higher.
.DESCRIPTION
    List Windows Version from computer. Compatible with PSVersion 3 or higher.
.PARAMETER ComputerName
    Name of server to list Windows Version from remote computer.
.PARAMETER SearchBase
    AD-SearchBase of server to list Windows Version from remote computer.
.PARAMETER History
    List History Windows Version from computer.
.PARAMETER Force
    Disable the built-in Format-Table and Sort-Object.
.NOTES
    Name: Get-WindowsVersion.psm1
    Author: Johannes Sebald
    Version: 1.2.5
    DateCreated: 2016-09-13
    DateEdit: 2018-07-11
.LINK
    https://www.dertechblog.de
.EXAMPLE
    Get-WindowsVersion
    List Windows Version on local computer with built-in Format-Table and Sort-Object.
.EXAMPLE
    Get-WindowsVersion -ComputerName pc1
    List Windows Version on remote computer with built-in Format-Table and Sort-Object.
.EXAMPLE
    Get-WindowsVersion -ComputerName pc1,pc2
    List Windows Version on multiple remote computer with built-in Format-Table and Sort-Object.
.EXAMPLE
    Get-WindowsVersion -SearchBase "OU=Computers,DC=comodo,DC=com" with built-in Format-Table and Sort-Object.
    List Windows Version on Active Directory SearchBase computer.
.EXAMPLE
    Get-WindowsVersion -ComputerName pc1,pc2 -Force
    List Windows Version on multiple remote computer and disable the built-in Format-Table and Sort-Object.
.EXAMPLE
    Get-WindowsVersion -History with built-in Format-Table and Sort-Object.
    List History Windows Version on local computer.
.EXAMPLE
    Get-WindowsVersion -ComputerName pc1,pc2 -History
    List History Windows Version on multiple remote computer with built-in Format-Table and Sort-Object.
.EXAMPLE
    Get-WindowsVersion -ComputerName pc1,pc2 -History -Force
    List History Windows Version on multiple remote computer and disable built-in Format-Table and Sort-Object.
#>
    [cmdletbinding()]
    param (
        [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$ComputerName = "localhost",
        [string]$SearchBase,
        [switch]$History,
        [switch]$Force
    )
    if ($($PsVersionTable.PSVersion.Major) -gt "2") {
        # SearchBase
        if ($SearchBase) {
            if (Get-Command Get-AD* -ErrorAction SilentlyContinue) {
                if (Get-ADOrganizationalUnit -Filter "distinguishedName -eq '$SearchBase'" -ErrorAction SilentlyContinue) {
                    $Table = Get-ADComputer -SearchBase $SearchBase -Filter *
                    $ComputerName = $Table.Name
                }
                else {Write-Warning "No SearchBase found"}
            }
            else {Write-Warning "No AD Cmdlet found"}
        }
        # Loop 1
        $Tmp = New-TemporaryFile
        foreach ($Computer in $ComputerName) {
            if (Test-Connection -ComputerName $Computer -Count 1 -ErrorAction SilentlyContinue) {
                try {
                    $WmiObj = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer
                }
                catch {
                    Write-Warning "$Computer no wmi access"
                }
                if ($WmiObj) {
                    # Variables
                    $WmiClass = [WmiClass]"\\$Computer\root\default:stdRegProv"
                    $HKLM = 2147483650
                    $Reg1 = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
                    $Reg2 = "SYSTEM\Setup"
                    if ($History) {$KeyArr = ($WmiClass.EnumKey($HKLM, $Reg2)).snames -like "Source*"} else {$KeyArr = $Reg1}
                    # Loop 2
                    foreach ($Key in $KeyArr) {
                        if ($History) {$Reg = "$Reg2\$Key"} else {$Reg = $Key}
                        $Major = $WmiClass.GetDWordValue($HKLM, $Reg, "CurrentMajorVersionNumber").UValue
                        $Minor = $WmiClass.GetDWordValue($HKLM, $Reg, "CurrentMinorVersionNumber").UValue
                        $Build = $WmiClass.GetStringValue($HKLM, $Reg, "CurrentBuildNumber").sValue
                        $UBR = $WmiClass.GetDWordValue($HKLM, $Reg, "UBR").UValue
                        $ReleaseId = $WmiClass.GetStringValue($HKLM, $Reg, "ReleaseId").sValue
                        $ProductName = $WmiClass.GetStringValue($HKLM, $Reg, "ProductName").sValue
                        $ProductId = $WmiClass.GetStringValue($HKLM, $Reg, "ProductId").sValue
                        $InstallTime1 = $WmiClass.GetQWordValue($HKLM, $Reg, "InstallTime").UValue
                        $InstallTime2 = ([datetime]::FromFileTime($InstallTime1))
                        # Variables Windows 6.x
                        if ($Major.Length -le 0) {$Major = $WmiClass.GetStringValue($HKLM, $Reg, "CurrentVersion").sValue}
                        if ($ReleaseId.Length -le 0) {$ReleaseId = $WmiClass.GetStringValue($HKLM, $Reg, "CSDVersion").sValue}
                        if ($InstallTime1.Length -le 0) {$InstallTime2 = ([WMI]"").ConvertToDateTime($WmiObj.InstallDate)}
                        # Add Points
                        if (-not($Major.Length -le 0)) {$Major = "$Major."}
                        if (-not($Minor.Length -le 0)) {$Minor = "$Minor."}
                        if (-not($UBR.Length -le 0)) {$UBR = ".$UBR"}
                        # Output
                        $Output = New-Object -TypeName PSobject
                        $Output | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.toUpper()
                        $Output | Add-Member -MemberType NoteProperty -Name ProductName -Value $ProductName
                        $Output | Add-Member -MemberType NoteProperty -Name WindowsVersion -Value $ReleaseId
                        $Output | Add-Member -MemberType NoteProperty -Name WindowsBuild -Value "$Major$Minor$Build$UBR"
                        $Output | Add-Member -MemberType NoteProperty -Name ProductId -Value $ProductId
                        $Output | Add-Member -MemberType NoteProperty -Name InstallTime -Value $InstallTime2
                        $Output | Export-Csv -Path $Tmp -Append
                    }
                }
            }
            else {Write-Warning "$Computer not reachable"}
        }
        # Output
        if ($Force) {Import-Csv -Path $Tmp} else {Import-Csv -Path $Tmp | Sort-Object -Property ComputerName, WindowsVersion | Format-Table -AutoSize}
    }
    else {Write-Warning "PSVersion to low"}
}