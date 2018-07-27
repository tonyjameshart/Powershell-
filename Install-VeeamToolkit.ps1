$productName = "Veeam Backup PowerShell Toolkit"
$host.ui.RawUI.WindowTitle = "$productName"

##http://grimkey.wordpress.com/2011/03/14/loading-a-net-4-0-snap-in-in-powershell-v2/
function Write-DotNet4Config
{
<#
.SYNOPSIS
Function to write .Net 4 configuration files for me so I don't have to remember

.EXAMPLE
Write-DotNet4Config "$env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe"

Writes a configuration for the given executable if none currently exists.
#>

	[CmdletBinding()]
	param($executable, [switch]$whatif)
	
	if (-not (test-path $executable)) 
	{
		"Cannot find executable $executable" | out-null
		if (-not $whatif.isPresent) { return }
	}
	
	if (test-path "$executable.config") 
	{
		"Path already exists" | out-null
		if (-not $whatif.isPresent) { return }
	}
	
	$versions = @("v4.0.30319", "v2.0.50727" )
	
	$config = @"
<?xml version="1.0"?>
<configuration>
    <startup useLegacyV2RuntimeActivationPolicy="true">
        $($versions | %{'<supportedRuntime version="{0}"/>' -f $_})
    </startup>
</configuration>
"@

	"XML: $config" | out-null
	"Output file: $executable.config" | out-null
	
	if ($whatif.isPresent)
	{
		"What if: Writes to file $executable.config" | out-null
		"What if: Writes xml: $config" | out-null
	}
	else
	{
		([xml]$config).Save("$executable.config")
	}
}

if ((Get-WMIObject win32_operatingsystem).OSArchitecture -match "32")
{
	Write-DotNet4Config "$env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe"
}
else
{
	Write-DotNet4Config "$env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe"
	Write-DotNet4Config "$env:SystemRoot\syswow64\WindowsPowerShell\v1.0\powershell.exe"
}

$initialize = split-path -parent $MyInvocation.MyCommand.Definition | join-path -ChildPath "Initialize-VeeamToolkit.ps1"
. $initialize
