$productName = "Veeam Backup and Replication PowerShell Toolkit"
$host.ui.RawUI.WindowTitle = "$productName"

$currentExecutingPath = $MyInvocation.MyCommand.Definition.Replace($MyInvocation.MyCommand.Name, "Veeam.Backup.PowerShell.dll")
$xmlPath = $MyInvocation.MyCommand.Definition.Replace($MyInvocation.MyCommand.Name, "Veeam.Backup.Powershell.Cmdlets.xml")
try
{
	[xml]$xml = Get-Content $xmlPath
}
catch
{

}

cd $MyInvocation.MyCommand.Definition.Replace($MyInvocation.MyCommand.Name, "")

if ((Get-WmiObject -Class Win32_Processor | Select-Object AddressWidth) -match "32")
{
	set-alias installutil $env:windir\Microsoft.NET\Framework\v4.0.30319\installutil 
}
else
{
	$arch = [reflection.assemblyname]::GetAssemblyName($currentExecutingPath).ProcessorArchitecture
	if ($arch -match "MSIL")
	{
		set-alias installutil $env:windir\Microsoft.NET\Framework\v4.0.30319\installutil 
	}
	else
	{
		set-alias installutil $env:windir\Microsoft.NET\Framework64\v4.0.30319\installutil 
	}
}

$installed = Get-PSSnapin -Registered | ? { $_.Name -eq "VeeamPSSnapIn"}
if (-not $installed)
{
	installutil $currentExecutingPath | out-null
}


[void][Reflection.Assembly]::LoadWithPartialName("VimService2")
[void][Reflection.Assembly]::LoadWithPartialName("VimService2.XmlSerializers")
[void][Reflection.Assembly]::LoadWithPartialName("VimService")
[void][Reflection.Assembly]::LoadWithPartialName("VimService.XmlSerializers")
[void][Reflection.Assembly]::LoadWithPartialName("VimServiceInt")
[void][Reflection.Assembly]::LoadWithPartialName("VimServiceInt.XmlSerializers")

Set-Alias -Name Add-VBRHPStorage -Value Add-HP4Storage -Scope Global -Description "Veeam VeeamPSSnapIn alias"
Set-Alias -Name Set-VBRHPStorage -Value Set-HP4Storage -Scope Global -Description "Veeam VeeamPSSnapIn alias"
Set-Alias -Name Add-VBRHPSnapshot -Value Add-HP4Snapshot -Scope Global -Description "Veeam VeeamPSSnapIn alias"
Set-Alias -Name Get-VBRHPSnapshot -Value Get-HP4Snapshot -Scope Global -Description "Veeam VeeamPSSnapIn alias"
Set-Alias -Name Get-VBRHPCluster -Value Get-HP4Cluster -Scope Global -Description "Veeam VeeamPSSnapIn alias"
Set-Alias -Name Get-VBRHPStorage -Value Get-HP4Storage -Scope Global -Description "Veeam VeeamPSSnapIn alias"
Set-Alias -Name Get-VBRHPVolume -Value Get-HP4Volume -Scope Global -Description "Veeam VeeamPSSnapIn alias"
Set-Alias -Name Remove-VBRHPSnapshot -Value Remove-HP4Snapshot -Scope Global -Description "Veeam VeeamPSSnapIn alias"
Set-Alias -Name Remove-VBRHPStorage -Value Remove-HP4Storage -Scope Global -Description "Veeam VeeamPSSnapIn alias"
Set-Alias -Name Sync-VBRHPStorage -Value Sync-HP4Storage -Scope Global -Description "Veeam VeeamPSSnapIn alias"
Set-Alias -Name Sync-VBRHPVolume -Value Sync-HP4Volume -Scope Global -Description "Veeam VeeamPSSnapIn alias"
Set-Alias -Name Clone-VBRJob -Value Copy-VBRJob -Scope Global -Description "Veeam VeeamPSSnapIn alias"


function Get-VBRCommand()
{

<#
.SYNOPSIS
Returns Veeam PowerShell cmdlets.

.DESCRIPTION

This cmdlet returns the list of available Veeam PowerShell cmdlets.
If you run this cmdlet without parameters, it will return the list of all Veeam cmdlets in the current session.

.PARAMETER Name
Specifies the name of the command. Lists the commands that match the specified name or regular name patterns. Accepts wildcard characters.

.PARAMETER Noun
Specifies the noun. Lists the commands using the specified noun name. Accepts wildcard characters.

.PARAMETER Verb
Specifies the command verb. Lists the commands using the specified verb name. Accepts wildcard characters.

.EXAMPLE
Get-VBRCommand Remove*

.EXAMPLE
Get-VBRCommand -Name *Zip*

.EXAMPLE
Get-VBRCommand -Verb Get, Set

.EXAMPLE
Get-VBRCommand -Noun *Job*, *Zip*

#>

	[CmdletBinding(DefaultParameterSetName = 'AllCommandSet')]

	param
	(
		[parameter(
			Mandatory = $false,
			Position = 0,
			ParameterSetName = 'NameSet',
			ValueFromPipeline = $true,
			ValueFromPipelineByPropertyName = $true,
			HelpMessage = "Specifies the name of the command. Lists the commands that match the specified name or regular name patterns. Accepts wildcard characters.")]
		$Name,

		[parameter(
			Mandatory = $false,
			ParameterSetName = 'NounVerbSet',
			ValueFromPipelineByPropertyName = $true,
			HelpMessage = "Specifies the noun. Lists the commands using the specified noun name. Accepts wildcard characters.")]
		$Noun,

		[parameter(
			Mandatory = $false,
			ParameterSetName = 'NounVerbSet',
			ValueFromPipelineByPropertyName = $true,
			HelpMessage = "Specifies the command verb. Lists the commands using the specified verb name. Accepts wildcard characters.")]
		$Verb,

		[Parameter(ParameterSetName='AllCommandSet')]
		[switch]$V95,
		
		[Parameter(ParameterSetName='AllCommandSet')]
		[switch]$V90,
		
		[Parameter(ParameterSetName='AllCommandSet')]
		[switch]$V80
	)

	process
	{
		if (-not $PsCmdlet.MyInvocation.BoundParameters.ContainsKey("V95") -and
			-not $PsCmdlet.MyInvocation.BoundParameters.ContainsKey("V90") -and
			-not $PsCmdlet.MyInvocation.BoundParameters.ContainsKey("V80"))
		{
			$V95 = $true
			$V90 = $true
			$V80 = $true
		}

		if (-not $xml)
		{
			throw "Failed to load commandlet data";
		}

		switch ($PsCmdlet.ParameterSetName)
		{
			'AllCommandSet'
			{
				$cmdlets = get-command -pssnapin "VeeamPSSnapIn" -type Cmdlet
			}
	
			'NameSet'
			{
				$cmdlets = get-command -pssnapin "VeeamPSSnapIn" -type Cmdlet -name $Name
			}

			'NounVerbSet'
			{
				$cmdlets = get-command -pssnapin "VeeamPSSnapIn" -noun $Noun -verb $Verb
			}
		}

		$info = New-Object System.Collections.Specialized.OrderedDictionary
		$cmdlets | Sort-Object { $_.Name } | % { $info.Add(($_.Name.ToLower() -replace "-", ""), $_) }

		$result = @{}
		$map = @{ "V90" = $V90; "V95" = $V95; "V80" = $V80 }
	
		foreach ($key in $map.Keys)
		{
			$cmdlets = $xml.Root.Release | ? { $_.Version -eq $key.ToString() } | select Cmdlet
			foreach ($cmdlet in [array]$cmdlets.Cmdlet)
			{
				$result.Add($cmdlet.Name.ToLower(), @{ IsHidden=[bool]::Parse($cmdlet.IsHidden); IsRequested=$map[$key] })
			}
		}
	
		foreach ($key in $info.Keys)
		{
			if (!$result.Contains($key))
			{
				Write-Error $key;
			}
			if ($result[$key].IsRequested -and -not $result[$key].IsHidden)
			{
				Write-Output $info[$key]
			}
		}
	}
}

Add-PSSnapin VeeamPSSnapIn

$snapIn = Get-PSSnapIn "VeeamPSSnapIn"
$xmlFilePath = [System.IO.Path]::Combine($snapIn.ApplicationBase, "Veeam.Backup.PowerShell.format.ps1xml")
update-formatdata -prependPath $xmlFilePath

cd $env:userprofile

write-host "          Welcome to the $productName!"
write-host ""
write-host "To list available commands, type " -NoNewLine
write-host "Get-VBRCommand" -foregroundcolor yellow
write-host "To open online documentation on all available commands, type " -NoNewLine
write-host "Get-VBRToolkitDocumentation" -foregroundcolor yellow  
write-host ""
write-host "       Copyright © Veeam Software AG. All rights reserved."
write-host ""
write-host ""

function Get-VBRToolkitDocumentation
{
	[System.Diagnostics.Process]::Start("https://helpcenter.veeam.com/docs/backup/powershell/")
}
