<#
The sample scripts are not supported under any Microsoft standard support 
program or service. The sample scripts are provided AS IS without warranty  
of any kind. Microsoft further disclaims all implied warranties including,  
without limitation, any implied warranties of merchantability or of fitness for 
a particular purpose. The entire risk arising out of the use or performance of  
the sample scripts and documentation remains with you. In no event shall 
Microsoft, its authors, or anyone else involved in the creation, production, or 
delivery of the scripts be liable for any damages whatsoever (including, 
without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use 
of or inability to use the sample scripts or documentation, even if Microsoft 
has been advised of the possibility of such damages.

read-host -assecurestring | convertfrom-securestring | out-file C:\cred.txt

#> 

#requires -Version 2

$password = get-content C:\cred.txt | convertto-securestring
$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist "xth",$password




Function Get-OSCServiceList
{
<#
	.SYNOPSIS
	Function Get-OSCServiceList is an advanced function which can list all services that using a particular user account as start account.
	.DESCRIPTION
	Function Get-OSCServiceList is an advanced function which can list all services that using a particular user account as start account.
	.PARAMETER $FilePath
	Get the path of specified CSV file
	.PARAMETER $ComputerName
	Get the specified computer 
	.PARAMETER $UserName
	Get the specified UserName 
	.EXAMPLE
	Get-OSCServiceList -FilePath "C:\New Text Document.csv"  -UserName "ADMINISTRATOR"
	
	List all services that using "administrator" as start account in computer(s) in the CSV file.
	.EXAMPLE
	Get-OSCServiceList -Computer "MININT-I5DE0FO" -UserName "ADMINISTRATOR"
	
	List all services that using "administrator" as start account in "MININT-I5DE0FO".	
#>
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$True,Position=0,ParameterSetName="FilePath")]
		[String]$FilePath,
		[Parameter(Mandatory=$False,Position=0,ParameterSetName="ComputerName")]
		[String]$ComputerName=$Env:COMPUTERNAME,
		[Parameter(Mandatory=$False,Position=1)]
		[String]$UserName=$Env:USERNAME
	)
	#Verify the type of input
	Switch($pscmdlet.ParameterSetName)
	{
		"FilePath"
		{	
			$Computers = Import-Csv  -Path $FilePath
			foreach ($computer in $Computers)
			{	
				$ComputerName = $computer.ComputerName
				CheckServicesOnComputer $ComputerName $UserName
			}
		}
		"ComputerName"
		{
			CheckServicesOnComputer $ComputerName $UserName		 
		}
	}
}

Function Set-OSCServicePSW
{	
<#
	.SYNOPSIS
	Function Set-OSCServicePSW is an advanced function which can set the new password for the specified service.
	.DESCRIPTION
	Function Set-OSCServicePSW is an advanced function which can set the new password for the specified service.
	.PARAMETER $ComputerName
	Get the specified computer 
	.PARAMETER $ServiceName
	Get the specified service
	.PARAMETER $UserName
	Get the specified user
	.PARAMETER $NewPSW
	Set the new password 
	
	.EXAMPLE
	Set-OSCServicePSW -ComputerName "MININT-I5DE0FO" -ServiceName "Spooler" -UserName "administrator" -NewPSW 12345678
	
	Set the new password "12345678" for "administrator" to the service "Spooler"
#>	
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$True,Position=0)]
		[String]$ComputerName,
		[Parameter(Mandatory=$True,Position=1)]
		[String]$ServiceName,
		[Parameter(Mandatory=$True,Position=2)]
		[String]$UserName,
		[Parameter(Mandatory=$True,Position=3)]
		[String]$NewPassWord
	)
	If(Test-Connection -ComputerName $ComputerName)
	{
		#Get the specified service
		$Service = Get-WmiObject win32_service -ComputerName $ComputerName -property name, startname, caption | 
			Where-Object { $_.Name -match $ServiceName}
		#Chech for the service
		If($Service -eq $null)
		{
			Write-Warning "The $ServiceName does not exist."
		}
		Else
		{	
			#Get the result of changing the password
			$Result = $Service.Change($null,$null,$null,$null,$null,$null, $UserName, $NewPassWord)
			#If the return value equals 0,success to set the password
			If($Result.ReturnValue -eq 0)
			{
				Write-Host "Set new password successfully."
			}
			#If the return value does not equal 0 ,fail to set the password
			Else
			{	
				Write-Host "Fail to set new password,please ensure the username valid and you have the permission."
			}
		}	
	}
	Else 
	{	
		Write-Warning "Can not connect to the computer $ComputerName,please close the firewall or make it online"
	}
}

Function CheckServicesOnComputer($strComputer,$UserName)
{
	#This function is used to list the service started up with specified account
	If(Test-Connection -ComputerName $strComputer -ErrorAction SilentlyContinue)
	{
		#Get services
	    $Services = Get-WmiObject win32_service -ComputerName $strComputer -property name, startname, caption | 
		Where-Object { $_.startname -match $UserName}  
		#If the $services is empty
		If($Services -eq $null)
		{
			Write-Warning "Some error occur in $strComputer  or there is no service start up with $UserName in $strComputer"
		}
		Else
		{	
			#List services
			$Services | Select-Object -Property @{Name="Name";Expression={$_.Name}},`
												@{Name="StartWithAccount";Expression={$_.StartName}},`
												@{Name="Caption";Expression={$_.Caption}},`
												@{Name="Computer";Expression={$strComputer}}
		}
	}
	Else 
	{
		Write-Warning "Can not connect to the computer $strComputer,please close the firewall or make it online"
	}
}
    