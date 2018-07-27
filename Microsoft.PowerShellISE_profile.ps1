# Admin Creds
# read-host -assecurestring | convertfrom-securestring | out-file C:\cred.txt
$pass = get-content C:\cred.txt | convertto-securestring
$creds = new-object -typename System.Management.Automation.PSCredential -argumentlist "xth",$pass

<#Module Browser Begin
#Version: 1.0.0
Add-Type -Path 'C:\Program Files (x86)\Microsoft Module Browser\ModuleBrowser.dll'
$moduleBrowser = $psISE.CurrentPowerShellTab.VerticalAddOnTools.Add('Module Browser', [ModuleBrowser.Views.MainView], $true)
$psISE.CurrentPowerShellTab.VisibleVerticalAddOnTools.SelectedAddOnTool = $moduleBrowser
#Module Browser End
 
# *** Variables *** 
New-Variable -Name ProfileFolder -Value (Split-Path $PROFILE -Parent)  
New-Variable -Name IseProfile -Value (Join-Path -Path (Split-Path $PROFILE -Parent ) -ChildPath Microsoft.PowerShellISE_profile.ps1 ` )



Set-Variable -Name MaximumHistoryCount -Value 128 
 
# *** Alias *** 
Get-Command -CommandType cmdlet |  
Foreach-Object {  
 Set-Alias -name ( $_.name -replace "-","") -value $_.name -description thart_Alias 
} #end Get-Command 
New-Alias -name gh -value Get-Help -description MrEd_Alias 
New-Alias -name i -value Invoke-History -description MrEd_Alias 
New-Alias -name p -value Get-Profile -description MrEd_Alias 
 
# *** PS Drive *** 
 
# *** Function *** 
Function Get-Profile 
{ 
 Notepad $profile 
} #end function get-profile 

####
# Modules
####

Get-Module -ListAvailable | Import-Module

<#

Import-Module ActiveDirectory
Import-Module MSOnline

#region: Load VMware Snapin or Module (if not already loaded)  
if (!(Get-Module -Name VMware.VimAutomation.Core) -and (Get-Module -ListAvailable -Name VMware.VimAutomation.Core)) {  
    Write-Output "loading the VMware COre Module..."  
    if (!(Import-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue)) {  
        # Error out if loading fails  
        Write-Error "`nERROR: Cannot load the VMware Module. Is the PowerCLI installed?"  
     }  
    $Loaded = $True  
    }  
    elseif (!(Get-PSSnapin -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue) -and !(Get-Module -Name VMware.VimAutomation.Core) -and ($Loaded -ne $True)) {  
        Write-Output "loading the VMware Core Snapin..."  
     if (!(Add-PSSnapin -PassThru VMware.VimAutomation.Core -ErrorAction SilentlyContinue)) {  
     # Error out if loading fails  
     Write-Error "`nERROR: Cannot load the VMware Snapin or Module. Is the PowerCLI installed?"  
     }  
    }  
#endregion  
#Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $true
Import-Module VMWare.PowerCLI
Import-Module Veeam.PowerCLI-Interactions
#>