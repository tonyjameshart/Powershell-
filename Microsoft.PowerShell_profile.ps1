## Admin Credentials
$pass = get-content C:\cred.txt | convertto-securestring
$creds = new-object -typename System.Management.Automation.PSCredential -argumentlist "xth",$pass


# *** Variables *** 

 
# *** PS Drive *** 

# *** Function *** 
Function Get-Profile 
{ 
 Notepad $profile 
} #end function get-profile 

####
# Modules
####

Import-Module ActiveDirectory
Import-Module MSOnline



#Load VMware Snapin or Module (if not already loaded)  
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

#Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $true
