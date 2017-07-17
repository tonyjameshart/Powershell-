$strComputer=$env:computername
$users = @("username, username")
$password = "<password>" #hulu WALMART & golf TOKYO 6 RIVER ROPE

ForEach ($user in $users){
  write-host $user
  write-host $password

#Create the user account and assign a default password

$objOU = [adsi]"WinNT://."
$objUser = $objOU.Create("User", $user)
$objuser.setPassword($password)
$objuser.setinfo()

#Enable [User must change password at next logon]
$objuser.PasswordExpired = 0
$objuser.SetInfo()

#Add the User account to the local xm8-admins (Administrators Group)

$computer = [ADSI]("WinNT://" + $strComputer + ",computer") 
$group = $computer.psbase.children.find("Administrators")  
$group.Add("WinNT://" + $strComputer + "/" + $user)  


}