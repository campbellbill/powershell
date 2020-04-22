$computers = Read-Host “What is the Computer Name?” #Enter the name of the computer you would like to modify
$userPW = Read-Host “What is the Password you would like to set?” #Enter the password you would like to set for the Administrator account.
$CurrentAdmin = Read-Host “What is the Current Administrator Name?” #Enter the name of the current administrator account.
$DisableDefaultAdminAccount = Read-Host “If you like to Enable the Default Administrator Account enter 0. If you would like to DISABLE the account enter 2” #Enter the status you would like the Administrator account to have. Enabled or Disabled.

foreach ($computer in $computers)
{
    ## This doesn’t need to be a function, I left it like this as it doesn’t hurt anything and if I Wanted to come back and actually create a LIST of computers I could.
    if (test-connection -computername $computer -quiet)
    {
        try
        {
            $localAdmin = [ADSI](“WinNT://” + $computer + “/” + $CurrentAdmin + “,User”)
            if($DisableDefaultAdminAccount -eq ‘0’)
            {
                $LocalAdmin.UserFlags = 65536 # UserFlags Value for the account to be active with a password set to never expire.
                $localAdmin.CommitChanges() # Commit the change
            }
            else
            {
                $LocalAdmin.UserFlags = 66083 #UserFlags Value for the account to be Disabled with password set to never expire.
                $localAdmin.CommitChanges() # Commit the change
            }

            $localAdmin.psbase.rename(‘SuperAdmin’)
            $localAdmin.setpassword($userPW)
            Write-Host “Successfully Renamed Administrator Account on $computer” -fore green
            $ObjComputer = [ADSI](“WinNT://” + $Computer)
            $DummyUser = $OBJComputer.Create(“User”, “Administrator”)
            $DummyUser.setPassword(“P@ssword1”)
            $DummyUser.SetInfo() #Commit this change of a new account with this password to the SAM DB – this makes the account visable and actable upon
            $DummyUser.Description = “Dummy Account” #Update the description of the account once commited to SAM
            $DummyUser.UserFlags = 66083
            $DummyUser.CommitChanges() # Commit the change of disabled and the description.
            Write-Host "Successfully Created Administrator Account on $computer" -fore green
        }
        catch {
        Write-Host “$_” -fore red
        }
    }
    else
    {
        Write-Host “Ping Failed to” $computer
    }
}
