Import-Module ActiveDirectory

$ResetPassword = Import-Csv "C:\Users\a.olumoh\Desktop\Scripts\Users.txt"
$CompletedLog = "C:\Users\a.olumoh\Desktop\Scripts\result.txt"
$ErrorLog = "C:\Users\a.olumoh\Desktop\Scripts\errorlog.txt"


#Password
$NewPassword = ConvertTo-SecureString "Welcome25" -AsPlainText -Force

#Ensures log files are cleared
Clear-Content -Path $CompletedLog -ErrorAction SilentlyContinue
Clear-Content -Path $ErrorLog -ErrorAction SilentlyContinue

#Loop through users in the txt file to reset password with error handling
foreach ($User in $ResetPassword) {
    try {
        Set-ADAccountPassword -Identity $User.ADUsername -Reset -NewPassword $NewPassword -ErrorAction Stop
        Set-ADUser -Identity $User.ADUsername -ChangePasswordAtLogon $true -ErrorAction Stop
        "Username: $($User.ADUsername) `| Password: Welcome25" | Out-File -Append -FilePath $CompletedLog
    }
    catch {
        "$($User.ADUsername) - Failed: $($_.Exception.Message)" | Out-File -Append -FilePath $ErrorLog
    }
}
Get-Date | Out-File -Append -FilePath $CompletedLog
Write-Host "Script completed."
Start-Sleep -Seconds 5
