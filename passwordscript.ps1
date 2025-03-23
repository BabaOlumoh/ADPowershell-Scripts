﻿Import-Module ActiveDirectory
Start-Process PowerShell -Credential "DOMAIN\a.olumoh" -NoNewWindow -ArgumentList "-File C:\scripts\reset_password.ps1"

$ResetPassword = Import-Csv "C:\Users\a.olumoh\Desktop\Scripts\Users.txt"
$CompletedLog = "C:\Users\a.olumoh\Desktop\Scripts\result.txt"
$ErrorLog = "C:\Users\a.olumoh\Desktop\Scripts\errorlog.txt"

#Password
$NewPassword = ConvertTo-SecureString "Welcome25" -AsPlainText -Force

#Ensure log files are cleared
Clear-Content -Path $CompletedLog -ErrorAction SilentlyContinue
Clear-Content -Path $ErrorLog -ErrorAction SilentlyContinue

#Loop through users in the txt file to reset password with error handling
foreach ($User in $ResetPassword) {
    try {
        Set-ADAccountPassword -Identity $User.ADUsername -Reset -NewPassword $NewPassword -ErrorAction Stop
        "Username: $($User.ADUsername) `| Password: Welcome25" | Out-File -Append -FilePath $CompletedLog
    }
    catch {
        "$($User.ADUsername) - Failed: $($_.Exception.Message)" | Out-File -Append -FilePath $ErrorLog
    }
}

Get-Date | Out-File -Append -FilePath $CompletedLog
Write-Host "Password reset process completed."
