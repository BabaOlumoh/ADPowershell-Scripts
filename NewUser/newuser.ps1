param(
    [int]$Length = 5,             
    [string]$Prefix = "MIBTempPass" 
)

Import-Module ActiveDirectory


#File location
$NewUser = Import-Csv "C:\Users\a.olumoh\Desktop\Scripts\Users.txt"
$CompletedLog = "C:\Users\a.olumoh\Desktop\Scripts\result.txt"
$ErrorLog = "C:\Users\a.olumoh\Desktop\Scripts\errorlog.txt"


foreach ($User in $NewUser){
    try {
        #Password generator
        $Numbers = '0123456789'
        $Symbols = '!@#$%&*()-_[]{}?'
        $AllChars = $Numbers + $Symbols

        $RandomPart = -join ((1..$Length) | ForEach-Object { 
            $AllChars[(Get-Random -Maximum $AllChars.Length)] 
        })
        
        $FirstName = $User.firstName
        $LastName = $User.lastName
        $DisplayName = $User.displayName
        $UserName = $User.userName
        $Email = $User.email
        $UPN = $User.upn
        $Department = $User.department
        $Company = $User.company
        $OU = $User.ou
        $JoinChars = "$Prefix$RandomPart"
        $Password = ConvertTo-SecureString -AsPlainText $JoinChars -Force

        if (Get-ADUser -F {SamAccountName -eq $UserName})
        {
            Write-Warning "An account with username $UserName already exists" | Out-File -Append -FilePath $ErrorLog
        }
        else {
            New-ADUser `
                
                -GivenName $FirstName `
                -Surname $LastName `
                -DisplayName $DisplayName `
                -SamAccountName $UserName `
                -UserPrincipalName $UPN `
                -EmailAddress $Email `
                -Department $Department `
                -Company $Company `
                -Path $OU `
                -ChangePasswordAtLogon $False `
                -AccountPassword $Password
        }
        "Username: $($UserName) `| Password: $($Password)" | Out-File -Append -FilePath $CompletedLog
    }
    catch {
        "$($UserName) - Failed: $($_.Exception.Message)" | Out-File -Append -FilePath $ErrorLog
    }
    
}

Write-Host "Script completed."









