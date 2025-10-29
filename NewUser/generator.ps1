param(
    [int]$Length = 5,             
    [string]$Prefix = "MIBTempPass" 
)


$numbers = '0123456789'
$symbols = '!@#$%&*()-_[]{}?'


$allChars = $numbers + $symbols

$randomPart = -join ((1..$Length) | ForEach-Object { 
    $allChars[(Get-Random -Maximum $allChars.Length)] 
})

$password = "$Prefix$randomPart"

# Output password
Write-Host "Generated Password: $password"
