param (
    $getDate = (Get-Date -Format "dd-MM-yyyy"),
    $localPath = "C:\Users\napol\Downloads\Practise\",
    $remotePath = "/Files/OPUS/",
    $failedLog = "C:\Users\napol\OPUS\Failed\Outbound\${getDate}.txt",
    $successLog = "C:\Users\napol\OPUS\Successful\Outbound\${getDate}.txt"
)

if (-not (Test-Path $failedLog)) {
    New-Item -ItemType File -Force -Path $failedLog | Out-Null
}
if (-not (Test-Path $successLog)) {
    New-Item -ItemType File -Force -Path $successLog | Out-Null
}
try {
    $smtpServer = "smtp.gmail.com"
    $smtpPort = "587"
    $smtpUser = (Get-Content -Path "C:\Users\napol\OneDrive\Desktop\Powershell\config.txt" -First 1)
    $smtpPwd = (Get-Content -Path "C:\Users\napol\OneDrive\Desktop\Powershell\config.txt" -Tail 1)
    $smtpCredentials = New-Object Management.Automation.PSCredential $smtpUser, ($smtpPwd | ConvertTo-SecureString -AsPlainText -Force)
}
catch {
    Write-Warning "Error: $($_.Exception.Message)" | Out-File -Append -FilePath $failedLog
} 
$subject = "`Daily Report " + (Get-Date -Format "dd/MM/yyyy")
$message1 = ""
$message2 = ""
$message3 = ""
$errorMessage1 = ""
$errorMessage2 = ""

try {
    # Add the WinSCP .NET assembly
    Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"

    # Create session options and configure them
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = "192.168.64.3"
        UserName = "testuser"
        Password = (Get-Content -Path "C:\Users\napol\OneDrive\Desktop\Powershell\crd.txt" -First 1)
        SshHostKeyFingerprint = (Get-Content -Path "C:\Users\napol\OneDrive\Desktop\Powershell\crd.txt" -Tail 1)
    }

    # Debug: Output session options to ensure they are correct
    Write-Output "Session Options: $sessionOptions" | Out-File -Append -FilePath $failedLog

    # Create a new WinSCP session object
    $session = New-Object WinSCP.Session

    try {
        # Open the session with the provided session options
        $session.Open($sessionOptions)

        # Check if the local directory exists
        if (-not (Test-Path $localPath)) {
            Write-Output "Local path does not exist. Creating directory: $localPath" | Out-File -Append -FilePath $failedLog
            New-Item -ItemType Directory -Force -Path $localPath
        }

        # Set up transfer options
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

        # Correct folder listing
        $directory = $session.ListDirectory("/Files/OPUS")
        $pdfFiles = $directory.Files | Where-Object { $_.Name -like "*.pdf" }

        
        if ($pdfFiles.Count -eq 0){
            Write-Output "No files available for processing" | Out-File -Append -FilePath $successLog
            $message1 = "No files available for processing" 
        }
        else{
            foreach ($file in $pdfFiles) {

            if ($file.Name -like "*.pdf") {

                $sourcePath = "/Files/OPUS/$($file.Name)"
                $destinationPath = Join-Path $localPath $file.Name

                # Download the file
                $session.GetFiles($sourcePath, $destinationPath, $false, $transferOptions).Check()

                Write-Output "Download of $($file.Name) succeeded" | Out-File -Append -FilePath $successLog
                $message2 += "Download of $($file.Name) -> $destinationPath succeeded `n"

                $session.RemoveFiles($sourcePath).Check()

                Write-Output "$($file.Name) has been deleted" | Out-File -Append -FilePath $successLog
                $message3 += "$($file.Name) has been deleted `n"
            }
        }
    }}
    catch{
        Write-Output "Error: $($_.Exception.Message)" | Out-File -Append -FilePath $failedLog
        $errorMessage1 = "Error: $($_.Exception.Message)"
        exit 0
    }
    finally {
        # Close the session after the transfer is done
        if ($session.Opened) {
            $session.Close()
        }
        $session.Dispose()

    }
}
catch {
    # Catch and print errors
    Write-Output "Error: $($_.Exception.Message)" | Out-File -Append -FilePath $failedLog
    $errorMessage2 = "Error: $($_.Exception.Message)"
    exit 1
}finally{
    $messages = @("$message1", "$message2", "$message3", "$errorMessage1", "$errorMessage2")
    $body = $messages -join "`n"
    Send-MailMessage -From "learningandcertifications@gmail.com" -To "learningandcertifications@gmail.com" -Subject $subject -Body $body -SmtpServer $smtpServer -Port $smtpPort -Credential $smtpCredentials -UseSsl
}