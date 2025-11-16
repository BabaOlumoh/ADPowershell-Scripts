param (
    $getDate = (Get-Date -Format "dd-MM-yyyy"),
    $localPath = "C:\Users\napol\OPUS\Processing\",
    $remotePath = "/Files/ScannedDocs",
    $archive = "C:\Users\napol\OPUS\Processing\Archive\",
    $failedLog = "C:\Users\napol\OPUS\Failed\Inbound\${getDate}.txt",
    $successLog = "C:\Users\napol\OPUS\Successful\Inbound\${getDate}.txt"
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
$errorMessage3 = ""

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
    Write-Output "Session Options: $sessionOptions" | Out-File -Append -FilePath $successLog

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
        $ignoreFolders = @("Archive")
        $directories = Get-ChildItem -Path $localPath -Directory -Recurse

        foreach ($dir in $directories) {

            # Skip Archive
            if ($ignoreFolders -contains $dir.Name){
                continue
            }

            $files = Get-ChildItem -Path $dir.FullName -File

            foreach ($file in $files){
                if ($file.Extension -eq ".pdf" -or $file.Extension -eq ".csv"){

                    $localSource = $file.FullName
                    $remoteDestination = "$remotePath/$($file.Name)"
                    
                    Write-Output "Uploading file: $localSource --> $remoteDestination" | Out-File -Append -FilePath $successLog
                    $session.PutFiles($localSource, $remoteDestination, $false, $transferOptions).Check()

                    $message2 += "$($file.Name) has been uploaded successfully `n"
                    Write-Output "$($file.Name) has been uploaded successfully" | Out-File -Append -FilePath $successLog
                    
                }

                if ($file.Extension -eq ".pdf"){
                    Start-Sleep -Seconds 2
                }

                if ($file.Extension -eq ".csv"){
                    Start-Sleep -Seconds 10
                }
            }
        }

    }
    catch{
        Write-Output "Error: $($_.Exception.Message)" | Out-File -Append -FilePath $failedLog
        $errorMessage1 = "Error: $($_.Exception.Message)"
        exit 0
    }

    try {
        $ignoreFolders = @("Archive")

        foreach($item in Get-ChildItem -Path $localPath -Directory){

            if ($ignoreFolders -contains $item.Name){
                continue
            }

            $source = $item.FullName
            $destination = Join-Path $archive $item.Name

            Move-Item -Path $source -Destination $destination

            $message3 += "$($item.Name) has been moved successfully to $archive `n"
            Write-Output "$($item.Name) has been moved successfully to $archive" | Out-File -Append -FilePath $successLog

        }
    }
    catch {
        # Catch and print errors
        Write-Output "Error: $($_.Exception.Message)" | Out-File -Append -FilePath $failedLog
        $errorMessage2 = "Error: $($_.Exception.Message)"
        exit 2
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
    $errorMessage3 = "Error: $($_.Exception.Message)"
    exit 2
}finally{
    $messages = @("$message1", "$message2", "$message3", "$errorMessage1", "$errorMessage2")
    $body = $messages -join "`n"
    Send-MailMessage -From "learningandcertifications@gmail.com" -To "learningandcertifications@gmail.com" -Subject $subject -Body $body -SmtpServer $smtpServer -Port $smtpPort -Credential $smtpCredentials -UseSsl
}