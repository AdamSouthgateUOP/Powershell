#SFTP RDP file (adm_bi_12.rpd) from K:\Central\IS\BI Team\BIRepository_backups\1. DEV to data/oracle/rpd on t-devbi-app-01


# Load WinSCP .NET assembly
Add-Type -Path 'C:\Users\southgaa\WinSCPnet.dll'

# Set up session options
$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
    Protocol              = [WinSCP.Protocol]::Sftp
    HostName              = 't-devbi-app-01.iso.port.ac.uk'
    UserName              = 'southgaa'
    Password              = 'H4rdfl00r.de'
    SshHostKeyFingerprint = 'ssh-ed25519 255 Cn2I73xHqfsNMFF/5yCmp2m7oliTdjlYJBX9d7/73x4='    
}

$session = New-Object WinSCP.Session
$session.ExecutablePath = 'C:\Users\southgaa\AppData\Local\Programs\WinSCP\WinSCP.exe'


try {
    # Connect
    $session.Open($sessionOptions)

    # Transfer files
    $remotePath = '/data/oracle/rpd/'
    $localPath = "\\sp1\Public\Central\IS\BI Team\BIRepository_backups\1. DEV\adm_bi_12.rpd"

    $transferResult = $session.PutFiles($localPath, $remotePath).Check()

    foreach ($transfer in $transferResult.Transfers) {
        Write-Host ('Upload of {0} succeeded' -f $transfer.FileName)
    }
    
    
}
finally {
    $session.Dispose()
}
