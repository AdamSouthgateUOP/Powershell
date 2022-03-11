Function ExportWSToCSV ($excelFileName, $csvLoc) {
    $excelFile = $csvLoc + $excelFileName + '.xlsx'
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false

    $wb = $E.Workbooks.Open($excelFile)

    foreach ($ws in $wb.Worksheets) {
        if ( $ws.name -match 'EXT_'  ) {
    
            $n = $ws.Name
            $ws.SaveAs($csvLoc + $n + '.csv', 6)
        }
    }
    $E.Quit()
}
ExportWSToCSV -excelFileName 'results-for-mres-postgrad-2021-08-31-0841' -csvLoc 'C:\Users\southgaa\Documents\'
# Load WinSCP .NET assembly
Add-Type -Path 'C:\Users\southgaa\WinSCPnet.dll'

# Set up session options
$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
    Protocol              = [WinSCP.Protocol]::Sftp
    HostName              = 't-devbi-db-01.iso.port.ac.uk'
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
    $remotePath = '/data/oracle/extern/pmr_files/*'

    $transferResult = $session.PutFiles('C:\Users\southgaa\Documents\EXT_*.csv', $remotePath).Check()

    foreach ($transfer in $transferResult.Transfers) {
        Write-Host ('Upload of {0} succeeded' -f $transfer.FileName)
    }
    
    
}
finally {
    $session.Dispose()
}