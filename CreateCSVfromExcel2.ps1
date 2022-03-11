Function ExportWSToCSV ($excelFileName, $csvLoc)
{
    $excelFile = $csvLoc + $excelFileName + ".xlsx"
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false

    $wb = $E.Workbooks.Open($excelFile)

    foreach ($ws in $wb.Worksheets)    
    {
        if  ( $ws.name -match 'EXT_'  ){
    
            $n = $ws.Name
            $ws.SaveAs($csvLoc + $n + ".csv", 6)
        }
    }
    $E.Quit()
}
ExportWSToCSV -excelFileName "results-for-mres-postgrad-2021-08-31-0841" -csvLoc "C:\Users\southgaa\Documents\"
# Load WinSCP .NET assembly
Add-Type -Path "C:\Users\southgaa\WinSCPnet.dll"

New-WinSCPSession -Hostname "t-devbi-db-01.iso.port.ac.uk" -Credential (Get-Credential) -SshHostKeyFingerprint "ssh-ed25519 255 Cn2I73xHqfsNMFF/5yCmp2m7oliTdjlYJBX9d7/73x4=" | 
     Receive-WinSCPItem -Path "C:\Users\southgaa\Documents\*.csv" -Destination "/data/oracle/extern/pmr_files/" -ExecutablePath = "C:\Users\southgaa\AppData\Local\Programs\WinSCP\WinSCP.exe"

