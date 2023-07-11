param (
    [Parameter(Mandatory=$true)]
    [string]$UserPrincipalName
)

function Write-Log($message, $foregroundColor = 'Green') {
    $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $colorMessage = "$date - $message"
    Microsoft.PowerShell.Utility\Write-Host -ForegroundColor $foregroundColor $colorMessage
}


function Draw-Logo {
    Clear-Host
    Write-Host @"
    @                                                             
    @   @@  @                                                        
     @@* @@ .@                                                       
   @   @@@ @@ @                                                      
    @@@@@@@@@&@@&@                                                 
    @@@@@@@&@@@@@%@                                                  
    ,@&@@@&@@@&@@&@@@&                                               
      @@@@&@@@@@@@@@@&%@                                             
      @@@@&@@@@@&@@@&%@@@@                         @@ @              
       @@@@@@@@@@@&@&&(@@@@@                    @@@@@@@              
        @@@&@@@@@@@@@@#%@@@@@@               @@ @@@@@@@@             
          @&@&@@@@@@@@%@&@@.#@@,           @@@@@@@@@@@@              
           .@@@@@@@@@@@@@@@@@&@@@       @&&@@@@@@@@@@                
            #*@@@@@@@@@@&@@@@@@/&@@  @%&&@%@@@@@@@@@                 
              @@@@@@@@@@@@@@@@@@@&&@@%*(((((*(&                      
                @@@@@@@@@@@@@&@&%&%#%(,,*,../*@@@                    
                   @@@@@@&&&&&&&&%#%%%*//%#&&@@@%@                   
                        &&&&&&%%&&&&@%@@@#       @*                  
                     &&&&&&&&%&@@@@@@@.                              
                @@@%&&@@@&@@&@@@@@@                                  
       @@@@@@@@@@@@#@@@@@&#@@@@@@                                    
    @@@@@@@@@@@@@@@@@@@@@@@.&&@                                      
      *@@@@@@@@@@@@@@@@@@@. &@@@@                                    
                &@,@@@ ,#@ ,     %  

"@
    Write-Log "Raven v0.4"
}


function Prompt-User {
    try {
        $startDaysAgo = [int](Read-Host "Please enter the start date (as a number of days ago between 0 and 90, defaults to 90)")
    }
    catch {
        Write-Log "Invalid input. Please enter a valid number." 'Red'
        exit
    }

    $directory = Read-Host "Please enter the directory to save the audit logs"

    if ([string]::IsNullOrWhiteSpace($startDaysAgo)) { $startDaysAgo = 90 }

    if ($startDaysAgo -lt 0 -or $startDaysAgo -gt 90) {
        Write-Log "Start date should be between 0 and 90 days ago" 'Red'
        exit
    }

    $outputFile = Join-Path -Path $directory -ChildPath "AuditLogRecords.csv"
    return $startDaysAgo, $outputFile
}


function Set-Dates($startDaysAgo) {
    [DateTime]$start = [DateTime]::UtcNow.AddDays(-$startDaysAgo)
    [DateTime]$end = Get-Date
    $end = $end.AddMinutes(-$end.Minute).AddSeconds(-$end.Second).AddMilliseconds(-$end.Millisecond)
    return $start, $end
}

function Connect-Exchange {
    if (!(Get-ConnectionInformation | Where-Object { $_.Name -match 'ExchangeOnline' -and $_.state -eq 'Connected' })) { 
        Connect-ExchangeOnline -ShowBanner:$false
    }
    else {
        Write-Log "Already connected to Exchange Online"
    }

    $config = Get-AdminAuditLogConfig
    if ($null -eq $config -or !$config.UnifiedAuditLogIngestionEnabled) {
        Write-Log "Audit logging not enabled on tenant" 'Red'
        exit
    }
}


function Retrieve-AuditRecords($start, $end, $outputFile) {
    $resultSize = 5000
    $intervalMinutes = 60
    $totalIntervals = [Math]::Ceiling(($end - $start).TotalMinutes / $intervalMinutes)
    $intervalCount = 0
    [DateTime]$currentStart = $start
    [DateTime]$currentEnd = $end
    $totalCount = 0

    while ($currentStart -lt $end) {
        $currentEnd = $currentStart.AddMinutes($intervalMinutes)
        if ($currentEnd -gt $end) {
            $currentEnd = $end
        }

        $sessionID = [Guid]::NewGuid().ToString() + "_" + "ExtractLogs" + (Get-Date).ToString("yyyyMMddHHmmssfff")

        do {
            $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -UserIds $UserPrincipalName

            if (($results | Measure-Object).Count -ne 0) {
                $results | export-csv -Path $outputFile -Append -NoTypeInformation
                $totalCount += $results.Count
            }
        }
        while (($results | Measure-Object).Count -ne 0)

        $currentStart = $currentEnd

        $intervalCount++
        $progress = [math]::Round((($intervalCount / $totalIntervals) * 100), 2)
        Write-Progress -Activity "Retrieving audit records" -Status "$progress% Complete:" -PercentComplete $progress
    }

    if ($totalCount -eq 0) {
        Write-Log "No audit logs were found in the specified date range." 'Red'
    }
    else {
        Write-Log "Audit logs retrieval completed. Total logs found: $totalCount"
    }
}

# Main script
$startDaysAgo, $outputFile = Prompt-User
$start, $end = Set-Dates $startDaysAgo
Draw-Logo
Connect-Exchange
Retrieve-AuditRecords $start $end $outputFile
