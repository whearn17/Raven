param (
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName
)

function Write-LogMessage($message, $foregroundColor = 'Green') {
    $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $colorMessage = "$date - $message"
    Microsoft.PowerShell.Utility\Write-Host -ForegroundColor $foregroundColor $colorMessage
}


function Show-Logo {
    Clear-Host
    $logo = @"
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
$logo.ToCharArray() | ForEach-Object {
    Write-Host $_ -NoNewline
    Start-Sleep -Milliseconds .02
}
    Write-Host ""
    Write-LogMessage "Raven v1.0"
}


function Expand-Property($object, $parentName = $null) {
    $properties = @{}
    foreach ($property in $object.PSObject.Properties) {
        $key = if ($parentName) { "$parentName.$($property.Name)" } else { $property.Name }
        if ($property.Value -is [PSCustomObject]) {
            $properties += Expand-Property -object $property.Value -parentName $key
        } else {
            $properties[$key] = $property.Value
        }
    }
    return $properties
}


function Format-AuditData($inputFile) {
    # Read the CSV file
    $data = Import-Csv -Path $inputFile

    $outputFile = $inputFile + "ParsedAuditLogs.csv"

    # Initialize an empty array to hold the parsed data
    $parsedData = @()

    # Loop over each row in the data
    foreach ($row in $data) {
        # Create a new dictionary
        $rowDict = @{}

        # Add properties from the original row to the dictionary
        foreach ($property in $row.PSObject.Properties) {
            if ($property.Name -ne 'AuditData') {
                $rowDict[$property.Name] = $property.Value
            }
        }

        # Parse the JSON in the 'AuditData' column
        $auditData = ConvertFrom-Json -InputObject $row.AuditData

        # Remove the 'RecordType' property from $auditData
        $auditData.PSObject.Properties.Remove('RecordType')

        # Add properties from the parsed JSON to the dictionary
        $rowDict += Expand-Property -object $auditData

        # Convert the dictionary to a PSObject and add it to the parsed data
        $parsedData += New-Object PSObject -Property $rowDict
    }

    # Write the parsed data to a new CSV file
    $parsedData | Export-Csv -Path $outputFile -NoTypeInformation
}



function Read-UserInput {
    $userInput = Read-Host "Please enter the start date (as a number of days ago between 0 and 90, defaults to 90)"
    $startDaysAgo = $null
    if (![int32]::TryParse($userInput, [ref]$startDaysAgo)) {
        Write-LogMessage "Invalid input. Please enter a valid number." 'Red'
        exit
    }

    $directory = Read-Host "Please enter the directory to save the audit logs"
    
    if (!(Test-Path $directory -PathType Container)) {
        Write-LogMessage "Invalid directory. Please enter a valid directory." 'Red'
        exit
    }

    if ([string]::IsNullOrWhiteSpace($startDaysAgo)) { $startDaysAgo = 90 }

    if ($startDaysAgo -lt 0 -or $startDaysAgo -gt 90) {
        Write-LogMessage "Start date should be between 0 and 90 days ago" 'Red'
        exit
    }

    $outputFile = Join-Path -Path $directory -ChildPath "AuditLogRecords.csv"
    return $startDaysAgo, $outputFile
}



function Get-DateRange($startDaysAgo) {
    [DateTime]$start = [DateTime]::UtcNow.AddDays(-$startDaysAgo)
    [DateTime]$end = Get-Date
    $end = $end.AddMinutes(-$end.Minute).AddSeconds(-$end.Second).AddMilliseconds(-$end.Millisecond)
    return $start, $end
}

function Connect-ExchangeService {
    if (!(Get-ConnectionInformation | Where-Object { $_.Name -match 'ExchangeOnline' -and $_.state -eq 'Connected' })) { 
        Connect-ExchangeOnline -ShowBanner:$false
    }
    else {
        Write-LogMessage "Already connected to Exchange Online"
    }

    $config = Get-AdminAuditLogConfig
    if ($null -eq $config -or !$config.UnifiedAuditLogIngestionEnabled) {
        Write-LogMessage "Audit logging not enabled on tenant" 'Red'
        exit
    }
}


function Get-AuditRecords($start, $end, $outputFile) {
    $resultSize = 5000
    $intervalMinutes = 60
    $totalIntervals = [Math]::Ceiling(($end - $start).TotalMinutes / $intervalMinutes)

    [DateTime]$currentEnd = $end

    for ($intervalCount = 0; $intervalCount -lt $totalIntervals; $intervalCount++) {
        
        $currentStart, $currentEnd = Get-TimeInterval $start $end $intervalMinutes $intervalCount
        $sessionID = New-SessionID

        do {
            $results = Search-AuditLog $currentStart $currentEnd $sessionID $resultSize $UserPrincipalName $outputFile
        } while ($results)

        Show-Progress $intervalCount $totalIntervals
    }

    Write-LogMessage "Raven finished"
}


function Get-TimeInterval($start, $end, $intervalMinutes, $intervalCount) {
    [DateTime]$currentStart = $start.AddMinutes($intervalCount * $intervalMinutes)
    [DateTime]$currentEnd = $currentStart.AddMinutes($intervalMinutes)

    if ($currentEnd -gt $end) {
        $currentEnd = $end
    }

    return $currentStart, $currentEnd
}


function New-SessionID() {
    return [Guid]::NewGuid().ToString() + "_" + "ExtractLogs" + (Get-Date).ToString("yyyyMMddHHmmssfff")
}


function Search-AuditLog($currentStart, $currentEnd, $sessionID, $resultSize, $UserPrincipalName, $outputFile) {
    $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize -UserIds $UserPrincipalName

    if (($results | Measure-Object).Count -ne 0) {
        $results | Export-Csv -Path $outputFile -Append -NoTypeInformation
        return $results
    }
}


function Show-Progress($intervalCount, $totalIntervals) {
    $progress = [math]::Round((($intervalCount / $totalIntervals) * 100), 0)
    Write-Progress -Activity "Retrieving audit records" -Status "$progress% Complete:" -PercentComplete $progress
}


function main {
    $startDaysAgo, $outputFile = Read-UserInput
    $start, $end = Get-DateRange $startDaysAgo
    Show-Logo
    Connect-ExchangeService
    Get-AuditRecords $start $end $outputFile

    Format-AuditData -inputFile $outputFile
}

main