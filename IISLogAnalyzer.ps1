param (
    [string] $AllReportsListCSV,
    [string] $DirectoryOfLogFiles
)

$allLogs = gci $DirectoryOfLogFiles
$logQty = $allLogs.count
$allReports = Import-Csv $AllReportsListCSV -Delimiter ';'
$endResults = New-Object -TypeName "System.Collections.ArrayList"
$repQty = $allReports.count

class Report {
    [string]$ReportGuid
    [string]$ReportPdfPath
    [System.Collections.ArrayList]$LastAccessTimePerIP
}

For ($r = 0; $r -lt $repQty; $r++) {
    $q = $r + 1
    $logCntr = 0
    $isNotPresent = $true
    Write-Progress -Activity "Report $q of $repQty [FILE: $(Split-path $AllReportsListCSV -Leaf)]" -Id 0 -PercentComplete ($q / $repQty * 100)

    Foreach ($l in $allLogs) {
        $logCntr ++
        Write-Progress -Activity "Log $logCntr of $logQty" -Id 1 -PercentComplete ($logCntr / $logQty * 100)
        $v = gi $l.Fullname | Select-string -Pattern $allReports[$r].reportguid 
        
        if ($v.count -gt 0) {
            $isNotPresent = $false

            foreach ($l in $v) {
                $obj = ConvertFrom-Csv $l.line -Delimiter ' ' -Header 'date', 'time', 's-ip', 'cs-method', 'cs-uri-stem', 'cs-uri-query', 's-port', 'cs-username', 'c-ip', 'cs-version', 'cs(User-Agent)', 'cs(Referer)', 'cs-host', 'sc-status', 'sc-substatus', 'sc-win32-status', 'sc-bytes', 'cs-bytes', 'time-taken'
                $actualCheck = $endResults | ? { $_.ReportGuid -eq $allReports[$r].reportguid }
                
                if ($actualCheck.Reportguid.Count -gt 0) {
                    
                    if ($actualCheck.LastAccessTimePerIP.Ip -contains $obj.'c-ip') {
                        
                        if ($actualCheck.LastAccessTimePerIP.LastAccessTime -lt (Get-date($obj.date))) {
                            
                            For ($z = 0; $z -lt $endResults[-1].LastAccessTimePerIP.Count; $z++) {
                                
                                if ($endResults[-1].LastAccessTimePerIP[$z].Ip -eq $obj.'c-ip') {
                                    $endResults[-1].LastAccessTimePerIP[$z].LastAccessTime = $obj.date
                                }
                            }
                        }
                    }
                    else {
                        $endResults | ? { $_.ReportGuid -eq $allReports[$r].reportguid } | % { 
                            $_.LastAccessTimePerIP.add(([PSCustomObject]@{
                                        Ip             = $obj.'c-ip'
                                        LastAccessTime = $obj.date
                                    }
                                )
                            ) > $null
                        }
                    }
                }
                else {
                    $endResults.Add(([Report]@{
                                ReportGuid          = $allReports[$r].reportguid
                                ReportPdfPath       = $allReports[$r].pdffilepath
                                LastAccessTimePerIP = @(@{
                                        Ip             = $obj.'c-ip'
                                        LastAccessTime = $obj.date
                                    }
                                )
                            }    
                        ) 
                    ) > $null
                }
            }
        }
    }

    if ($isNotPresent) {
        $endResults.Add(([Report]@{
                    ReportGuid          = $allReports[$r].reportguid
                    ReportPdfPath       = $allReports[$r].pdffilepath
                    LastAccessTimePerIP = @(@{
                            Ip             = $null
                            LastAccessTime = $null
                        }
                    )
                }    
            ) 
        ) > $null
    }
}

try {
    if (!(Test-Path "$(Split-Path $AllReportsListCSV -Parent)\results")) { New-Item "$(Split-Path $AllReportsListCSV -Parent)\results" -ItemType Directory }
    $endResults | ConvertTo-Json -Depth 4 | Out-File "$(Split-Path $AllReportsListCSV -Parent)\results\results_$((Split-Path $AllReportsListCSV -Leaf).Split('.')[0]).json" -Force -Verbose
}
catch {
    "Error exporting '$(Split-Path $AllReportsListCSV -Leaf)': $($_.exception.message)" | Out-File "$(Split-Path $AllReportsListCSV -Parent)\error_log.txt"
}
