# === CONFIGURATION ===
$tempRoots = @(
    "C:\Users\user\AppData\Local\Temp",
    "C:\Users\user2\AppData\Local\Temp"            # add more paths as needed
)
$daysToKeep = 1
$logDir = "D:\TEMP_CleanUp_LOG"
$timestamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
$rawStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = "$logDir\TempCleanupLog_$rawStamp.txt"

# === DESTINATION SHARED FOLDERS ===
$destinationPaths = @(
    "\\SRV-Storage\Logs_Cleaning_Reports\TEMP_Logs"
)

# === Ensure log directory exists ===
if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}

# === System Info ===
$hostname = $env:COMPUTERNAME
$ipAddress = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
    $_.IPAddress -notlike '169.*' -and $_.InterfaceAlias -notlike 'Loopback*' -and $_.PrefixOrigin -ne 'WellKnown'
} | Select-Object -First 1 -ExpandProperty IPAddress)

# === Function to get drive space ===
function Get-DriveFreeSpace($letter) {
    $drive = Get-PSDrive $letter
    return "{0:N2} GB" -f ($drive.Free / 1GB)
}

# === INIT LOG ===
"=== Multi-Temp Repository Cleanup Started: $timestamp ===" | Out-File -FilePath $logFile
"Host: $hostname ($ipAddress)" | Out-File -Append $logFile
"" | Out-File -Append $logFile
"Before Cleanup:" | Out-File -Append $logFile
"C:\ Free Space: $(Get-DriveFreeSpace 'C')" | Out-File -Append $logFile
"D:\ Free Space: $(Get-DriveFreeSpace 'D')" | Out-File -Append $logFile
"" | Out-File -Append $logFile

# === MAIN CLEANUP LOGIC ===
$totalDeletedSize = 0
$cutoffDate = (Get-Date).AddDays(-$daysToKeep)
$cleanedItemsLog = @()

foreach ($targetRoot in $tempRoots) {
    if (Test-Path $targetRoot) {
        "----- Cleaning: $targetRoot -----" | Out-File -Append $logFile

        # Get all child items (files/folders) older than cutoff
        $items = Get-ChildItem -Path $targetRoot -Force | Where-Object {
            $_.LastWriteTime -lt $cutoffDate
        }

        foreach ($item in $items) {
            try {
                $lastMod = $item.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                if ($item.PSIsContainer) {
                    $sizeBytes = (Get-ChildItem -LiteralPath $item.FullName -Recurse -Force -ErrorAction SilentlyContinue |
                                  Where-Object { -not $_.PSIsContainer } | Measure-Object Length -Sum).Sum
                    $sizeMB = if ($sizeBytes) { [math]::Round($sizeBytes / 1MB, 2) } else { 0 }
                    Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
                    $totalDeletedSize += $sizeBytes
                    $cleanedItemsLog += "🗑️ Deleted folder: $($item.FullName) | Size: $sizeMB MB | Last Modified: $lastMod"
                } else {
                    $sizeMB = [math]::Round($item.Length / 1MB, 2)
                    Remove-Item -LiteralPath $item.FullName -Force -ErrorAction Stop
                    $totalDeletedSize += $item.Length
                    $cleanedItemsLog += "🗑️ Deleted file:   $($item.FullName) | Size: $sizeMB MB | Last Modified: $lastMod"
                }
            } catch {
                $cleanedItemsLog += "⚠️ Failed to delete $($item.FullName) | Error: $($_.Exception.Message)"
            }
        }

        "Completed sweep of $targetRoot" | Out-File -Append $logFile
        "" | Out-File -Append $logFile
    } else {
        "❌ Target path not found: $targetRoot" | Out-File -Append $logFile
        "" | Out-File -Append $logFile
    }
}

# === POST-CLEANUP STATS ===
$totalDeletedMB = [math]::Round($totalDeletedSize / 1MB, 2)
"Total Deleted Size Across All Targets: $totalDeletedMB MB" | Out-File -Append $logFile
"" | Out-File -Append $logFile
"After Cleanup:" | Out-File -Append $logFile
"C:\ Free Space: $(Get-DriveFreeSpace 'C')" | Out-File -Append $logFile
"D:\ Free Space: $(Get-DriveFreeSpace 'D')" | Out-File -Append $logFile
"" | Out-File -Append $logFile
"=========== Deleted Items Details ============" | Out-File -Append $logFile
$cleanedItemsLog | Out-File -Append $logFile
"=== Cleanup Completed: $(Get-Date -Format 'MM/dd/yyyy HH:mm:ss') ===" | Out-File -Append $logFile

# === COPY TO SHARED FOLDERS ===
$copySuccessCount = 0
$totalDestinations = $destinationPaths.Count

foreach ($dest in $destinationPaths) {
    if (Test-Path -Path $dest) {
        try {
            Copy-Item -Path $logFile -Destination $dest -Force
            "✅ Log copied to $dest" | Out-File -Append $logFile
            $copySuccessCount++
        } catch {
            "❌ Failed to copy to $dest - $_" | Out-File -Append $logFile
        }
    } else {
        "⚠️ Destination not found: $dest" | Out-File -Append $logFile
    }
}

# === CONDITIONAL CLEANUP OF LOCAL LOG ===
$deleteLog = "$logDir\CleanupDeleteLog_$rawStamp.txt"
if ($copySuccessCount -eq $totalDestinations) {
    try {
        Remove-Item -Path $logFile -Force
        "🗑️ Original log file deleted after successful copy to all destinations." | Out-File -FilePath $deleteLog
    } catch {
        "⚠️ Failed to delete original log file: $_" | Out-File -FilePath $deleteLog
    }
} else {
    "🚫 Original log file NOT deleted. Only $copySuccessCount out of $totalDestinations destinations succeeded." | Out-File -FilePath $deleteLog
}
