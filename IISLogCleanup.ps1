# === CONFIGURATION ===
$logFolders = @(
  "C:\inetpub\logs\LogFiles\W3SVC1",
    "C:\inetpub\logs\LogFiles\W3SVC2",
    "C:\inetpub\logs\LogFiles\W3SVC3",
    "C:\inetpub\logs\LogFiles\W3SVC4",
    "C:\inetpub\logs\LogFiles\W3SVC5",
    "C:\inetpub\logs\LogFiles\W3SVC6",
    "C:\inetpub\logs\LogFiles\W3SVC7",
    "C:\inetpub\logs\LogFiles\W3SVC8",
    "C:\inetpub\logs\LogFiles\W3SVC9",
    "C:\inetpub\logs\LogFiles\W3SVC10",
    "C:\inetpub\logs\LogFiles\W3SVC11",
    "C:\inetpub\logs\LogFiles\W3SVC12",
    "C:\inetpub\logs\LogFiles\W3SVC13",
    "C:\inetpub\logs\LogFiles\W3SVC14",
    "C:\inetpub\logs\LogFiles\W3SVC15",
    "C:\inetpub\logs\LogFiles\W3SVC16",
    "C:\inetpub\logs\LogFiles\W3SVC17",
    "C:\inetpub\logs\LogFiles\W3SVC18",
    "C:\inetpub\logs\LogFiles\W3SVC19",
    "C:\inetpub\logs\LogFiles\W3SVC20",
    "C:\inetpub\logs\LogFiles\W3SVC21",
    "C:\inetpub\logs\LogFiles\W3SVC22",
    "C:\inetpub\logs\LogFiles\W3SVC23",
    "C:\inetpub\logs\LogFiles\W3SVC24"
)
$daysToKeep = 2
$logDir = "D:\IIS_CleanUp_Report"
$timestamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
$rawStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = "$logDir\CleanupLog_$rawStamp.txt"

# === DESTINATION SHARED FOLDERS ===
$destinationPaths = @(
    "\\Storage\Logs_Cleaning_Reports\$hostname--$ipAddress"
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
"=== IIS Log Cleanup Started: $timestamp ===" | Out-File -FilePath $logFile
"Hostname: $hostname" | Out-File -Append $logFile
"Host IP: $ipAddress" | Out-File -Append $logFile
"" | Out-File -Append $logFile
"Before Cleanup:" | Out-File -Append $logFile
"C:\ Free Space: $(Get-DriveFreeSpace 'C')" | Out-File -Append $logFile
"D:\ Free Space: $(Get-DriveFreeSpace 'D')" | Out-File -Append $logFile
"" | Out-File -Append $logFile

# === MAIN CLEANUP LOGIC ===
$totalDeletedSize = 0
$cutoffDate = (Get-Date).Date.AddDays(-$daysToKeep)
$cleanedFilesLog = @()

foreach ($folder in $logFolders) {
    if ([string]::IsNullOrWhiteSpace($folder)) {
        continue
    }

    $folderPath = $folder.TrimEnd('\')
    $folderDeletedSize = 0
    $deletedCount = 0
    $fileDetails = @()

    if (Test-Path $folderPath) {
        $files = Get-ChildItem -Path $folderPath -File -Recurse | Where-Object {
            $_.LastWriteTime -lt $cutoffDate
        }

        foreach ($file in $files) {
            try {
                $sizeMB = [math]::Round($file.Length / 1MB, 2)
                $lastMod = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                $fileDetails += "🗑️ Deleted: $($file.FullName) | Size: $sizeMB MB | Last Modified: $lastMod"
                $folderDeletedSize += $file.Length
                $deletedCount++
                Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
            } catch {
                $fileDetails += "⚠️ Failed: $($file.FullName) | Error: $_"
            }
        }

        $folderDeletedMB = [math]::Round($folderDeletedSize / 1MB, 2)
        "Checking: $folderPath" | Out-File -Append $logFile
        "Deleted $deletedCount files from $folderPath | Size: $folderDeletedMB MB" | Out-File -Append $logFile
        "" | Out-File -Append $logFile

        $totalDeletedSize += $folderDeletedSize

        if ($deletedCount -gt 0) {
            $fileDetails += "✅ Deleted $deletedCount files from $folderPath | Total Size: $folderDeletedMB MB"
            $cleanedFilesLog += $fileDetails + ""
        }
    } else {
        "❌ Folder not found: $folderPath" | Out-File -Append $logFile
        "" | Out-File -Append $logFile
    }
}

# === FINAL STATS ===
$totalDeletedMB = [math]::Round($totalDeletedSize / 1MB, 2)
"Total Deleted Size: $totalDeletedMB MB" | Out-File -Append $logFile
"" | Out-File -Append $logFile
"After Cleanup:" | Out-File -Append $logFile
"C:\ Free Space: $(Get-DriveFreeSpace 'C')" | Out-File -Append $logFile
"D:\ Free Space: $(Get-DriveFreeSpace 'D')" | Out-File -Append $logFile
"" | Out-File -Append $logFile
"=== Cleanup Completed: $(Get-Date -Format 'MM/dd/yyyy HH:mm:ss') ===" | Out-File -Append $logFile

# === COPY TO SHARED FOLDERS ===
$copySuccessCount = 0
$totalDestinations = $destinationPaths.Count

foreach ($dest in $destinationPaths) {
    if (Test-Path -path $dest) {
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

# === APPEND CLEANED FILE DETAILS ===
"" | Out-File -Append $logFile
"===========Cleaned Files============" | Out-File -Append $logFile
$cleanedFilesLog | Out-File -Append $logFile

# === DELETE LOCAL LOG IF COPIED ===
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