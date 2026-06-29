
# ================== CONFIGURATION ==================
$BasePath = "D:\test_shared\"
$NetworkShare = "\\SRV-SHARED\Destination\TEst_SRV_Path\"
$LogPath = "D:\test_shared\Logs\ZipProcess.log"


$Prefix = "Test-"

# ================== LOG FUNCTION ==================
function Write-Log {
    param ([string]$Message)
    $TimeStamp = (Get-Date).ToString("dd-mm-yyyy HH:mm:ss")
    "$TimeStamp - $Message" | Out-File -FilePath $LogPath -Append -Encoding utf8
}

# ================== DATE FORMAT ==================
$Today    = Get-Date
$TodayStr = $Prefix + $Today.ToString("dd-MMM-yyyy").ToUpper()

$Tomorrow    = $Today.AddDays(1)
$TomorrowStr = $Prefix + $Tomorrow.ToString("dd-MMM-yyyy").ToUpper()

# ================== PATHS ==================
$SourceFolder   = Join-Path $BasePath $TodayStr
$ZipFileName    = "$TodayStr.zip"
$DestinationZip = Join-Path $NetworkShare $ZipFileName

Write-Log "---------------- START PROCESS ----------------"
Write-Log "Target Folder: $SourceFolder"

# ================== CHECK FOLDER ==================
if (-not (Test-Path $SourceFolder)) {
    Write-Log "ERROR: Folder not found: $SourceFolder"
    exit
}

# ================== CALCULATE ORIGINAL SIZE ==================
Try {
    $OriginalSize = (Get-ChildItem $SourceFolder -Recurse -File | Measure-Object Length -Sum).Sum
    $OriginalSizeMB = [math]::Round($OriginalSize / 1MB, 2)
    Write-Log "Original Size: $OriginalSizeMB MB"
}
Catch {
    Write-Log "ERROR calculating original size: $_"
}

# ================== COMPRESSION ==================
$CompressionWatch = [System.Diagnostics.Stopwatch]::StartNew()

Try {
    Compress-Archive -Path $SourceFolder -DestinationPath $DestinationZip -Force
    $CompressionWatch.Stop()
    Write-Log "Compression SUCCESS"
}
Catch {
    $CompressionWatch.Stop()
    Write-Log "ERROR during compression: $_"
    exit
}

$CompressionDuration = [math]::Round($CompressionWatch.Elapsed.TotalSeconds, 2)
Write-Log "Compression Time: $CompressionDuration sec"

# ================== ZIP SIZE ==================
Try {
    $ZipSize = (Get-Item $DestinationZip).Length
    $ZipSizeMB = [math]::Round($ZipSize / 1MB, 2)
    Write-Log "ZIP Size: $ZipSizeMB MB"
}
Catch {
    Write-Log "ERROR getting ZIP size: $_"
}

# ================== DELETE ORIGINAL FOLDER ==================
Try {
    Remove-Item -Path $SourceFolder -Recurse -Force
    Write-Log "Original folder deleted"
}
Catch {
    Write-Log "ERROR deleting source folder: $_"
}

# ================== CREATE TOMORROW FOLDER ==================
$NewFolder = Join-Path $BasePath $TomorrowStr

Try {
    if (-not (Test-Path $NewFolder)) {
        New-Item -Path $NewFolder -ItemType Directory | Out-Null
        Write-Log "Created next folder: $NewFolder"
    }
    else {
        Write-Log "Next folder already exists: $NewFolder"
    }
}
Catch {
    Write-Log "ERROR creating next folder: $_"
}

# ================== END ==================
Write-Log "Process completed"
Write-Log "---------------- END PROCESS ----------------`n"
