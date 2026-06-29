# ================= CONFIG =================
$SourceRoot     = "E:\test"
$DestinationRoot  = "\\backup-test\BACKUP\Backup-ANY" # Mandatory UNC Path

$Prefix        = "ANY-"
$BackupRetentionDays = 3
$LogRetentionDays  = 7
$LogRoot       = "E:\test\ZIP_Logs"
$7ZipPath      = "C:\Program Files\7-Zip\7z.exe"

# ================= EMAIL CONFIGURATION (ZONE 1 MATCH) =================
[string]$SmtpServer          = "smtp-mail.outlook.com"
[int]$SmtpPort               = 587
[string]$EmailFrom           = "notifier@infoxtek.com"      # Must match authenticated mailbox
[string]$HotmailPlainPassword= "SMTPpassword_apppassword"          # Your 16-Character App Password
[string]$EmailTo             = "admin1@infoxtek.com"  
[string]$CriticalEmailCc     = "manager1@infoxtek.com"

# --- RETENTION WATCHER LOGIC ---
$MaxWaitMinutes = 30
$CheckIntervalSeconds = 60
#>

# Network Authentication Credentials (For Task Scheduler Context)
$NetUser = "infoxtekdomain.local\serviceacc_name" 
$NetPass = "serviceacc_PW"
<#
=================================================================================
  ENTERPRISE DATABASE BACKUP AUTOMATION PIPELINE (FIXED: STRING PARSER SYNTAX)
=================================================================================
#>


# ================= DYNAMIC RUNTIME PATH GENERATION =================
$DateStamp    = Get-Date -Format "dd-MMM-yyyy"
$SourceFolder = Join-Path $SourceRoot "$Prefix$DateStamp"
if (-not (Test-Path $LogRoot)) { New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null }

$ProcessLog = Join-Path $LogRoot "Backup_Process_$DateStamp.log"
$ReportLog  = Join-Path $LogRoot "Backup_Report_$DateStamp.log"

# --- GLOBAL TELEMETRY TRACKING ARRAYS (FOR EMAIL METRICS REPORTING) ---
$Global:PurgedAssets     = [System.Collections.Generic.List[string]]::new()
$Global:MaintainedAssets = [System.Collections.Generic.List[string]]::new()
$Global:FinalDestSize    = "0.00 GB"
$Global:WorkflowRoute    = "Pending Assessment"

Function Write-Report ($Message) {
    $Stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[REPORTS] $Stamp - $Message" | Out-File -FilePath $ReportLog -Append -Encoding UTF8 -Force
    Write-Output "[REPORTS] $Stamp - $Message"
}

# --- CENTRALIZED MULTI-COLOR HTML EMAIL ENGINE ---
Function Send-StructuralEmail ($Status, $Subject, $FailedStage = "", $ErrorException = "") {
    try {
        $SecurePass = ConvertTo-SecureString $HotmailPlainPassword -AsPlainText -Force
        $Cred = New-Object System.Management.Automation.PSCredential($EmailFrom, $SecurePass)
        $Timestamp = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
        $NextDayDate = (Get-Date).AddDays(1).ToString("dd-MMM-yyyy")
        $NextSequenceTarget = Join-Path $SourceRoot "$Prefix$NextDayDate"

        # Dynamically set headers based on runtime status
        switch ($Status) {
            "SUCCESS" {
                $HeaderBgColor = "#28a745"
                $BannerText    = "✅ DB BACKUP COPIED  AUTOMATION SUCCESSFUL"
            }
            "WARNING" {
                $HeaderBgColor = "#f0ad4e"
                $BannerText    = "⚠️ DB BACKUP COPIED  COMPLETED WITH WARNINGS"
            }
            "ERROR" {
                $HeaderBgColor = "#d9534f"
                $BannerText    = "❌ DB BACKUP COPIED  CRITICAL FAILURE"
            }
        }

        # Build dynamic HTML lists for assets
        $PurgedHtml = ""
        foreach ($Asset in $Global:PurgedAssets) { $PurgedHtml += "<li style='margin-bottom: 4px;'>$Asset</li>" }
        if ($PurgedHtml -eq "") { $PurgedHtml = "<li>No assets were modified or deleted during this lifecycle window.</li>" }

        $MaintainedHtml = ""
        foreach ($Asset in $Global:MaintainedAssets) { $MaintainedHtml += "<li style='margin-bottom: 4px;'>$Asset</li>" }
        if ($MaintainedHtml -eq "") { $MaintainedHtml = "<li>No historical dependencies found within log/backup paths.</li>" }

        $Body = @"
<html>
<body style='font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 20px;'>
    <div style='background-color: $HeaderBgColor; color: white; padding: 12px 20px; font-size: 18px; font-weight: bold; border-radius: 4px; margin-bottom: 20px;'>
        $BannerText
    </div>
    
    <h3 style='color: #444; margin-bottom: 10px;'>Operational Summary Metrics</h3>
    <table border='1' cellpadding='10' cellspacing='0' style='border-collapse: collapse; width: 100%; max-width: 900px; border: 1px solid #cccccc; font-size: 14px;'>
        <tbody>
            <tr style='background-color: #f9f9f9;'>
                <td style='width: 30%; font-weight: bold; border: 1px solid #cccccc;'>Target Server Hostname:</td>
                <td style='border: 1px solid #cccccc;'>$env:COMPUTERNAME</td>
            </tr>
            <tr>
                <td style='font-weight: bold; border: 1px solid #cccccc;'>Predefined Prefix Used:</td>
                <td style='border: 1px solid #cccccc; color: #d9534f; font-weight: bold;'>$Prefix</td>
            </tr>
            <tr style='background-color: #f9f9f9;'>
                <td style='font-weight: bold; border: 1px solid #cccccc;'>Workflow Route Tracked:</td>
                <td style='border: 1px solid #cccccc; font-weight: bold;'>$Global:WorkflowRoute</td>
            </tr>
            $(if ($Status -eq "ERROR") {
            "<tr style='background-color: #fff5f5;'>
                <td style='font-weight: bold; color: #d9534f; border: 1px solid #cccccc;'>Failed Stage:</td>
                <td style='font-weight: bold; color: #d9534f; border: 1px solid #cccccc;'>$FailedStage</td>
            </tr>
            <tr style='background-color: #fff5f5;'>
                <td style='font-weight: bold; color: #d9534f; border: 1px solid #cccccc;'>Error Exception:</td>
                <td style='font-family: Consolas, monospace; color: #d9534f; border: 1px solid #cccccc;'>$ErrorException</td>
            </tr>"
            })
            <tr>
                <td style='font-weight: bold; border: 1px solid #cccccc;'>Staging Root Path:</td>
                <td style='border: 1px solid #cccccc;'>$SourceRoot</td>
            </tr>
            <tr style='background-color: #f9f9f9;'>
                <td style='font-weight: bold; border: 1px solid #cccccc;'>Network Shares Location:</td>
                <td style='border: 1px solid #cccccc; color: #0275d8;'>$DestinationRoot</td>
            </tr>
            <tr>
                <td style='font-weight: bold; border: 1px solid #cccccc;'>Double-Verified Size:</td>
                <td style='border: 1px solid #cccccc; font-weight: bold;'>$Global:FinalDestSize</td>
            </tr>
            <tr style='background-color: #f9f9f9;'>
                <td style='font-weight: bold; border: 1px solid #cccccc;'>Next Sequence Target Created:</td>
                <td style='border: 1px solid #cccccc;'>$NextSequenceTarget</td>
            </tr>
        </tbody>
    </table>
    
    <h3 style='color: #d9534f; margin-top: 25px; margin-bottom: 5px;'>Purged & Deleted System Assets</h3>
    <ul style='padding-left: 20px; font-size: 14px; color: #555;'>
        $PurgedHtml
    </ul>
    
    <h3 style='color: #0275d8; margin-top: 25px; margin-bottom: 5px;'>Maintained System Assets (ZIP Retention: $BackupRetentionDays Days | Log Retention: $LogRetentionDays Days)</h3>
    <ul style='padding-left: 20px; font-size: 14px; color: #555;'>
        $MaintainedHtml
    </ul>
    
    <hr style='border: 0; border-top: 1px solid #e0e0e0; margin-top: 30px;'>
    <div style='font-size: 12px; color: #888; line-height: 1.6;'>
        Automated task transmission event protocol. Raw output transaction telemetry logs appended directly inside root network systems.<br>
        <b>Audit Execution Timestamp:</b> $Timestamp | <b>System IP Context:</b> $((Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object {$_.IPAddress}).IPAddress | Select-Object -First 1)
    </div>
</body>
</html>
"@

        $MailParams = @{
            SMTPServer  = $SmtpServer
            Port        = $SmtpPort
            UseSsl      = $true
            Credential  = $Cred
            From        = $EmailFrom
            To          = $EmailTo
            Subject     = $Subject
            Body        = $Body
            BodyAsHtml  = $true
            Priority    = if ($Status -eq "ERROR") { "High" } else { "Normal" }
        }
        if ($CriticalEmailCc) { $MailParams.Cc = $CriticalEmailCc }
        
        Send-MailMessage @MailParams
        Write-Report "SUCCESS: Status email notification sent to recipients."
    } catch {
        Write-Report "CRITICAL ERROR: Failed to dispatch SMTP alert message. Error: $_"
    }
}

# --- INITIALIZE LIFECYCLE ---
Write-Report "===== BACKUP PIPELINE START ====="
Write-Report "System Configured Prefix: $Prefix"
Write-Report "Source Folder Configured: $SourceFolder"
Write-Report "Destination Path Configured: $DestinationRoot"

# 1. NETWORK AUTHENTICATION AND 1219 CONFLICT RESOLUTION
try {
    Write-Report "AUTHENTICATION: Purging pre-existing overlapping connections to target network space..."
    $ServerUNC = ($DestinationRoot -split '\\')[2]
    net use | Out-String | ForEach-Object {
        if ($_ -match "\\\\$ServerUNC") {
            net use "\\$ServerUNC" /delete /y | Out-Null
        }
    }
    
    Write-Report "AUTHENTICATION: Mapping secure system credential context for Destination Share..."
    $NetUseCmd = "net use `"$DestinationRoot`" `"$NetPass`" /USER:`"$NetUser`" /PERSISTENT:NO"
    $NetOutput = Invoke-Expression $NetUseCmd 2>&1 | Out-String
    
    if (-not (Test-Path $DestinationRoot)) { throw "Network path unreachable. Output: $NetOutput" }
    Write-Report "SUCCESS: Authenticated and verified destination share path connection."
} catch {
    $ErrMsg = "Network share target destination not accessible via Cloud Task Scheduler: $DestinationRoot"
    Write-Report "CRITICAL AUTHENTICATION FAILURE: $ErrMsg"
    Send-StructuralEmail -Status "ERROR" -Subject "⚠️ CRITICAL BACKUP ALERT: Server [$env:COMPUTERNAME] Failed during execution!" -FailedStage "Authentication" -ErrorException "$ErrMsg (Internal: $_)"
    Exit
}

# 2. EVALUATE WORKFLOW ROUTE
$PreExistingZip = Get-ChildItem -Path $SourceRoot -Filter "$Prefix*.zip" | Select-Object -First 1

$ZipVerified100Success = $false
$FolderVerified100Success = $false
$ZipFileName   = "$Prefix$DateStamp.zip"
$ZipToCopyPath = Join-Path $SourceRoot $ZipFileName

if ($PreExistingZip) {
    # ------------------ ROUTE A: PRE-EXISTING ZIP MOVE ------------------
    $ZipToCopyPath = $PreExistingZip.FullName
    $ZipFileName   = $PreExistingZip.Name
    $Global:WorkflowRoute = "Pre-Existing ZIP Move"
    Write-Report "WORKFLOW ROUTE FOUND: Pre-existing ZIP file matching prefix discovered -> [$ZipFileName]. Skipping compression phase."
    
    $ZipVerified100Success = $true
    $Global:MaintainedAssets.Add("Original Database Folder Retained (Bypassed via Pre-Existing Workflow Path): $SourceFolder")
} else {
    # ------------------ ROUTE B: FULL COMPRESSION AND MOVE ------------------
    $Global:WorkflowRoute = "Full Compression and Move"
    if (-not (Test-Path $SourceFolder)) {
        $ErrMsg = "Source directory not found: $SourceFolder"
        Write-Report "CRITICAL WORKFLOW FAILURE: $ErrMsg"
        Send-StructuralEmail -Status "ERROR" -Subject "⚠️ CRITICAL BACKUP ALERT: Server [$env:COMPUTERNAME] Failed during execution!" -FailedStage "Validation" -ErrorException $ErrMsg
        Exit
    }

    Write-Report "WORKFLOW ROUTE FOUND: No pre-existing ZIP file found. Initializing 7-Zip compression for prefix folder: $SourceFolder"
    $ZipArgs = "a -tzip -mx3 -mmt=on `"$ZipToCopyPath`" `"$SourceFolder`""
    Write-Report "ZIPPING: Deploying .NET Diagnostic Native Process Engine Handle..."
    
    $StartInfo = New-Object System.Diagnostics.ProcessStartInfo -Property @{
        FileName               = $7ZipPath
        Arguments              = $ZipArgs
        RedirectStandardOutput = $true
        RedirectStandardError  = $true
        UseShellExecute        = $false
        CreateNoWindow         = $true
    }
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $StartInfo
    
    $StartTime = Get-Date
    if ($Process.Start()) {
        $CapturedPID = $Process.Id
        Write-Report "SUCCESS: 7-Zip engine engaged safely with native Process ID [PID: $CapturedPID]."
    } else {
        $ErrMsg = "Engine Launch Failure: System diagnostics failed to secure process thread initialization."
        Write-Report "CRITICAL WORKFLOW FAILURE: $ErrMsg"
        Send-StructuralEmail -Status "ERROR" -Subject "📂 ❌ BACKUP CRITICAL FAILURE - Compression Corrupted" -FailedStage "Compression Initialization" -ErrorException $ErrMsg
        Exit
    }
    
    $OriginalSizeGB = [Math]::Round((Get-ChildItem $SourceFolder -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum / 1GB, 2)
    if ($OriginalSizeGB -eq 0) { $OriginalSizeGB = 0.01 }

    while (-not $Process.HasExited) {
        Start-Sleep -Seconds 15
        $TotalDurationMinutes = [Math]::Round(((Get-Date) - $StartTime).TotalMinutes, 2)
        
        if ($TotalDurationMinutes -gt 180) { # 3-Hour Safety Breaker
            Write-Report "CRITICAL SYSTEM EXCEPTION: Runtime exceeded limits. Terminating PID $CapturedPID..."
            $Process.Kill()
            break
        }

        $CurrentZipSizeGB = if (Test-Path $ZipToCopyPath) { [Math]::Round((Get-Item $ZipToCopyPath).Length / 1GB, 2) } else { 0 }
        $Speed = if ($TotalDurationMinutes -gt 0) { [Math]::Round($CurrentZipSizeGB / $TotalDurationMinutes, 2) } else { 0 }
        $Ratio = [Math]::Round(($CurrentZipSizeGB / $OriginalSizeGB) * 100, 0)
        if ($Ratio -gt 100) { $Ratio = 100 }

        $LogBlock = "================ ZIPPING PROCESS =================`r`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - PID=$CapturedPID | Elapsed=$TotalDurationMinutes min | ETA=DYNAMIC | Original=$OriginalSizeGB GB | Processed~$CurrentZipSizeGB GB | Compressed=$CurrentZipSizeGB GB | Ratio~$Ratio % | Speed=$Speed GB/min | Location=LOCAL"
        $LogBlock | Out-File -FilePath $ProcessLog -Append -Encoding UTF8 -Force
    }

    if ($Process.ExitCode -eq 0 -and (Test-Path $ZipToCopyPath)) {
        $LocalZipSize = (Get-Item $ZipToCopyPath).Length
        if ($LocalZipSize -gt 0) {
            Write-Report "VERIFICATION SUCCESS: 7-Zip compression finalized 100% successfully. Exit Code 0."
            $ZipVerified100Success = $true
            $FolderVerified100Success = $true
        }
    }
    
    if (-not $ZipVerified100Success) {
        $ErrMsg = "7-Zip check failed! Process exited with error code ($($Process.ExitCode)). Source folder preservation lock engaged."
        Write-Report "CRITICAL WORKFLOW FAILURE: $ErrMsg"
        Send-StructuralEmail -Status "ERROR" -Subject "📂 ❌ BACKUP CRITICAL FAILURE - Compression Corrupted" -FailedStage "Compression" -ErrorException $ErrMsg
        Exit
    }
}

# 3. SECURE MULTI-THREADED NETWORK DATA TRANSFER (ROBOCOPY ENGINE)
Write-Report "TRANSFER: Initializing multi-threaded Robocopy operational sequence to destination share..."
$SourceDirOnly = Split-Path $ZipToCopyPath
$DestDirOnly   = Join-Path $DestinationRoot "$Prefix$DateStamp"

if (-not (Test-Path $DestDirOnly)) { New-Item -ItemType Directory -Path $DestDirOnly -Force | Out-Null }

$RoboArgs = @("`"$SourceDirOnly`"", "`"$DestDirOnly`"", "`"$ZipFileName`"", "/MT:8", "/Z", "/R:3", "/W:5", "/V", "/TS", "/FP", "/LOG+:`"$ReportLog`"")
$RoboProcess = Start-Process -FilePath "robocopy.exe" -ArgumentList $RoboArgs -PassThru -NoNewWindow -Wait

if ($RoboProcess.ExitCode -ge 8) {
    $ErrMsg = "Robocopy transfer failed with network fault. Exit Code: $($RoboProcess.ExitCode)"
    Write-Report "CRITICAL DATA CORRUPTION WARNING: $ErrMsg"
    Send-StructuralEmail -Status "ERROR" -Subject "⚠️ CRITICAL BACKUP ALERT: Server [$env:COMPUTERNAME] Failed during execution!" -FailedStage "Network Transfer" -ErrorException $ErrMsg
    Exit
}

# 4. STRICT DOUBLE-VERIFICATION PROTOCOL GATE
Write-Report "VERIFICATION: Initiating explicit bit-level file synchronization validation check..."
$TargetDestinationFile = Join-Path $DestDirOnly $ZipFileName
$FileCopied100Success = $false

if (Test-Path $TargetDestinationFile) {
    $LocalSizeB  = (Get-Item $ZipToCopyPath).Length
    $RemoteSizeB = (Get-Item $TargetDestinationFile).Length
    
    if ($LocalSizeB -eq $RemoteSizeB -and $LocalSizeB -gt 0) {
        $Global:FinalDestSize = "$([Math]::Round($RemoteSizeB / 1GB, 2)) GB"
        Write-Report "VERIFICATION SUCCESS: File match confirmed. ($LocalSizeB B) matches perfectly. 100% transfer success validated."
        $FileCopied100Success = $true
    } else {
        Write-Report "VERIFICATION FAILED: Byte size mismatch! Source: $LocalSizeB B vs Destination: $RemoteSizeB B"
    }
} else {
    Write-Report "VERIFICATION FAILED: Output archive completely missing from destination share location."
}

if (-not $FileCopied100Success) {
    $ErrMsg = "Double-verification gate rejected file transfer! File sizes do not match."
    Write-Report "CRITICAL SECURITY BREACH: $ErrMsg"
    Send-StructuralEmail -Status "ERROR" -Subject "⚠️ CRITICAL BACKUP ALERT: Server [$env:COMPUTERNAME] Failed during execution!" -FailedStage "Destination Data Verification" -ErrorException $ErrMsg
    Exit
}

# 5. SAFE CLEANUP LOGIC WITH METRIC TRACKING
Write-Report "CLEANUP: Safe confirmation flags matched. Initiating targeted source asset cleanup..."
if ($FileCopied100Success -and (Test-Path $ZipToCopyPath)) {
    Remove-Item -Path $ZipToCopyPath -Force -Confirm:$false
    Write-Report "CLEANUP SUCCESS: Local source zip archive erased safely from staging partition."
    $Global:PurgedAssets.Add("Local Source ZIP Archive Safely Cleaned: $ZipToCopyPath")
}
if ($FolderVerified100Success -and (Test-Path $SourceFolder)) {
    Remove-Item -Path $SourceFolder -Recurse -Force -Confirm:$false
    Write-Report "CLEANUP SUCCESS: Original raw directory purged cleanly from local storage system."
    $Global:PurgedAssets.Add("Local Source Raw Directory Safely Purged: $SourceFolder")
}

# 6. PROVISION NEXT DAY PLACEHOLDER CONTAINER
$NextDayDate   = (Get-Date).AddDays(1).ToString("dd-MMM-yyyy")
$NextDayFolder = Join-Path $SourceRoot "$Prefix$NextDayDate"
if (-not (Test-Path $NextDayFolder)) {
    New-Item -ItemType Directory -Path $NextDayFolder -Force | Out-Null
    Write-Report "LIFECYCLE PROVISIONING: Next sequence placeholder folder container generated at: $NextDayFolder"
}

# 7. RETENTION POLICY ENGINE WITH LOG & METRIC AUDITING
$Now = Get-Date
$HasWarnings = $false

# A. Evaluate Historical Network Folders (ZIP Containers)
$NetworkFolders = Get-ChildItem -Path $DestinationRoot -Directory | Where-Object { $_.Name -like "$Prefix*" }
foreach ($Folder in $NetworkFolders) {
    $Age = [Math]::Floor(($Now - $Folder.CreationTime).TotalDays)
    if ($Age -ge $BackupRetentionDays) {
        try {
            Remove-Item -Path $Folder.FullName -Recurse -Force -Confirm:$false
            Write-Report "RETENTION PURGE: Erased expired historic network archive directory container: $($Folder.Name)"
            $Global:PurgedAssets.Add("Network Storage Expired ZIP Container Purged: $($Folder.FullName) (Age: $Age Days Old)")
        } catch {
            Write-Report "RETENTION WARNING: Failed to remove expired network folder: $($Folder.FullName). Error: $_"
            $HasWarnings = $true
        }
    } else {
        $Global:MaintainedAssets.Add("Network Storage Active ZIP Archive Maintained: $($Folder.Name) (Age: $Age Days Old)")
    }
}

# B. Evaluate Historical System Engine Logs
$HistoricalLogs = Get-ChildItem -Path $LogRoot -File | Where-Object { $_.Name -like "*.log" }
foreach ($LogFile in $HistoricalLogs) {
    $Age = [Math]::Floor(($Now - $LogFile.CreationTime).TotalDays)
    if ($Age -ge $LogRetentionDays) {
        try {
            Remove-Item -Path $LogFile.FullName -Force -Confirm:$false
            Write-Report "RETENTION PURGE: Cleaned historical engine log asset: $($LogFile.Name)"
            $Global:PurgedAssets.Add("Local Historical Log File Purged: $($LogFile.Name) (Age: $Age Days Old)")
        } catch {
            Write-Report "RETENTION WARNING: Failed to purge local log file: $($LogFile.FullName). Error: $_"
            $HasWarnings = $true
        }
    } else {
        $Global:MaintainedAssets.Add("Local Staging Log Maintained: $($LogFile.Name) (Age: $Age Days Old)")
    }
}

# 8. TRANSMIT FINAL EXECUTION METRICS ALERT
$FinalStatus = if ($HasWarnings) { "WARNING" } else { "SUCCESS" }
$SubjectEmoji = if ($HasWarnings) { "⚠️" } else { "✅" }

Send-StructuralEmail -Status $FinalStatus -Subject "$SubjectEmoji BACKUP REPLICATION COMPLETION - $Prefix - Size: $Global:FinalDestSize"
Write-Report "===== BACKUP PIPELINE END ====="
