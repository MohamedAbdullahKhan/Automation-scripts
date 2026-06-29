<#
=================================================================================
  ENTERPRISE DATABASE BACKUP AUTOMATION PIPELINE (DIRECT REPLICATION REVISION)
=================================================================================
  System Architecture: High-Throughput Unbuffered Directory Replication Engine
  Target Payload: Raw Database Directory Array (Uncompressed)
  Network Optimization: Data/Timestamp Only Stream (Bypasses Owner ACL Latency)
=================================================================================
#>

# --- INITIALIZATION BLOCK ---
$SourceRoot      = "D:\test"
$DestinationRoot  = "\\backup-test\BACKUP\Backup-test" # Mandatory UNC Path
$Prefix        = "any-"
$BackupRetentionDays = 3
$LogRetentionDays  = 7
$LogRoot      = "D:\Backup-test\ZIP_Logs"

# Safety Fallback Enforcement
if ([string]::IsNullOrEmpty($SourceRoot)) { $SourceRoot = "D:\Backup-test" }
if ([string]::IsNullOrEmpty($DestinationRoot)) { $DestinationRoot = "\\backup-test\BACKUP\Backup-test" }
if ([string]::IsNullOrEmpty($Prefix)) { $Prefix = "any-" }
if ([string]::IsNullOrEmpty($LogRoot)) { $LogRoot = "D:\Backup-test\ZIP_Logs" }

# ================= SMTP EMAIL CONFIGURATION =================
[string]$SmtpServer          = "smtp-mail.outlook.com"
[int]$SmtpPort               = 587
[string]$EmailFrom           = "notification_sender@infoxtek.com"      # Must match authenticated mailbox
[string]$HotmailPlainPassword= "yousmtporapppassword"          # Your 16-Character App Password
[string]$EmailTo             = "support1@infoxtek.com"  
[string]$CriticalEmailCc     = "admin@infoxtek.com"

# Network Authentication Space
$NetUser = "infoxtekdomain.local\serviceaccount_name" 
$NetPass = "serviceaccount_pw"

# ================= DYNAMIC RUNTIME PATH GENERATION =================
$DateStamp           = Get-Date -Format "dd-MMM-yyyy"
$SourceFolder        = Join-Path $SourceRoot "$Prefix$DateStamp"
$DestDirContainerName= "$Prefix$DateStamp"
$DestinationFolder   = Join-Path $DestinationRoot $DestDirContainerName

if (-not (Test-Path $LogRoot)) { 
    New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null 
}

$ReportLog  = Join-Path $LogRoot "Replication_Report_$DateStamp.log"

# --- GLOBAL LIFECYCLE MONITORING ARRAYS ---
$Global:PurgedAssets     = [System.Collections.Generic.List[string]]::new()
$Global:MaintainedAssets = [System.Collections.Generic.List[string]]::new()
$Global:FinalDestSize    = "0.00 GB"
$Global:TransferMetrics  = "N/A"

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

        switch ($Status) {
            "SUCCESS" {
                $HeaderBgColor = "#28a745"
                $BannerText    = "✅ DATABASE DIRECT REPLICATION PIPELINE SUCCESSFUL"
            }
            "WARNING" {
                $HeaderBgColor = "#f0ad4e"
                $BannerText    = "⚠️ DATABASE REPLICATION PIPELINE COMPLETED WITH WARNINGS"
            }
            "ERROR" {
                $HeaderBgColor = "#d9534f"
                $BannerText    = "❌ DATABASE REPLICATION PIPELINE CRITICAL FAILURE"
            }
        }

        $PurgedHtml = ""
        foreach ($Asset in $Global:PurgedAssets) { $PurgedHtml += "<li style='margin-bottom: 4px;'>$Asset</li>" }
        if ($PurgedHtml -eq "") { $PurgedHtml = "<li>No assets were modified or deleted during this lifecycle window.</li>" }

        $MaintainedHtml = ""
        foreach ($Asset in $Global:MaintainedAssets) { $MaintainedHtml += "<li style='margin-bottom: 4px;'>$Asset</li>" }
        if ($MaintainedHtml -eq "") { $MaintainedHtml = "<li>No historical dependencies found within network backup paths.</li>" }

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
                <td style='width: 30%; font-weight: bold; border: 1px solid #cccccc;'>Host Node Name:</td>
                <td style='border: 1px solid #cccccc;'>$env:COMPUTERNAME</td>
            </tr>
            <tr>
                <td style='font-weight: bold; border: 1px solid #cccccc;'>Active Base Prefix:</td>
                <td style='border: 1px solid #cccccc; font-weight: bold; color: #1A365D;'>$Prefix</td>
            </tr>
            $(if ($Status -eq "ERROR") {
            "<tr style='background-color: #fff5f5;'>
                <td style='font-weight: bold; color: #d9534f; border: 1px solid #cccccc;'>Failed Stage:</td>
                <td style='font-weight: bold; color: #d9534f; border: 1px solid #cccccc;'>$FailedStage</td>
            </tr>
            <tr style='background-color: #fff5f5;'>
                <td style='font-weight: bold; color: #d9534f; border: 1px solid #cccccc;'>Error Details:</td>
                <td style='font-family: Consolas, monospace; color: #d9534f; border: 1px solid #cccccc;'>$ErrorException</td>
            </tr>"
            })
            <tr style='background-color: #f9f9f9;'>
                <td style='font-weight: bold; border: 1px solid #cccccc;'>Source Folder:</td>
                <td style='border: 1px solid #cccccc;'>$SourceFolder</td>
            </tr>
            <tr>
                <td style='font-weight: bold; border: 1px solid #cccccc;'>Target Share Location:</td>
                <td style='border: 1px solid #cccccc; color: #0275d8;'>$DestinationFolder</td>
            </tr>
            <tr style='background-color: #f9f9f9;'>
                <td style='font-weight: bold; border: 1px solid #cccccc;'>Verified Payload Size:</td>
                <td style='border: 1px solid #cccccc; font-weight: bold;'>$Global:FinalDestSize</td>
            </tr>
            <tr>
                <td style='font-weight: bold; border: 1px solid #cccccc;'>Robocopy Performance:</td>
                <td style='border: 1px solid #cccccc; font-family: Consolas, monospace; background-color: #f7f7f7;'>$Global:TransferMetrics</td>
            </tr>
            <tr style='background-color: #f9f9f9;'>
                <td style='font-weight: bold; border: 1px solid #cccccc;'>Next Day Provisioning:</td>
                <td style='border: 1px solid #cccccc;'>$NextSequenceTarget</td>
            </tr>
        </tbody>
    </table>
    
    <h3 style='color: #d9534f; margin-top: 25px; margin-bottom: 5px;'>Purged Assets (Retention Cleaned)</h3>
    <ul style='padding-left: 20px; font-size: 14px; color: #555;'>
        $PurgedHtml
    </ul>
    
    <h3 style='color: #0275d8; margin-top: 25px; margin-bottom: 5px;'>Maintained Active Assets</h3>
    <ul style='padding-left: 20px; font-size: 14px; color: #555;'>
        $MaintainedHtml
    </ul>
    
    <hr style='border: 0; border-top: 1px solid #e0e0e0; margin-top: 30px;'>
    <div style='font-size: 12px; color: #888;'>
        Automated Task Scheduler Pipeline Engine Transaction.<br>
        <b>Execution Timestamp:</b> $Timestamp
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
Write-Report "Source Folder Configured: $SourceFolder"
Write-Report "Destination Path Configured: $DestinationFolder"

# 1. LIVE REVALIDATION GATES
if (-not (Test-Path $SourceFolder)) {
    $ErrMsg = "Source data directory missing: $SourceFolder. Execution aborted."
    Write-Report "CRITICAL VALIDATION FAILURE: $ErrMsg"
    Send-StructuralEmail -Status "ERROR" -Subject "❌ REPLICATION CRITICAL FAILURE - Source Missing" -FailedStage "Source Verification" -ErrorException $ErrMsg
    Exit
}

# 2. SMB CONNECTIVITY AUTHENTICATION
try {
    Write-Report "AUTHENTICATION: Purging overlapping network connections..."
    $ServerUNC = ($DestinationRoot -split '\\')[2]
    
    net use | Out-String | ForEach-Object {
        if ($_ -match "\\\\$ServerUNC") {
            Invoke-Expression "net use `"$DestinationRoot`" /delete /y" 2>$null | Out-Null
        }
    }
    
    Write-Report "AUTHENTICATION: Mapping secure system credential context for Destination Share..."
    $NetUseCmd = "net use `"$DestinationRoot`" `"$NetPass`" /USER:`"$NetUser`" /PERSISTENT:NO"
    $NetOutput = Invoke-Expression $NetUseCmd 2>&1 | Out-String
    
    if (-not (Test-Path $DestinationRoot)) { throw "Network path unreachable via security validation gate. Output: $NetOutput" }
    Write-Report "SUCCESS: Authenticated and verified destination share path connection."
} catch {
    $ErrMsg = "Network target share destination not accessible: $DestinationRoot"
    Write-Report "CRITICAL AUTHENTICATION FAILURE: $ErrMsg"
    Send-StructuralEmail -Status "ERROR" -Subject "❌ REPLICATION CRITICAL FAILURE - Network Authentication" -FailedStage "Authentication" -ErrorException "$ErrMsg (Internal: $_)"
    Exit
}

# 3. DESTINATION RETENTION ENFORCEMENT ENGINE (MAX 3 DAYS HELD)
Write-Report "RETENTION: Auditing target share history folder constraints..."
$Now = Get-Date
$NetworkFolders = Get-ChildItem -Path $DestinationRoot -Directory | Where-Object { $_.Name -like "Backup-$Prefix*" }

foreach ($Folder in $NetworkFolders) {
    $Age = [Math]::Floor(($Now - $Folder.CreationTime).TotalDays)
    if ($Age -ge $BackupRetentionDays) {
        try {
            $TargetToPurge = $Folder.FullName
            Write-Report "RETENTION PURGE: Deleting expired historical network directory: $($Folder.Name) (Age: $Age Days)"
            Remove-Item -Path $TargetToPurge -Recurse -Force -Confirm:$false
            $Global:PurgedAssets.Add("Network Storage Expired Container Deleted: $($Folder.Name) ($Age days old)")
        } catch {
            Write-Report "RETENTION WARNING: Failed to remove expired network target directory: $($Folder.FullName). Error: $_"
        }
    } else {
        Write-Report "RETENTION MAINTENANCE: Keeping active backup history: $($Folder.Name) (Age: $Age Days)"
        $Global:MaintainedAssets.Add("Network Storage Active Backup Kept: $($Folder.Name) ($Age days old)")
    }
}

# 4. SECURE RE-PROVISIONING FOR TARGET DESTINATION BLOCK
if (-not (Test-Path $DestinationFolder)) {
    New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null
    Write-Report "PROVISIONING: Created clean target container on remote share: $DestinationFolder"
}

# 5. HIGH-SPEED UNBUFFERED ROBOCOPY EXECUTION SEQUENCE
Write-Report "TRANSFER: Launching unbuffered multi-threaded direct folder replication..."
$RoboArgs = @("`"$SourceFolder`"", "`"$DestinationFolder`"", "/E", "/IS", "/MT:8", "/J", "/COPY:DT", "/R:1", "/W:2", "/V", "/TS", "/FP", "/LOG:`"$ReportLog`"")

$StartTime = Get-Date
Write-Report "TRANSFER PROTOCOL ENGAGED: Monitoring raw stream lines..."

$RoboProcess = Start-Process -FilePath "robocopy.exe" -ArgumentList $RoboArgs -PassThru -NoNewWindow -Wait

$EndTime = Get-Date
$Elapsed = $EndTime - $StartTime
$ElapsedFormatted = "{0:hh\:mm\:ss}" -f $Elapsed

# Interpret standard Robocopy exit code metrics (Codes under 8 indicate non-error operations)
if ($RoboProcess.ExitCode -lt 8) {
    Write-Report "TRANSFER SUCCESS: Robocopy replication completed successfully. Exit Code: $($RoboProcess.ExitCode)"
    Write-Report "TOTAL TIME ELAPSED: $ElapsedFormatted"
} else {
    $ErrMsg = "Robocopy structural replication failed with exit token code: $($RoboProcess.ExitCode)"
    Write-Report "CRITICAL DATA COPY ERROR: $ErrMsg"
    Send-StructuralEmail -Status "ERROR" -Subject "❌ REPLICATION CRITICAL FAILURE - Copy Interrupted" -FailedStage "Robocopy Processing" -ErrorException $ErrMsg
    Exit
}

# 6. DOUBLE-VERIFICATION REPLICATION VALIDATION GATE
Write-Report "VERIFICATION: Initiating explicit structural size synchronization validation check..."
$SourceSizeB = (Get-ChildItem $SourceFolder -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
$DestSizeB   = (Get-ChildItem $DestinationFolder -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum

$Global:FinalDestSize = "$([Math]::Round($DestSizeB / 1GB, 2)) GB"

# FIXED: Nested evaluation removed to prevent runtime parser failure
$MinutesDivisor = if ($Elapsed.TotalMinutes -gt 0) { $Elapsed.TotalMinutes } else { 1 }
$SpeedMetricMBMin = [Math]::Round(($DestSizeB / 1MB) / $MinutesDivisor, 2)

$Global:TransferMetrics = "Elapsed Time: $ElapsedFormatted | Average Wire Speed: $SpeedMetricMBMin MB/min"

Write-Report "METRICS AUDIT: Source Size = $SourceSizeB Bytes | Destination Size = $DestSizeB Bytes"

if ($SourceSizeB -eq $DestSizeB -and $SourceSizeB -gt 0) {
    Write-Report "VERIFICATION SUCCESS: Structural byte-match validated perfectly. 100% network sync verified."
    
    # 7. SAFE LOCAL SOURCE PURGE
    Write-Report "CLEANUP: Safe confirmation flags matched. Purging local database source directory..."
    try {
        Remove-Item -Path $SourceFolder -Recurse -Force -Confirm:$false
        Write-Report "CLEANUP SUCCESS: Raw source directory completely erased from local production block storage."
        $Global:PurgedAssets.Add("Local Source Raw Directory Wiped: $SourceFolder")
    } catch {
        Write-Report "CLEANUP WARNING: Storage array lock prevented source purge: $_"
    }
} else {
    $ErrMsg = "Size validation gate rejected data. Source ($SourceSizeB B) does not match Destination ($DestSizeB B)."
    Write-Report "CRITICAL DATA DISCREPANCY: $ErrMsg"
    Send-StructuralEmail -Status "ERROR" -Subject "❌ REPLICATION CRITICAL FAILURE - Byte Size Mismatch" -FailedStage "Post-Copy Verification Gate" -ErrorException $ErrMsg
    Exit
}

# 8. PROVISION NEXT DAY STORAGE CONTAINER PLACEHOLDER
$NextDayDate   = (Get-Date).AddDays(1).ToString("dd-MMM-yyyy")
$NextDayFolder = Join-Path $SourceRoot "$Prefix$NextDayDate"
if (-not (Test-Path $NextDayFolder)) {
    New-Item -ItemType Directory -Path $NextDayFolder -Force | Out-Null
    Write-Report "LIFECYCLE PROVISIONING: Next sequence folder container initialized at: $NextDayFolder"
}

# 9. CENTRAL LOG ROTATION HANDLING ENGINE (7-DAY ROTATION RULES)
Write-Report "LOG RETENTION: Sweeping tracking files for assets older than 7 days..."
$HistoricalLogs = Get-ChildItem -Path $LogRoot -File | Where-Object { $_.Name -like "Replication_Report_*.log" }

foreach ($LogFile in $HistoricalLogs) {
    $LogAge = [Math]::Floor(($Now - $LogFile.CreationTime).TotalDays)
    if ($LogAge -ge $LogRetentionDays) {
        try {
            $LogPathToPurge = $LogFile.FullName
            Remove-Item -Path $LogPathToPurge -Force -Confirm:$false
            Write-Report "LOG PURGE: Rotated out old infrastructure log sheet: $($LogFile.Name)"
            $Global:PurgedAssets.Add("Local System Engine Log File Cleaned: $($LogFile.Name)")
        } catch {
            Write-Report "LOG WARNING: Unable to clean historical log asset: $($LogFile.Name)"
        }
    } else {
        $Global:MaintainedAssets.Add("Local Staging Log Maintained: $($LogFile.Name) ($LogAge days old)")
    }
}

# 10. TRANSMIT OPERATIONAL NOTIFICATION SUCCESS
Send-StructuralEmail -Status "SUCCESS" -Subject "✅ DB BACKUP REPLICATION SUCCESS - Size: $Global:FinalDestSize"
Write-Report "===== BACKUP PIPELINE END ====="