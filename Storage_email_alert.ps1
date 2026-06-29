# ==============================================================================
# ðŸŸ© ZONE 1: MANUAL USER CONFIGURATIONS & INPUTS (CONFIGURATION AREA)
# ==============================================================================
# Change the values in this zone to match your environment. Do not change code below this zone.

# --- PERSONAL OUTLOOK ACCREDITATION SETTINGS ---
# --- ALERTS INPUT CONFIGURATIONS (DECLARED CREDENTIAL VARIABLES) ---
[string]$SmtpServer          = "smtp-mail.outlook.com"
[int]$SmtpPort               = 587
[string]$EmailFrom           = "senderemail@infoxtek.com"   # ðŸ‘ˆ Your Outlook/Hotmail address
[string]$HotmailPlainPassword= "YourSmtppassword"               # ðŸ‘ˆ Your 16-Character App Password



# 📬 GENERAL / DAILY REPORTS RECEIVERS (Always on the "To:" line)
[string]$EmailTo             = "admin1@infoxtek.com , admin2@infoxtek.com , support1@infoxtek.com , support2@infoxtek.com"  

# 👥 CRITICAL ALERT CC RECEIVERS (Only added to "CC:" line when storage is low)
[string]$CriticalEmailCc     = "Manager1@infoxtek.com , Manager2@infoxtek.com"

[double]$CriticalThresholdGB = 10.0                                # Safety boundary threshold

# --- NEW DIAGNOSTIC EMAIL LOG CONFIGURATIONS ---
$EmailLogDir                 = "\\SRV-Storage\Logs\Email_storage_alert_logs"


# ==============================================================================
# 🟦 ZONE 2: RUNTIME CORE LOGIC (DO NOT EDIT)
# ==============================================================================

# --- GET HOST CONTEXT & DATE STAMP ---
$hostname = $env:COMPUTERNAME
$ipAddress = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
    $_.IPAddress -notlike '169.*' -and $_.InterfaceAlias -notlike 'Loopback*' -and $_.PrefixOrigin -ne 'WellKnown'
} | Select-Object -First 1 -ExpandProperty IPAddress)

# Generate custom DDMMYYYY date format
$dateStamp = Get-Date -Format "ddMMyyyy"

# Dynamically construct the trace log filename using the exact format requested
$EmailLogFile = "$EmailLogDir\$hostname--$ipAddress-$dateStamp.log"

# Ensure logging directory path exists
if (-not (Test-Path $EmailLogDir)) {
    New-Item -Path $EmailLogDir -ItemType Directory -Force | Out-Null
}

# --- RETENTION POLICY SUB-SYSTEM (7 DAYS) ---
# Finds and purges logs matching this host's pattern older than 7 days before creating the new one
$retentionCutoff = (Get-Date).AddDays(-7).Date
if (Test-Path $EmailLogDir) {
    Get-ChildItem -Path $EmailLogDir -File | Where-Object {
        $_.Name -like "$hostname--$ipAddress-*.log" -and $_.LastWriteTime -lt $retentionCutoff
    } | ForEach-Object {
        try {
            Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue
        } catch {
            # Fail silently to prevent interrupting the core monitoring pipeline
        }
    }
}

function Write-EmailLog([string]$Message) {
    $LogStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$LogStamp] $Message" | Out-File -FilePath $EmailLogFile -Append -Encoding utf8
}

Write-EmailLog "--------------------------------------------------------"
Write-EmailLog "PROCESS START: Initiating System Space Verification Pipeline."

# --- PROCESS DRIVE METRICS AND CAPACITY ---
$DrivesToAudit = @('C', 'D')
$CriticalDrivesDetected = @()
$HtmlDriveRows = ""

foreach ($Letter in $DrivesToAudit) {
    if (Get-PSDrive -Name $Letter -ErrorAction SilentlyContinue) {
        $DriveObj = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$($Letter):'"
        [double]$FreeGB      = [math]::Round($DriveObj.FreeSpace / 1GB, 2)
        [double]$SizeGB      = [math]::Round($DriveObj.Size / 1GB, 2)
        [double]$PercentFree = [math]::Round(($FreeGB / $SizeGB) * 100, 1)

        # Determine individual drive status styling
        if ($FreeGB -lt $CriticalThresholdGB) {
            $StatusStyle = "color: #FF0000; font-weight: bold; background-color: #FCE4D6;"
            $StatusText  = "CRITICAL ALERT"
            $CriticalDrivesDetected += "${Letter}:\ ($FreeGB GB left)"
        } else {
            $StatusStyle = "color: #008000; font-weight: normal; background-color: #E2EFDA;"
            $StatusText  = "HEALTHY"
        }

        $HtmlDriveRows += @"
        <tr style="$StatusStyle">
            <td style="padding: 8px; border: 1px solid #ddd; text-align: center; font-weight: bold;">${Letter}:\</td>
            <td style="padding: 8px; border: 1px solid #ddd; text-align: right; font-weight: bold;">$FreeGB GB</td>
            <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">$SizeGB GB</td>
            <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">$PercentFree %</td>
            <td style="padding: 8px; border: 1px solid #ddd; text-align: center; font-weight: bold;">$StatusText</td>
        </tr>
"@
    }
}

# --- INITIALIZE ROUTING STRING CONTAINERS ---
$FinalCcRouting = ""

# --- DYNAMIC SUBJECT & DESIGN LOOKUP ---
if ($CriticalDrivesDetected.Count -gt 0) {
    $AlertSubject = "⚠️ CRITICAL STORAGE ALERT: Server [$hostname] is running low on disk space!"
    $HeaderColor  = "#d9534f" # Red header
    $StatusTitle  = "Critical Storage Alert - System Action Required"
    $IntroText    = "The following storage volumes remain **below** the safety threshold of <strong>$CriticalThresholdGB GB</strong>:"
    
    # Map critical CC audience
    $FinalCcRouting = $CriticalEmailCc
} else {
    $AlertSubject = "✅ STORAGE REPORT: Server [$hostname] status is Healthy"
    $HeaderColor  = "#28a745" # Green header
    $StatusTitle  = "Routine Storage Status Report"
    $IntroText    = "All storage volumes are operating safely **above** the safety threshold of <strong>$CriticalThresholdGB GB</strong>:"
    
    # Leave CC routing blank for routine reports
    $FinalCcRouting = ""
}

# --- GENERATE HTML EMAIL ---
$EmailBody = @"
<html>
<body style="font-family: Arial, sans-serif; color: #333; line-height: 1.5;">
    <div style="background-color: $HeaderColor; color: white; padding: 15px; font-size: 18px; font-weight: bold; border-radius: 4px 4px 0 0;">
        $StatusTitle
    </div>
    <div style="padding: 20px; border: 1px solid $HeaderColor; border-top: none; background-color: #fff;">
        <p>Hello Administrator,</p>
        <p>$IntroText</p>
        
        <table style="width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px;">
            <thead>
                <tr style="background-color: #f2f2f2; border-bottom: 2px solid #ddd;">
                    <th style="padding: 10px; border: 1px solid #ddd; text-align: center;">Drive</th>
                    <th style="padding: 10px; border: 1px solid #ddd; text-align: right;">Available Space</th>
                    <th style="padding: 10px; border: 1px solid #ddd; text-align: right;">Total Capacity</th>
                    <th style="padding: 10px; border: 1px solid #ddd; text-align: right;">% Free Space</th>
                    <th style="padding: 10px; border: 1px solid #ddd; text-align: center;">System Status</th>
                </tr>
            </thead>
            <tbody>
                $HtmlDriveRows
            </tbody>
        </table>

        <div style="margin-top: 25px; background-color: #f9f9f9; padding: 15px; border-left: 4px solid #5bc0de; font-size: 12px; color: #666;">
            <strong>Execution Server Context:</strong><br>
            • Hostname: $hostname<br>
            • System IP: $ipAddress<br>
            • Audit Execution Timestamp: $(Get-Date -Format "MM/dd/yyyy HH:mm:ss")
        </div>
    </div>
</body>
</html>
"@

# --- OUTBOUND MAILING SUB-SYSTEM ---
try {
    Write-EmailLog "CONNECT: Initializing specialized .NET secure mail pipeline client wrapper..."
    
    $Mail = New-Object System.Net.Mail.MailMessage
    $Mail.From = New-Object System.Net.Mail.MailAddress($EmailFrom)
    
    # Process multiple receivers correctly by splitting on commas and trimming out whitespace
    $EmailTo.Split(',') | ForEach-Object {
        $cleanEmail = $_.Trim()
        if (-not [string]::IsNullOrEmpty($cleanEmail)) {
            $Mail.To.Add($cleanEmail)
        }
    }

    # Process "CC" recipients dynamically if a critical storage issue was logged
    if (-not [string]::IsNullOrWhiteSpace($FinalCcRouting)) {
        Write-EmailLog "ROUTING: Critical threshold compromised. Populating CC routing configurations..."
        $FinalCcRouting.Split(',') | ForEach-Object {
            $cleanCc = $_.Trim()
            if (-not [string]::IsNullOrEmpty($cleanCc)) {
                $Mail.CC.Add($cleanCc)
            }
        }
    }

    # Force explicit UTF8 Encoding properties on the .NET Mail Object
    $Mail.SubjectEncoding = [System.Text.Encoding]::UTF8
    $Mail.BodyEncoding    = [System.Text.Encoding]::UTF8

    $Mail.Subject = $AlertSubject
    $Mail.Body = $EmailBody
    $Mail.IsBodyHtml = $true

    # Setup custom modern client submission stack parameters explicitly
    $Smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
    $Smtp.EnableSsl = $true
    $Smtp.TargetName = "STARTTLS/smtp-mail.outlook.com" 
    $Smtp.Credentials = New-Object System.Net.NetworkCredential($EmailFrom, $HotmailPlainPassword)

    Write-EmailLog "CONNECT: Dispatching transmission through explicit target socket tunnel..."
    $Smtp.Send($Mail)
    
    # Dispose memory allocations
    $Mail.Dispose()
    $Smtp.Dispose()
    
    Write-EmailLog "SUCCESS: Personal Outlook transmission cleared verification. Dispatch achieved."
} catch {
    Write-EmailLog "❌ FAILED: Pipeline compilation failed to push data to Outlook gateway."
    Write-EmailLog "EXCEPTIONAL ERROR EXPLANATION: $_"
    if ($_.Exception.InnerException) {
        Write-EmailLog "INNER CORE EXCEPTION DETAILS: $($_.Exception.InnerException.Message)"
    }
}