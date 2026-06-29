# ==============================================================================
# 🟩 ZONE 1: MANUAL USER CONFIGURATIONS & INPUTS (CONFIGURATION AREA)
# ==============================================================================
# Change the values in this zone to match your environment. Do not change code below this zone.

# Path to the shared Excel checklist tracker (must already exist with row 1
# headers containing "HOSTNAME--IPADDRESS" for each of the 11 servers,
# matching merged column-block headers, e.g. B1:C1 = "SERVER01--10.0.0.5")
[string]$ExcelTrackerPath = "\\SRV-Storage\Logs\Email_storage_alert_logs\Sample_Excel_Email_Report_Checklist.xlsx"

# Folder for this script's own run log
[string]$TrackerLogDir    = "\\SRV-Storage\Logs\Email_storage_alert_logs"

# Subject text fragments used to identify storage alert emails (matches the
# subjects produced by the per-server scripts)
[string]$SubjectKeyword1  = "CRITICAL STORAGE ALERT"
[string]$SubjectKeyword2  = "STORAGE REPORT"

# Outlook folder to scan (default = Inbox of the default profile)
# Set to $null to use the default Inbox, or specify a sub-folder name.
[string]$OutlookSubFolder = $null


# ==============================================================================
# 🟦 ZONE 2: RUNTIME CORE LOGIC (DO NOT EDIT)
# ==============================================================================

# --- SETUP RUN LOG ---
if (-not (Test-Path $TrackerLogDir)) {
    New-Item -Path $TrackerLogDir -ItemType Directory -Force | Out-Null
}
$RunDateStamp = Get-Date -Format "ddMMyyyy"
$TrackerLogFile = "$TrackerLogDir\Excel_Tracker_Update--$RunDateStamp.log"

function Write-TrackerLog([string]$Message) {
    $LogStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$LogStamp] $Message" | Out-File -FilePath $TrackerLogFile -Append -Encoding utf8
}

Write-TrackerLog "--------------------------------------------------------"
Write-TrackerLog "PROCESS START: Daily Excel Tracker Update from Outlook Mailbox."

# --- CONNECT TO OUTLOOK ---
$Outlook   = $null
$Namespace = $null
$Folder    = $null

try {
    $Outlook   = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $Folder    = $Namespace.GetDefaultFolder(6)   # 6 = olFolderInbox

    if ($OutlookSubFolder) {
        $Folder = $Folder.Folders.Item($OutlookSubFolder)
    }

    Write-TrackerLog "CONNECT: Outlook mailbox connected. Scanning folder '$($Folder.Name)'."
} catch {
    Write-TrackerLog "❌ FAILED: Could not connect to Outlook. $_"
    Write-TrackerLog "PROCESS END: Aborted due to Outlook connection failure."
    return
}

# --- COLLECT TODAY'S ALERT EMAILS ---
$Today = (Get-Date).Date
$Items = $Folder.Items
$Items.Sort("[ReceivedTime]", $true)   # newest first

$AlertEmails = @()
foreach ($mail in $Items) {
    try {
        if ($mail.ReceivedTime.Date -ne $Today) {
            # Items are sorted newest-first; once older than today, stop scanning
            if ($mail.ReceivedTime.Date -lt $Today) { break }
            else { continue }
        }
        if ($mail.Subject -match [regex]::Escape($SubjectKeyword1) -or $mail.Subject -match [regex]::Escape($SubjectKeyword2)) {
            $AlertEmails += $mail
        }
    } catch {
        continue
    }
}

Write-TrackerLog "SCAN: Found $($AlertEmails.Count) alert email(s) received today."

if ($AlertEmails.Count -eq 0) {
    Write-TrackerLog "PROCESS END: No alert emails found for today. Nothing to update."
    return
}

# --- PARSE DATA FROM EACH EMAIL BODY (HTML) ---
# Expected HTML fragment from the per-server script:
#   • Hostname: SERVER01
#   • System IP: 10.0.0.5
#   <tr ...><td ...>C:\</td><td ...>123.45 GB</td> ... </tr>
#   <tr ...><td ...>D:\</td><td ...>67.89 GB</td> ... </tr>

$ParsedEntries = @()

foreach ($mail in $AlertEmails) {
    try {
        $body = $mail.HTMLBody

        # Extract hostname
        $hostMatch = [regex]::Match($body, "Hostname:\s*</?[^>]*>?\s*([^<\r\n]+)")
        if (-not $hostMatch.Success) {
            $hostMatch = [regex]::Match($body, "Hostname:\s*([^<\r\n]+)")
        }
        $parsedHost = if ($hostMatch.Success) { $hostMatch.Groups[1].Value.Trim() } else { $null }

        # Extract IP address
        $ipMatch = [regex]::Match($body, "System IP:\s*</?[^>]*>?\s*([^<\r\n]+)")
        if (-not $ipMatch.Success) {
            $ipMatch = [regex]::Match($body, "System IP:\s*([^<\r\n]+)")
        }
        $parsedIp = if ($ipMatch.Success) { $ipMatch.Groups[1].Value.Trim() } else { $null }

        if (-not $parsedHost -or -not $parsedIp) {
            Write-TrackerLog "⚠️ SKIP: Could not parse hostname/IP from email '$($mail.Subject)' received $($mail.ReceivedTime)."
            continue
        }

        $hostIpKey = "$parsedHost--$parsedIp"

        # Extract C: and D: drive Free GB values from table rows
        $cMatch = [regex]::Match($body, "C:\\<\/td>\s*<td[^>]*>([\d\.]+)\s*GB")
        $dMatch = [regex]::Match($body, "D:\\<\/td>\s*<td[^>]*>([\d\.]+)\s*GB")

        $cValue = if ($cMatch.Success) { $cMatch.Groups[1].Value } else { $null }
        $dValue = if ($dMatch.Success) { $dMatch.Groups[1].Value } else { $null }

        $ParsedEntries += [PSCustomObject]@{
            HostIpKey = $hostIpKey
            Hostname  = $parsedHost
            IPAddress = $parsedIp
            CDriveGB  = $cValue
            DDriveGB  = $dValue
            Received  = $mail.ReceivedTime
        }

        Write-TrackerLog "PARSED: '$hostIpKey' -> C: $cValue GB, D: $dValue GB (Subject: $($mail.Subject))"
    } catch {
        Write-TrackerLog "⚠️ SKIP: Error parsing email '$($mail.Subject)'. $_"
        continue
    }
}

if ($ParsedEntries.Count -eq 0) {
    Write-TrackerLog "PROCESS END: No valid entries parsed from today's emails."
    return
}

# Deduplicate: keep only the most recent email per host--IP (in case multiple
# emails arrived from the same server today)
$ParsedEntries = $ParsedEntries | Group-Object HostIpKey | ForEach-Object {
    $_.Group | Sort-Object Received -Descending | Select-Object -First 1
}

Write-TrackerLog "DEDUP: $($ParsedEntries.Count) unique server entr(y/ies) ready to write."


# ==============================================================================
# 🟨 ZONE 3: EXCEL CHECKLIST UPDATE VIA EXCEL COM AUTOMATION
# ==============================================================================

$excelApp  = $null
$workbook  = $null
$worksheet = $null

try {
    if (-not (Test-Path $ExcelTrackerPath)) {
        Write-TrackerLog "❌ FAILED: Excel tracker file not found at '$ExcelTrackerPath'."
        Write-TrackerLog "PROCESS END: Aborted - tracker file missing."
        return
    }

    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible       = $false
    $excelApp.DisplayAlerts = $false

    $workbook = $excelApp.Workbooks.Open($ExcelTrackerPath)

    $MonthSheet = (Get-Date).ToString("MMMM")   # e.g. "June", "July"
    $DayOfMonth = (Get-Date).Day
    $RowNumber  = $DayOfMonth + 2               # Row 3 = day 1 of the month

    $worksheet = $null
    foreach ($ws in $workbook.Worksheets) {
        if ($ws.Name -eq $MonthSheet) { $worksheet = $ws; break }
    }
    if (-not $worksheet) {
        Write-TrackerLog "⚠️ Worksheet '$MonthSheet' not found. Creating it as a copy of layout."
        $worksheet = $workbook.Worksheets.Add()
        $worksheet.Name = $MonthSheet
        $daysInMonth = [DateTime]::DaysInMonth((Get-Date).Year, (Get-Date).Month)
        for ($d = 1; $d -le $daysInMonth; $d++) {
            $worksheet.Cells.Item($d + 2, 1) = (Get-Date -Year (Get-Date).Year -Month (Get-Date).Month -Day $d)
        }
    }

    # Read row 1 headers to build a HostIpKey -> starting column index map
    # Headers are merged cells (e.g. B1:C1), so read the value from each
    # block's first cell. Column blocks run B,D,F,H,J,L,N,P,R,T,V (cols 2-22, step 2)
    $HeaderColumnMap = @{}
    for ($col = 2; $col -le 22; $col += 2) {
        $headerVal = $worksheet.Cells.Item(1, $col).Text
        if (-not [string]::IsNullOrWhiteSpace($headerVal)) {
            $HeaderColumnMap[$headerVal.Trim()] = $col
        }
    }

    Write-TrackerLog "HEADERS: Found $($HeaderColumnMap.Count) registered server header(s) in '$MonthSheet' row 1."

    # Track which registered servers got an entry today (everyone else stays
    # empty/red, as those servers did not send an email)
    $UpdatedKeys = @()

    foreach ($entry in $ParsedEntries) {
        if ($HeaderColumnMap.ContainsKey($entry.HostIpKey)) {
            $colC = $HeaderColumnMap[$entry.HostIpKey]       # C Drive column
            $colD = $colC + 1                                 # D Drive column

            $cellC = $worksheet.Cells.Item($RowNumber, $colC)
            $cellD = $worksheet.Cells.Item($RowNumber, $colD)

            if ($entry.CDriveGB) { $cellC.Value2 = "$($entry.CDriveGB) GB" } else { $cellC.Value2 = "" }
            if ($entry.DDriveGB) { $cellD.Value2 = "$($entry.DDriveGB) GB" } else { $cellD.Value2 = "" }

            # Clear any prior red highlight (data successfully received)
            $cellC.Interior.ColorIndex = -4142   # xlNone
            $cellD.Interior.ColorIndex = -4142
            $cellC.Font.Color = 0
            $cellD.Font.Color = 0

            $UpdatedKeys += $entry.HostIpKey

            Write-TrackerLog "WRITE: '$($entry.HostIpKey)' -> $MonthSheet!$($cellC.Address($false,$false))=$($entry.CDriveGB) GB, $($cellD.Address($false,$false))=$($entry.DDriveGB) GB"
        } else {
            Write-TrackerLog "⚠️ UNREGISTERED: '$($entry.HostIpKey)' does not match any header in row 1. Skipped (logged only)."
        }
    }

    # --- Mark missing servers (no email today) as empty + RED for today's row ---
    foreach ($key in $HeaderColumnMap.Keys) {
        if ($UpdatedKeys -notcontains $key) {
            $colC = $HeaderColumnMap[$key]
            $colD = $colC + 1

            $cellC = $worksheet.Cells.Item($RowNumber, $colC)
            $cellD = $worksheet.Cells.Item($RowNumber, $colD)

            $cellC.Value2 = ""
            $cellD.Value2 = ""
            $cellC.Interior.Color = 255    # Red
            $cellD.Interior.Color = 255    # Red

            Write-TrackerLog "MISSING: '$key' did not send an alert email today. Marked $MonthSheet!$($cellC.Address($false,$false)):$($cellD.Address($false,$false)) as RED/EMPTY."
        }
    }

    $workbook.Save()
    $workbook.Close($true)
    $excelApp.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)  | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp)  | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-TrackerLog "SUCCESS: Excel checklist tracker updated for $($MonthSheet) row $RowNumber."
} catch {
    Write-TrackerLog "❌ FAILED: Excel update encountered an error. $_"

    try {
        if ($workbook)  { $workbook.Close($false) }
        if ($excelApp)  { $excelApp.Quit() }
        if ($worksheet) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null }
        if ($workbook)  { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)  | Out-Null }
        if ($excelApp)  { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp)  | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    } catch {}
}

Write-TrackerLog "PROCESS END: Daily Excel Tracker Update complete."
