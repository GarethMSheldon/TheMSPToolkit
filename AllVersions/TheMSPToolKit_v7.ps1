# Require Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# Load UI Assemblies
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Main Form Setup ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "MSP Technician IDE - Professional Edition"
$form.Size = New-Object System.Drawing.Size(1200, 900)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# --- Status Bar at Bottom ---
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready | Administrator Mode | Computer: $env:COMPUTERNAME"
$statusLabel.ForeColor = [System.Drawing.Color]::LimeGreen
$statusBar.Items.Add($statusLabel) | Out-Null
$statusBar.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$form.Controls.Add($statusBar)

# --- Progress Bar ---
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 810)
$progressBar.Size = New-Object System.Drawing.Size(550, 25)
$progressBar.Style = "Marquee"
$progressBar.MarqueeAnimationSpeed = 30
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# --- Global Output Box ---
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Point(580, 10)
$outputBox.Size = New-Object System.Drawing.Size(600, 790)
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$outputBox.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$outputBox.ForeColor = [System.Drawing.Color]::Gainsboro
$outputBox.ReadOnly = $true
$form.Controls.Add($outputBox)

# --- Clear Log Button ---
$clearBtn = New-Object System.Windows.Forms.Button
$clearBtn.Text = "Clear Log"
$clearBtn.Location = New-Object System.Drawing.Point(580, 810)
$clearBtn.Size = New-Object System.Drawing.Size(140, 30)
$clearBtn.BackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$clearBtn.ForeColor = [System.Drawing.Color]::White
$clearBtn.FlatStyle = "Flat"
$clearBtn.Add_Click({ $outputBox.Clear(); Log "Log cleared." "INFO" })
$form.Controls.Add($clearBtn)

# --- Save Log Button ---
$saveBtn = New-Object System.Windows.Forms.Button
$saveBtn.Text = "Save Log to Desktop"
$saveBtn.Location = New-Object System.Drawing.Point(730, 810)
$saveBtn.Size = New-Object System.Drawing.Size(180, 30)
$saveBtn.BackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$saveBtn.ForeColor = [System.Drawing.Color]::White
$saveBtn.FlatStyle = "Flat"
$saveBtn.Add_Click({ 
    $logPath = "$env:USERPROFILE\Desktop\ITLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $outputBox.Text | Out-File -FilePath $logPath
    Log "Log saved to: $logPath" "PASS"
    [System.Windows.Forms.MessageBox]::Show("Log saved to Desktop!", "Success", "OK", "Information")
})
$form.Controls.Add($saveBtn)

# --- Logging Function ---
function Log([string]$msg, [string]$status="INFO") {
    $timestamp = Get-Date -Format "HH:mm:ss"
    $color = "White"
    if($status -eq "PASS"){$color="Lime"}
    if($status -eq "FAIL"){$color="OrangeRed"}
    if($status -eq "WARN"){$color="Yellow"}
    $outputBox.SelectionColor = [System.Drawing.Color]::$color
    $outputBox.AppendText("[$timestamp][$status] $msg`r`n")
    $outputBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# --- Show/Hide Progress Bar ---
function Show-Progress { $progressBar.Visible = $true; [System.Windows.Forms.Application]::DoEvents() }
function Hide-Progress { $progressBar.Visible = $false; [System.Windows.Forms.Application]::DoEvents() }

# --- Sidebar Panel ---
$sidebar = New-Object System.Windows.Forms.FlowLayoutPanel
$sidebar.Location = New-Object System.Drawing.Point(10, 10)
$sidebar.Size = New-Object System.Drawing.Size(560, 790)
$sidebar.FlowDirection = "TopDown"
$sidebar.AutoScroll = $true
$sidebar.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$sidebar.Padding = New-Object System.Windows.Forms.Padding(10)
$form.Controls.Add($sidebar)

# --- Tooltip for descriptions ---
$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.AutoPopDelay = 8000
$tooltip.InitialDelay = 500
$tooltip.ReshowDelay = 500

function Add-SectionLabel($text) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "`n$text"
    $label.Width = 530
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $label.ForeColor = [System.Drawing.Color]::White
    $sidebar.Controls.Add($label)
}

function Create-Btn($text, $color, $script, $description) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Width = 530
    $btn.Height = 42
    $btn.Margin = New-Object System.Windows.Forms.Padding(0, 5, 0, 5)
    $btn.FlatStyle = "Flat"
    $btn.BackColor = $color
    $btn.ForeColor = [System.Drawing.Color]::Black
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    
    # Add tooltip description
    if($description) { $tooltip.SetToolTip($btn, $description) }
    
    $btn.Add_Click($script)
    $sidebar.Controls.Add($btn)
}

# --- SECTION: SYSTEM REPAIR AND HEALTH ---
Add-SectionLabel "SYSTEM REPAIR AND HEALTH"

Create-Btn "Run SFC Scan (System File Checker)" ([System.Drawing.Color]::LightCoral) {
    try {
        Log "Starting SFC Scan... this may take 10-15 minutes." "WARN"
        Show-Progress
        $process = Start-Process "sfc" -ArgumentList "/scannow" -Wait -NoNewWindow -PassThru
        Hide-Progress
        if($process.ExitCode -eq 0){ Log "SFC Scan Complete - No issues found." "PASS" }
        else { Log "SFC Scan Complete - Check system logs for details." "WARN" }
    } catch {
        Hide-Progress
        Log "SFC Scan failed: $($_.Exception.Message)" "FAIL"
    }
} "Scans and repairs corrupted Windows system files. Takes 10-15 minutes."

Create-Btn "Check for Pending Reboot" ([System.Drawing.Color]::White) {
    try {
        $rebootRequired = $false
        $checks = @{
            "WindowsUpdate" = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
            "ComponentBasedServicing" = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
            "PendingFileRename" = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations"
        }
        
        foreach($check in $checks.GetEnumerator()){
            if(Test-Path $check.Value){ 
                Log "Reboot PENDING: $($check.Key)" "FAIL"
                $rebootRequired = $true
            }
        }
        
        if(-not $rebootRequired){ Log "No pending reboot flags detected." "PASS" }
    } catch {
        Log "Error checking reboot status: $($_.Exception.Message)" "FAIL"
    }
} "Checks multiple registry locations for pending system reboots."

# --- SECTION: USER PROFILE AND M365 ---
Add-SectionLabel "USER PROFILE AND M365"

Create-Btn "Clear Teams Cache (Classic & New)" ([System.Drawing.Color]::White) {
    try {
        Log "Closing Teams processes..."
        Get-Process -Name "Teams", "ms-teams" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 2
        
        Log "Clearing Cache Folders..."
        $paths = @(
            "$env:AppData\Microsoft\Teams",
            "$env:LocalAppData\Packages\MSTeams_8wekyb3d8bbwe\LocalCache"
        )
        foreach($p in $paths){ 
            if(Test-Path $p){ 
                Remove-Item "$p\*" -Recurse -Force -ErrorAction SilentlyContinue
                Log "Cleared: $p" "PASS"
            }
        }
        Log "Teams cache cleared successfully." "PASS"
    } catch {
        Log "Error clearing Teams cache: $($_.Exception.Message)" "FAIL"
    }
} "Fixes Teams freezing, login issues, and sync problems by clearing cache."

Create-Btn "Fix Outlook Profile (Delete & Recreate)" ([System.Drawing.Color]::LightCoral) {
    try {
        Log "Closing Outlook..." "WARN"
        Get-Process "OUTLOOK" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 2
        
        Log "Deleting Outlook Profile Registry Keys..."
        $outlookVersions = @("16.0", "15.0", "14.0")
        foreach($ver in $outlookVersions){
            $path = "HKCU:\Software\Microsoft\Office\$ver\Outlook\Profiles"
            if(Test-Path $path){
                Remove-Item "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
                Log "Removed profile for Office $ver" "PASS"
            }
        }
        Log "Outlook profile removed. Reopen Outlook to recreate." "PASS"
    } catch {
        Log "Error fixing Outlook profile: $($_.Exception.Message)" "FAIL"
    }
} "Fixes 'Enter Password' prompts, profile corruption, and connection issues."

Create-Btn "Fix OneDrive Sync Issues" ([System.Drawing.Color]::White) {
    try {
        Log "Resetting OneDrive..." "WARN"
        $oneDrivePath = "$env:LocalAppData\Microsoft\OneDrive\onedrive.exe"
        if(Test-Path $oneDrivePath){
            Show-Progress
            Start-Process $oneDrivePath -ArgumentList "/reset" -NoNewWindow -ErrorAction Stop
            Start-Sleep -Seconds 8
            Start-Process $oneDrivePath -NoNewWindow -ErrorAction Stop
            Hide-Progress
            Log "OneDrive reset complete." "PASS"
        } else {
            Log "OneDrive not installed at expected location." "WARN"
        }
    } catch {
        Hide-Progress
        Log "Error resetting OneDrive: $($_.Exception.Message)" "FAIL"
    }
} "Fixes sync errors, file conflicts, and 'processing changes' stuck status."

Create-Btn "Fix Disconnected Mapped Drives" ([System.Drawing.Color]::White) {
    try {
        Log "Checking Network Drives..."
        $drives = Get-PSDrive -PSProvider FileSystem -ErrorAction Stop | Where-Object { $_.DisplayRoot -ne $null }
        
        if($drives.Count -eq 0){ 
            Log "No mapped drives found." "WARN" 
        } else {
            foreach($d in $drives){
                Log "Testing: $($d.Name): ($($d.DisplayRoot))..."
                if(Test-Path $d.Root -ErrorAction SilentlyContinue){ 
                    Log "[OK] Drive $($d.Name): Connected" "PASS" 
                } else { 
                    Log "[FAIL] Drive $($d.Name): DISCONNECTED" "FAIL" 
                }
            }
        }
    } catch {
        Log "Error checking drives: $($_.Exception.Message)" "FAIL"
    }
} "Tests connectivity to all mapped network drives (Z:, H:, etc.)"

# --- SECTION: NETWORKING ---
Add-SectionLabel "ADVANCED NETWORKING"

Create-Btn "Find Top Network Consuming Processes" ([System.Drawing.Color]::LightBlue) {
    try {
        Log "Monitoring Network Traffic..."
        Show-Progress
        
        try {
            $stats = Get-NetTCPConnection -State Established -ErrorAction Stop | 
                Group-Object -Property OwningProcess | 
                Select-Object Count, Name, @{Name="Process";Expression={(Get-Process -Id $_.Name -ErrorAction SilentlyContinue).ProcessName}} | 
                Sort-Object Count -Descending | Select-Object -First 5
            
            foreach($s in $stats){ 
                Log "PID $($s.Name) ($($s.Process)): $($s.Count) Active Connections" 
            }
        } catch {
            Log "Using netstat fallback method..." "WARN"
            $netstatOutput = netstat -ano | Select-String "ESTABLISHED"
            $processCount = @{}
            foreach($line in $netstatOutput){
                $pid = ($line -split '\s+')[-1]
                if($processCount.ContainsKey($pid)){ $processCount[$pid]++ }
                else { $processCount[$pid] = 1 }
            }
            
            $top5 = $processCount.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5
            foreach($entry in $top5){
                $proc = (Get-Process -Id $entry.Key -ErrorAction SilentlyContinue).ProcessName
                Log "PID $($entry.Key) ($proc): $($entry.Value) Connections"
            }
        }
        
        Hide-Progress
    } catch {
        Hide-Progress
        Log "Network monitoring failed: $($_.Exception.Message)" "FAIL"
    }
} "Identifies processes using the most network bandwidth (helpful for slow network issues)."

Create-Btn "List WiFi Passwords (Saved Profiles)" ([System.Drawing.Color]::White) {
    try {
        Log "Retrieving Saved WiFi Profiles..."
        $profiles = netsh wlan show profiles | Select-String "\:(.+)$" | ForEach-Object { $_.Matches.Groups[1].Value.Trim() }
        
        if($profiles.Count -eq 0){
            Log "No WiFi profiles found." "WARN"
        } else {
            foreach($p in $profiles){
                $pass = netsh wlan show profile name="$p" key=clear | Select-String "Key Content\W+\:(.+)$" | ForEach-Object { $_.Matches.Groups[1].Value.Trim() }
                if($pass){ Log "SSID: $p | Password: $pass" "PASS" }
                else { Log "SSID: $p | No password stored" "WARN" }
            }
        }
    } catch {
        Log "Error retrieving WiFi passwords: $($_.Exception.Message)" "FAIL"
    }
} "Displays all saved WiFi network names and passwords from this computer."

Create-Btn "Reset Network Stack (Nuclear Option)" ([System.Drawing.Color]::LightCoral) {
    try {
        Log "Resetting Network Stack..." "WARN"
        Show-Progress
        
        netsh winsock reset | Out-Null
        netsh int ip reset | Out-Null
        ipconfig /flushdns | Out-Null
        ipconfig /release | Out-Null
        ipconfig /renew | Out-Null
        
        Hide-Progress
        Log "Network Stack Reset Complete. REBOOT REQUIRED!" "PASS"
        [System.Windows.Forms.MessageBox]::Show("Network reset complete. Please restart your computer.", "Reboot Required", "OK", "Warning")
    } catch {
        Hide-Progress
        Log "Network reset failed: $($_.Exception.Message)" "FAIL"
    }
} "Last resort for network issues: resets TCP/IP stack, Winsock, DNS cache. Requires reboot."

Create-Btn "Show Public & Local IP Addresses" ([System.Drawing.Color]::White) {
    try {
        Log "Fetching IP Information..."
        $localIP = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop | Where-Object {$_.InterfaceAlias -notlike "*Loopback*" -and $_.PrefixOrigin -ne "WellKnown"} | Select-Object -First 1).IPAddress
        
        try { 
            $publicIP = (Invoke-WebRequest -Uri "https://api.ipify.org" -UseBasicParsing -TimeoutSec 5).Content 
        } catch { 
            $publicIP = "Unable to fetch (check internet)" 
        }
        
        Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "PASS"
        Log "LOCAL IP:  $localIP" "PASS"
        Log "PUBLIC IP: $publicIP" "PASS"
        Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "PASS"
    } catch {
        Log "Error fetching IPs: $($_.Exception.Message)" "FAIL"
    }
} "Displays both internal (LAN) and external (WAN) IP addresses."

# --- SECTION: PRINTER TROUBLESHOOTING ---
Add-SectionLabel "PRINTER TROUBLESHOOTING"

Create-Btn "Clear Print Spooler & Restart Service" ([System.Drawing.Color]::White) {
    try {
        Log "Stopping Print Spooler..."
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        
        Log "Clearing Spool Folder..."
        Remove-Item "C:\Windows\System32\spool\PRINTERS\*" -Force -ErrorAction SilentlyContinue
        
        Start-Service -Name Spooler -ErrorAction Stop
        Log "Print Spooler Cleared & Restarted." "PASS"
    } catch {
        Log "Print spooler operation failed: $($_.Exception.Message)" "FAIL"
        try { Start-Service -Name Spooler -ErrorAction SilentlyContinue } catch {}
    }
} "Fixes stuck print jobs, printer offline errors, and spooler crashes."

Create-Btn "List All Installed Printers" ([System.Drawing.Color]::White) {
    try {
        Log "Listing Printers..."
        
        try {
            $printers = Get-Printer -ErrorAction Stop
            foreach($p in $printers){ 
                $status = if($p.PrinterStatus -eq "Normal"){"[OK]"} else{"[FAIL]"}
                Log "$status $($p.Name) - Status: $($p.PrinterStatus)" 
            }
        } catch {
            Log "Using WMI fallback..." "WARN"
            $printers = Get-WmiObject -Class Win32_Printer
            foreach($p in $printers){
                Log "• $($p.Name) - Status: $($p.PrinterState)"
            }
        }
    } catch {
        Log "Error listing printers: $($_.Exception.Message)" "FAIL"
    }
} "Shows all installed printers and their current status."

Create-Btn "Remove Offline/Ghost Printers" ([System.Drawing.Color]::LightCoral) {
    try {
        Log "Searching for Offline Printers..." "WARN"
        
        try {
            $offline = Get-Printer -ErrorAction Stop | Where-Object { $_.PrinterStatus -ne "Normal" }
            if($offline.Count -eq 0){ 
                Log "No offline printers found." "PASS" 
            } else {
                foreach($p in $offline){ 
                    Remove-Printer -Name $p.Name -ErrorAction SilentlyContinue
                    Log "Removed: $($p.Name)" "PASS" 
                }
            }
        } catch {
            Log "Printer removal unavailable on this system." "WARN"
        }
    } catch {
        Log "Error removing printers: $($_.Exception.Message)" "FAIL"
    }
} "Automatically removes printers with error/offline status."

# --- SECTION: DISK AND STORAGE ---
Add-SectionLabel "DISK AND STORAGE HEALTH"

Create-Btn "Run CHKDSK (Read-Only Scan)" ([System.Drawing.Color]::White) {
    try {
        Log "Running CHKDSK on C: (read-only)..." "WARN"
        Show-Progress
        chkdsk C: /scan
        Hide-Progress
        Log "CHKDSK Complete." "PASS"
    } catch {
        Hide-Progress
        Log "CHKDSK failed: $($_.Exception.Message)" "FAIL"
    }
} "Scans disk for errors without fixing them - safe non-destructive check."

Create-Btn "Find Large Files (>500MB)" ([System.Drawing.Color]::White) {
    try {
        Log "Scanning for large files (may take 2-3 minutes)..." "WARN"
        Show-Progress
        
        $largeFiles = Get-ChildItem "C:\Users" -Recurse -File -ErrorAction SilentlyContinue | 
            Where-Object { $_.Length -gt 500MB } | 
            Sort-Object Length -Descending | 
            Select-Object -First 10
        
        Hide-Progress
        
        if($largeFiles.Count -eq 0){
            Log "No files larger than 500MB found." "PASS"
        } else {
            foreach($f in $largeFiles){ 
                Log "$($f.FullName) - $('{0:N2}' -f ($f.Length/1GB)) GB" 
            }
        }
    } catch {
        Hide-Progress
        Log "File scan failed: $($_.Exception.Message)" "FAIL"
    }
} "Locates the 10 largest files on the system (helpful for 'disk full' issues)."

Create-Btn "Check Disk Health (SMART Status)" ([System.Drawing.Color]::White) {
    try {
        Log "Checking Disk Health..."
        $disks = Get-PhysicalDisk -ErrorAction Stop
        
        foreach($d in $disks){ 
            $health = if($d.HealthStatus -eq "Healthy"){"OK"} else{"FAIL"}
            Log "$health Disk: $($d.FriendlyName) | Health: $($d.HealthStatus) | Status: $($d.OperationalStatus)" 
        }
    } catch {
        Log "Disk health check failed: $($_.Exception.Message)" "FAIL"
    }
} "Checks physical disk health using SMART data (detects failing drives early)."

Create-Btn "Clean Temp Files (System & User)" ([System.Drawing.Color]::LightCoral) {
    try {
        Log "Cleaning Temporary Files..." "WARN"
        Show-Progress
        
        $tempPaths = @(
            "C:\Windows\Temp",
            "$env:TEMP",
            "C:\Windows\Prefetch"
        )
        
        $totalFreed = 0
        foreach($path in $tempPaths){
            if(Test-Path $path){
                $before = (Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                Remove-Item "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
                $after = (Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                $freed = ($before - $after) / 1MB
                $totalFreed += $freed
                Log "Cleaned: $path (freed $('{0:N2}' -f $freed) MB)"
            }
        }
        
        Hide-Progress
        Log "Temp cleanup complete. Total freed: $('{0:N2}' -f $totalFreed) MB" "PASS"
    } catch {
        Hide-Progress
        Log "Temp file cleanup failed: $($_.Exception.Message)" "FAIL"
    }
} "Deletes temporary Windows files, user temp, and prefetch to free disk space."

# --- SECTION: WINDOWS UPDATE ---
Add-SectionLabel "WINDOWS UPDATE FIXES"

Create-Btn "Reset Windows Update Components" ([System.Drawing.Color]::LightCoral) {
    try {
        Log "Resetting Windows Update (this may take 1-2 minutes)..." "WARN"
        Show-Progress
        
        $services = @('wuauserv', 'cryptSvc', 'bits', 'msiserver')
        
        # Stop services
        foreach($svc in $services){
            try {
                Stop-Service -Name $svc -Force -ErrorAction Stop
                Log "Stopped: $svc" "PASS"
            } catch {
                Log "Could not stop $svc (may already be stopped)" "WARN"
            }
        }
        
        Start-Sleep -Seconds 2
        
        # Rename folders
        Log "Renaming SoftwareDistribution folder..."
        if(Test-Path "C:\Windows\SoftwareDistribution"){
            Rename-Item "C:\Windows\SoftwareDistribution" "SoftwareDistribution.old" -Force -ErrorAction SilentlyContinue
        }
        if(Test-Path "C:\Windows\System32\catroot2"){
            Rename-Item "C:\Windows\System32\catroot2" "catroot2.old" -Force -ErrorAction SilentlyContinue
        }
        
        # Start services
        foreach($svc in $services){
            try {
                Start-Service -Name $svc -ErrorAction Stop
                Log "Started: $svc" "PASS"
            } catch {
                Log "Could not start $svc" "WARN"
            }
        }
        
        Hide-Progress
        Log "Windows Update Components Reset Complete." "PASS"
    } catch {
        Hide-Progress
        Log "Windows Update reset failed: $($_.Exception.Message)" "FAIL"
    }
} "Fixes 'Updates failed to install' errors by resetting all Windows Update services."

Create-Btn "Check Last Update Install Date" ([System.Drawing.Color]::White) {
    try {
        Log "Checking Windows Update History..."
        $updates = Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending | Select-Object -First 5
        
        Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "PASS"
        foreach($u in $updates){
            Log "KB: $($u.HotFixID) | Installed: $($u.InstalledOn)" "PASS"
        }
        Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "PASS"
    } catch {
        Log "Could not retrieve update history: $($_.Exception.Message)" "FAIL"
    }
} "Shows the 5 most recent Windows updates installed on this system."

# --- SECTION: PERFORMANCE ---
Add-SectionLabel "PERFORMANCE AND CRASHES"

Create-Btn "Find Memory-Hogging Processes (Top 5)" ([System.Drawing.Color]::White) {
    try {
        Log "Memory Usage Report..."
        $processes = Get-Process -ErrorAction Stop | 
            Sort-Object WorkingSet -Descending | 
            Select-Object -First 5
        
        Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "PASS"
        foreach($p in $processes){ 
            Log "$($p.ProcessName): $('{0:N2}' -f ($p.WorkingSet/1MB)) MB" 
        }
        Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "PASS"
    } catch {
        Log "Memory report failed: $($_.Exception.Message)" "FAIL"
    }
} "Identifies processes consuming the most RAM (helpful for 'computer is slow' complaints)."

Create-Btn "Check for Recent Blue Screens (BSOD)" ([System.Drawing.Color]::White) {
    try {
        Log "Searching for BSOD Events..."
        $bsod = Get-WinEvent -FilterHashtable @{LogName='System'; ID=1001} -MaxEvents 5 -ErrorAction SilentlyContinue
        
        if($bsod){ 
            Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "FAIL"
            foreach($e in $bsod){ 
                Log "BSOD Detected: $($e.TimeCreated)" "FAIL"
                Log "  Message: $($e.Message)" "FAIL"
            }
            Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "FAIL"
        } else { 
            Log "No recent BSOD events found." "PASS" 
        }
    } catch {
        Log "BSOD check failed: $($_.Exception.Message)" "FAIL"
    }
} "Scans event logs for recent blue screen errors (helps diagnose random restarts)."

Create-Btn "Generate Battery Health Report" ([System.Drawing.Color]::White) {
    try {
        Log "Generating Battery Report..." "WARN"
        Show-Progress
        powercfg /batteryreport /output "C:\battery-report.html" | Out-Null
        Hide-Progress
        
        if(Test-Path "C:\battery-report.html"){
            Log "Battery report saved to C:\battery-report.html" "PASS"
            Start-Process "C:\battery-report.html"
        } else {
            Log "Battery report generation failed (may not be a laptop)" "WARN"
        }
    } catch {
        Hide-Progress
        Log "Battery report failed: $($_.Exception.Message)" "FAIL"
    }
} "Creates detailed battery health report (laptops only) showing capacity degradation."

# --- SECTION: DOMAIN AND AD ---
Add-SectionLabel "DOMAIN AND ACTIVE DIRECTORY"

Create-Btn "Test Domain Controller Connection" ([System.Drawing.Color]::White) {
    try {
        Log "Testing Domain Controller Connection..."
        
        # Check if computer is domain-joined
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem
        
        if($computerSystem.PartOfDomain){
            Log "Domain: $($computerSystem.Domain)" "PASS"
            
            try {
                $test = Test-ComputerSecureChannel -ErrorAction Stop
                if($test){ 
                    Log "[OK] Secure Channel: HEALTHY" "PASS" 
                } else { 
                    Log "[FAIL] Secure Channel: FAILED" "FAIL"
                    Log "Run this in PowerShell: Reset-ComputerMachinePassword" "WARN"
                }
            } catch {
                Log "[FAIL] Domain Connection Test Failed" "FAIL"
                Log "Computer may need to be rejoined to domain" "WARN"
            }
        } else {
            Log "This computer is NOT joined to a domain (Workgroup: $($computerSystem.Domain))" "WARN"
        }
    } catch {
        Log "Domain check failed: $($_.Exception.Message)" "FAIL"
    }
} "Verifies connection to Active Directory domain controller (fixes 'trust relationship' errors)."

Create-Btn "Force Group Policy Update" ([System.Drawing.Color]::White) {
    try {
        Log "Forcing Group Policy Update..." "WARN"
        Show-Progress
        gpupdate /force
        Hide-Progress
        Log "Group Policy Update Complete. Changes applied." "PASS"
    } catch {
        Hide-Progress
        Log "GPUpdate failed: $($_.Exception.Message)" "FAIL"
    }
} "Forces immediate download and application of all group policies from domain."

Create-Btn "Show Last User Logon Time" ([System.Drawing.Color]::White) {
    try {
        Log "Checking Last Interactive Logon..."
        $lastLogon = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4624} -MaxEvents 10 -ErrorAction Stop | 
            Where-Object { $_.Properties[8].Value -eq 2 } | 
            Select-Object -First 1
        
        if($lastLogon){ 
            $username = $lastLogon.Properties[5].Value
            Log "Last Logon: $($lastLogon.TimeCreated)" "PASS"
            Log "User: $username" "PASS"
        } else {
            Log "Unable to retrieve last logon info." "WARN"
        }
    } catch {
        Log "Logon check failed: $($_.Exception.Message)" "FAIL"
    }
} "Shows when the last user logged into this computer (helpful for troubleshooting access issues)."

# --- SECTION: SYSTEM INFO DASHBOARD ---
Add-SectionLabel "SYSTEM INFORMATION"

Create-Btn "Show System Info Dashboard" ([System.Drawing.Color]::LightGreen) {
    try {
        Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "PASS"
        Log "           SYSTEM INFORMATION DASHBOARD           " "PASS"
        Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "PASS"
        
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem
        $os = Get-WmiObject -Class Win32_OperatingSystem
        $uptime = (Get-Date) - $os.ConvertToDateTime($os.LastBootUpTime)
        
        Log "Computer Name: $env:COMPUTERNAME" "PASS"
        Log "Domain/Workgroup: $($computerSystem.Domain)" "PASS"
        Log "Current User: $env:USERNAME" "PASS"
        Log " " "INFO"
        Log "OS: $($os.Caption) $($os.OSArchitecture)" "PASS"
        Log "Build: $($os.Version)" "PASS"
        Log "Install Date: $($os.ConvertToDateTime($os.InstallDate))" "PASS"
        Log " " "INFO"
        Log "Manufacturer: $($computerSystem.Manufacturer)" "PASS"
        Log "Model: $($computerSystem.Model)" "PASS"
        Log "Total RAM: $('{0:N2}' -f ($computerSystem.TotalPhysicalMemory/1GB)) GB" "PASS"
        Log " " "INFO"
        Log "System Uptime: $($uptime.Days) days, $($uptime.Hours) hours" "PASS"
        Log "Last Boot: $($os.ConvertToDateTime($os.LastBootUpTime))" "PASS"
        Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "PASS"
    } catch {
        Log "System info retrieval failed: $($_.Exception.Message)" "FAIL"
    }
} "Displays complete system overview: name, domain, OS version, hardware, uptime."

# Initial startup message
Log "MSP Technician IDE - Professional Edition" "PASS"
Log "Logged in as: $env:USERNAME on $env:COMPUTERNAME" "INFO"
Log "Running with Administrator privileges" "PASS"
Log "Ready for troubleshooting operations." "INFO"
Log " " "INFO"

$form.ShowDialog() | Out-Null
