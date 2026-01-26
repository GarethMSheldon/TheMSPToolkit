## Merged file: TheMSPToolKit_v2.0.0.ps1
## Sources merged: TheMSPToolKit_v2.0.0.ps1 (2025-03-15) and TheMSPToolKit_v2.ps1 (2025-10-01 snapshot)
## Created: 2026-01-26 by organizer

#Requires -RunAsAdministrator
<#
.SYNOPSIS
MSP Technician Toolkit — v2.0.0 (consolidated)

.DESCRIPTION
Consolidated v2.0.0 script combining the official v2.0.0 release snapshot and
an intermediate v2 snapshot. This canonical file preserves release metadata,
includes change notes when available, enforces Administrator elevation, and
provides the GUI and helper functions.

.NOTES
Merged sources: TheMSPToolKit_v2.0.0.ps1, TheMSPToolKit_v2.ps1
Merged on: 2026-01-26
#>

# ============================================================================
# Release Information
# ============================================================================
$Version = '2.0.0'
$ReleaseDate = '2025-03-15'
$Summary = 'Initial Public Release with core automation and GUI migration.'
$Changes = @(
    'System File Checker automation',
    'CHKDSK scanning capability',
    'Teams cache cleaner',
    'Network stack reset',
    'Print spooler clearing',
    'Memory-hogging process identification',
    'SMART status checking',
    'Domain connection testing',
    'Group policy updater',
    'Basic logging functionality'
)

function Show-Changes {
    Write-Host "MSP Toolkit - $Version ($ReleaseDate)" -ForegroundColor Cyan
    Write-Host $Summary
    Write-Host "`nChanges:" -ForegroundColor Yellow
    $Changes | ForEach-Object { Write-Host "- $_" }
}

# Include release notes file (if present)
if (Test-Path (Join-Path $PSScriptRoot 'change_2.0.0.ps1')) { . (Join-Path $PSScriptRoot 'change_2.0.0.ps1') }

# Display changes if dot-sourced
if ($MyInvocation.InvocationName -eq '.') { Show-Changes }

# ============================================================================
# Administrator Privilege Check
# ============================================================================
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# ============================================================================
# UI Assembly Loading
# ============================================================================
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# ============================================================================
# Main Form Setup
# ============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "MSP Technician IDE - v$Version (Running as Admin)"
$form.Size = New-Object System.Drawing.Size(1150, 850)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

# ============================================================================
# Global Output Box
# ============================================================================
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Dock = "Right"
$outputBox.Width = 600
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$outputBox.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$outputBox.ForeColor = [System.Drawing.Color]::Gainsboro
$form.Controls.Add($outputBox)

# ============================================================================
# Logging Function
# ============================================================================
function Log([string]$msg, [string]$status="INFO") {
    $color = "White"
    if($status -eq "PASS"){$color="Lime"}
    if($status -eq "FAIL"){$color="OrangeRed"}
    if($status -eq "WARN"){$color="Yellow"}
    $outputBox.SelectionColor = [System.Drawing.Color]::$color
    $outputBox.AppendText("[$status] $msg`r`n")
    $outputBox.ScrollToCaret()
}

# ============================================================================
# Sidebar Panel
# ============================================================================
$sidebar = New-Object System.Windows.Forms.FlowLayoutPanel
$sidebar.Dock = "Left"
$sidebar.Width = 520
$sidebar.FlowDirection = "TopDown"
$sidebar.AutoScroll = $true
$sidebar.Padding = New-Object System.Windows.Forms.Padding(10)
$sidebar.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 235)
$form.Controls.Add($sidebar)

# ============================================================================
# UI Helper Functions
# ============================================================================
function Add-SectionLabel($text) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "`n$text"
    $label.Width = 480
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $sidebar.Controls.Add($label)
}

function Create-Btn($text, $color, $script) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Width = 470
    $btn.Height = 40
    $btn.Margin = New-Object System.Windows.Forms.Padding(0, 5, 0, 5)
    $btn.FlatStyle = "Flat"
    $btn.BackColor = $color
    $btn.Add_Click($script)
    $sidebar.Controls.Add($btn)
}

# ============================================================================
# SYSTEM REPAIR & HEALTH
# ============================================================================
Add-SectionLabel "SYSTEM REPAIR & HEALTH"

Create-Btn "Run SFC Scan (System File Checker)" "LightCoral" {
    Log "Starting SFC Scan... this may take several minutes." "WARN"
    Start-Process "sfc" -ArgumentList "/scannow" -Wait -NoNewWindow
    Log "SFC Scan Complete. Check console for results." "PASS"
}

Create-Btn "Check for Pending Reboot" "White" {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
    if (Test-Path $regPath) { Log "Reboot is PENDING (Registry flag found)." "FAIL" }
    else { Log "No pending reboot flags detected." "PASS" }
}

# ============================================================================
# USER PROFILE & M365
# ============================================================================
Add-SectionLabel "USER PROFILE & M365"

Create-Btn "Clear Teams Cache (Classic & New)" "White" {
    Log "Closing Teams..."
    Get-Process "Teams" -ErrorAction SilentlyContinue | Stop-Process -Force
    Log "Clearing Cache Folders..."
    $paths = @("$env:AppData\Microsoft\Teams", "$env:LocalAppData\Packages\MSTeams_8wekyb3d8bbwe\LocalCache")
    foreach($p in $paths){ if(Test-Path $p){ Remove-Item "$p\*" -Recurse -Force -ErrorAction SilentlyContinue } }
    Log "Teams cache cleared." "PASS"
}

Create-Btn "Fix Outlook Profile (Delete & Recreate)" "LightCoral" {
    Log "Closing Outlook..." "WARN"
    Get-Process "OUTLOOK" -ErrorAction SilentlyContinue | Stop-Process -Force
    Log "Deleting Outlook Profile Registry Keys..."
    Remove-Item "HKCU:\Software\Microsoft\Office\*\Outlook\Profiles\Outlook" -Recurse -Force -ErrorAction SilentlyContinue
    Log "Outlook profile removed. Reopen Outlook to recreate." "PASS"
}

Create-Btn "Fix OneDrive Sync Issues" "White" {
    Log "Resetting OneDrive..." "WARN"
    Start-Process "$env:LocalAppData\Microsoft\OneDrive\onedrive.exe" -ArgumentList "/reset" -NoNewWindow
    Start-Sleep -Seconds 5
    Start-Process "$env:LocalAppData\Microsoft\OneDrive\onedrive.exe" -NoNewWindow
    Log "OneDrive reset complete." "PASS"
}

Create-Btn "Fix Disconnected Mapped Drives" "White" {
    Log "Checking Network Drives..."
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot -ne $null }
    foreach($d in $drives){
        Log "Testing connection to $($d.DisplayRoot)..."
        if(Test-Path $d.Root){ Log "Drive $($d.Name) is Connected." "PASS" }
        else { Log "Drive $($d.Name) is DISCONNECTED." "FAIL" }
    }
}

# ============================================================================
# ADVANCED NETWORKING
# ============================================================================
Add-SectionLabel "ADVANCED NETWORKING"

Create-Btn "Find Top 5 Network Consuming Processes" "LightBlue" {
    Log "Monitoring Network Traffic..."
    $stats = Get-NetTCPConnection | Group-Object -Property OwningProcess | Select-Object Count, Name, @{Name="Process";Expression={(Get-Process -Id $_.Name -ErrorAction SilentlyContinue).ProcessName}} | Sort-Object Count -Descending | Select-Object -First 5
    foreach($s in $stats){ Log "PID $($s.Name) ($($s.Process)): $($s.Count) Connections" }
}

Create-Btn "List WiFi Passwords (Saved Profiles)" "White" {
    Log "Retrieving Saved WiFi Profiles..."
    $profiles = netsh wlan show profiles | Select-String "\:(.+)$" | ForEach-Object { $_.Matches.Groups[1].Value.Trim() }
    foreach($p in $profiles){
        $pass = netsh wlan show profile name="$p" key=clear | Select-String "Key Content\W+\:(.+)$" | ForEach-Object { $_.Matches.Groups[1].Value.Trim() }
        Log "SSID: $p | Key: $pass"
    }
}

Create-Btn "Reset Network Stack (Full Nuclear Option)" "LightCoral" {
    Log "Resetting Network Stack..." "WARN"
    netsh winsock reset | Out-Null
    netsh int ip reset | Out-Null
    ipconfig /flushdns | Out-Null
    Log "Network Stack Reset. Reboot required!" "PASS"
}

Create-Btn "Show Public & Local IP Addresses" "White" {
    Log "Fetching IP Information..."
    $localIP = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.InterfaceAlias -notlike "*Loopback*"} | Select-Object -First 1).IPAddress
    try { $publicIP = (Invoke-WebRequest -Uri "https://api.ipify.org" -UseBasicParsing).Content }
    catch { $publicIP = "Unable to fetch" }
    Log "Local IP: $localIP" "PASS"
    Log "Public IP: $publicIP" "PASS"
}

# ============================================================================
# PRINTER TROUBLESHOOTING
# ============================================================================
Add-SectionLabel "PRINTER TROUBLESHOOTING"

Create-Btn "Clear Print Spooler & Restart Service" "White" {
    Log "Stopping Print Spooler..."
    Stop-Service -Name Spooler -Force
    Log "Clearing Spool Folder..."
    Remove-Item "C:\Windows\System32\spool\PRINTERS\*" -Force -ErrorAction SilentlyContinue
    Start-Service -Name Spooler
    Log "Print Spooler Cleared & Restarted." "PASS"
}

Create-Btn "List All Installed Printers" "White" {
    Log "Listing Printers..."
    Get-Printer | ForEach-Object { Log "$($_.Name) - Status: $($_.PrinterStatus)" }
}

Create-Btn "Remove Offline/Ghost Printers" "LightCoral" {
    Log "Removing Offline Printers..."
    $offline = Get-Printer | Where-Object { $_.PrinterStatus -eq "Offline" }
    foreach($p in $offline){ Remove-Printer -Name $p.Name -ErrorAction SilentlyContinue; Log "Removed: $($p.Name)" "PASS" }
    if($offline.Count -eq 0){ Log "No offline printers found." "PASS" }
}

# ============================================================================
# DISK & STORAGE HEALTH
# ============================================================================
Add-SectionLabel "DISK & STORAGE HEALTH"

Create-Btn "Run CHKDSK (Read-Only Scan)" "White" {
    Log "Running CHKDSK on C:..." "WARN"
    chkdsk C: /scan
    Log "CHKDSK Complete." "PASS"
}

Create-Btn "Find Large Files (>500MB)" "White" {
    Log "Scanning for large files (this may take time)..." "WARN"
    Get-ChildItem "C:\Users" -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Length -gt 500MB } | Sort-Object Length -Descending | Select-Object -First 10 | ForEach-Object { Log "$($_.FullName) - $('{0:N2}' -f ($_.Length/1GB)) GB" }
}

Create-Btn "Check Disk Health (SMART Status)" "White" {
    Log "Checking Disk Health..."
    Get-PhysicalDisk | ForEach-Object { Log "Disk $($_.FriendlyName): Health=$($_.HealthStatus) OpStatus=$($_.OperationalStatus)" }
}

Create-Btn "Clean Temp Files (System & User)" "LightCoral" {
    Log "Cleaning Temporary Files..." "WARN"
    Remove-Item "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
    Log "Temp files cleaned." "PASS"
}

# ============================================================================
# WINDOWS UPDATE FIXES
# ============================================================================
Add-SectionLabel "WINDOWS UPDATE FIXES"

Create-Btn "Reset Windows Update Components" "LightCoral" {
    Log "Stopping Update Services..." "WARN"
    Stop-Service -Name wuauserv, cryptSvc, bits, msiserver -Force
    Log "Renaming SoftwareDistribution folder..."
    Rename-Item "C:\Windows\SoftwareDistribution" "SoftwareDistribution.old" -Force -ErrorAction SilentlyContinue
    Rename-Item "C:\Windows\System32\catroot2" "catroot2.old" -Force -ErrorAction SilentlyContinue
    Start-Service -Name wuauserv, cryptSvc, bits, msiserver
    Log "Windows Update Components Reset." "PASS"
}

Create-Btn "Check Last Update Install Date" "White" {
    Log "Checking Windows Update History..."
    $lastUpdate = Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1
    Log "Last Update: $($lastUpdate.HotFixID) installed on $($lastUpdate.InstalledOn)" "PASS"
}

# ============================================================================
# PERFORMANCE & CRASHES
# ============================================================================
Add-SectionLabel "PERFORMANCE & CRASHES"

Create-Btn "Find Memory-Hogging Processes (Top 5)" "White" {
    Log "Memory Usage Report..."
    Get-Process | Sort-Object WorkingSet -Descending | Select-Object -First 5 | ForEach-Object { Log "$($_.ProcessName): $('{0:N2}' -f ($_.WorkingSet/1MB)) MB" }
}

Create-Btn "Check for Recent Blue Screen Logs" "White" {
    Log "Searching for BSOD Events..."
    $bsod = Get-WinEvent -FilterHashtable @{LogName='System'; ID=1001} -MaxEvents 5 -ErrorAction SilentlyContinue
    if($bsod){ foreach($e in $bsod){ Log "BSOD on $($e.TimeCreated): $($e.Message)" "FAIL" } }
    else { Log "No recent BSOD events found." "PASS" }
}

Create-Btn "Generate Battery Health Report" "White" {
    Log "Generating Battery Report..." "WARN"
    powercfg /batteryreport /output "C:\battery-report.html"
    Log "Battery report saved to C:\battery-report.html" "PASS"
    Start-Process "C:\battery-report.html"
}

# ============================================================================
# DOMAIN & AD ISSUES
# ============================================================================
Add-SectionLabel "DOMAIN & AD ISSUES"

Create-Btn "Test Domain Controller Connection" "White" {
    Log "Testing Domain Controller..."
    $test = Test-ComputerSecureChannel -Verbose
    if($test){ Log "Domain Connection: HEALTHY" "PASS" }
    else { Log "Domain Connection: FAILED - Run 'Reset-ComputerMachinePassword'" "FAIL" }
}

Create-Btn "Force Group Policy Update" "White" {
    Log "Forcing Group Policy Update..." "WARN"
    gpupdate /force
    Log "Group Policy Update Complete." "PASS"
}

Create-Btn "Show Last User Logon Time" "White" {
    Log "Checking Last Interactive Logon..."
    $lastLogon = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4624} -MaxEvents 1 -ErrorAction SilentlyContinue | Where-Object { $_.Properties[8].Value -eq 2 }
    if($lastLogon){ Log "Last Logon: $($lastLogon.TimeCreated) - User: $($lastLogon.Properties[5].Value)" }
    else { Log "Unable to retrieve last logon info." "WARN" }
}

# ============================================================================
# Application Entry Point
# ============================================================================
Log "MSP Technician Toolkit v$Version loaded" "PASS"
Log "Release Date: $ReleaseDate" "INFO"
Log "Ready for use." "PASS"

# Show the form
$form.ShowDialog()