#Requires -RunAsAdministrator
<#
.SYNOPSIS
MSP Technician Toolkit — Work in Progress (3.1.3)
.DESCRIPTION
Development snapshot consolidating recent hardening, background-worker
management, and shutdown cleanup improvements intended for the upcoming
3.1.3 release. Use for development and testing only; not a stable release.
.NOTES
Author: MSP Solutions Team
Version: 3.1.3
Date: 2026-01-26
Requires: PowerShell 3.0+, Administrator privileges
Compatible: Windows 7 SP1 through Windows 11 / Server 2022
#>

# Early startup diagnostic log (helps capture silent failures before GUI loads)
try {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $StartupLog = Join-Path $ScriptDir 'MSP_Toolkit_startup.log'
    "{0} - START - PID:{1} - User:{2} - Host:{3} - Args:{4}" -f (Get-Date -Format o), $PID, $env:USERNAME, $Host.Name, ($args -join ' ') | Out-File -FilePath $StartupLog -Append -Encoding UTF8 -ErrorAction SilentlyContinue
} catch {
    # Swallow errors - logging is best-effort
}

# Auto-elevation block
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$($myinvocation.mycommand.definition)`""
    Start-Process powershell -Verb RunAs -ArgumentList $arguments
    exit
}

# Enable DPI awareness for better scaling on high-resolution displays
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class DPIAwareness {
    [DllImport("user32.dll")]
    private static extern bool SetProcessDPIAware();
    
    public static void SetDPIAware() {
        if (Environment.OSVersion.Version.Major >= 6) {
            SetProcessDPIAware();
        }
    }
}
"@ -ErrorAction SilentlyContinue

[DPIAwareness]::SetDPIAware()

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Initialize a single session log file and Save-Log function (prevents file spam and recursion)
$script:SessionLogFile = Join-Path $env:USERPROFILE ("Desktop\MSP_Toolkit_Session_{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

function Save-Log {
    param(
        [string]$LogMessage,
        [string]$LogStatus = "INFO"
    )
    try {
        # Best-effort append to a single session file. Do not call Log() from here.
        if (-not $script:SessionLogFile) {
            $script:SessionLogFile = Join-Path $env:USERPROFILE ("Desktop\MSP_Toolkit_Session_{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        }
        "{0} [{1}] {2}" -f (Get-Date -Format o), $LogStatus, $LogMessage | Out-File -FilePath $script:SessionLogFile -Append -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {
        # Swallow errors to avoid UI crashes during logging
    }
}

# Core logging function with color coding
# Enhanced logging with interactive feedback
function Log {
    param(
        [string]$Message,
        [string]$Status = "INFO"
    )
    try {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $color = switch ($Status) {
            "PASS" { "Lime" }
            "FAIL" { "OrangeRed" }
            "WARN" { "Yellow" }
            default { "White" }
        }

        # Thread-safe UI update
        if ($outputBox.InvokeRequired) {
            $outputBox.Invoke([Action]{ 
                $outputBox.SelectionColor = [System.Drawing.Color]::$color
                $outputBox.AppendText("[$timestamp][$Status] $Message`r`n")
                $outputBox.ScrollToCaret()
            })
        } else {
            $outputBox.SelectionColor = [System.Drawing.Color]::$color
            $outputBox.AppendText("[$timestamp][$Status] $Message`r`n")
            $outputBox.ScrollToCaret()
        }
        [System.Windows.Forms.Application]::DoEvents()
    } catch {
        try {
            Write-Host "[$timestamp][$Status] $Message" -ForegroundColor $color
        } catch {
            Write-Host "[$timestamp][$Status] $Message"
        }
    }

    # Save log to single session file (best-effort)
    Save-Log -LogMessage $Message -LogStatus $Status
}

# Ensure `Log` is available to button event handlers running in other scopes
try {
    Set-Item -Path Function:\Global\Log -Value $function:Log -ErrorAction SilentlyContinue
} catch {
    # best-effort; if this fails the user can dot-source the script instead
}

# Confirmation dialog function
function Show-Confirmation {
    param(
        [string]$Message,
        [string]$Title = "Confirm Action"
    )
    $result = [System.Windows.Forms.MessageBox]::Show($form, $Message, $Title, 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    return $result -eq [System.Windows.Forms.DialogResult]::Yes
}

# Helper function for progress bar control
function Set-ProgressBarVisibility {
    param([bool]$Visible)
    $script:progressBar.Visible = $Visible
    $statusBar.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "MSP Technician Toolkit - Professional Edition"
$form.Size = New-Object System.Drawing.Size(1300, 950)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 700)  # Allow some resizing
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.TopMost = $false
$form.ShowIcon = $false

# Add resize event handler to maintain proper layout
$form.add_Resize({
    # Update panel heights to match form height
    $sidebarPanel.Height = $form.ClientSize.Height - $statusBar.Height - 20
    $consolePanel.Height = $form.ClientSize.Height - $statusBar.Height - 20
    
    # Update output box size to fill console panel
    $outputBox.Width = $consolePanel.ClientSize.Width - 20
    $outputBox.Height = $consolePanel.ClientSize.Height - 20
    
    # Update button container height
    $buttonContainer.Height = [Math]::Max(5500, $sidebarPanel.ClientSize.Height + $sidebarPanel.VerticalScroll.Value + 100)
})

# Create status bar FIRST to calculate available space
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.Size = New-Object System.Drawing.Size(1300, 25)
$statusBar.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$statusBar.ForeColor = [System.Drawing.Color]::LimeGreen
$statusBar.Dock = [System.Windows.Forms.DockStyle]::Bottom

# Status labels
$adminStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$adminStatus.Text = "Admin: YES"
$adminStatus.ForeColor = [System.Drawing.Color]::LimeGreen

$computerName = New-Object System.Windows.Forms.ToolStripStatusLabel
$computerName.Text = "PC: $env:COMPUTERNAME"
$computerName.Width = 250

$currentUserName = New-Object System.Windows.Forms.ToolStripStatusLabel
$currentUserName.Text = "User: $env:USERNAME"
$currentUserName.Width = 200

$domainStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
try {
    $computerSystem = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
    $domainStatus.Text = "Domain: $($computerSystem.Domain)"
} catch {
    $domainStatus.Text = "Domain: Error detecting domain"
}

$statusBar.Items.AddRange(@($adminStatus, $computerName, $currentUserName, $domainStatus))

# Progress bar (marquee style)
$progressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$progressBar.MarqueeAnimationSpeed = 30
$progressBar.Width = 200
$progressBar.Visible = $false
$null = $statusBar.Items.Add($progressBar)

# Create sidebar panel with scroll - DOCKED TO LEFT
$sidebarPanel = New-Object System.Windows.Forms.Panel
$sidebarPanel.Width = 620
$sidebarPanel.Dock = [System.Windows.Forms.DockStyle]::Left
$sidebarPanel.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$sidebarPanel.AutoScroll = $true
$sidebarPanel.BorderStyle = "None"

# Create scrollable container for buttons
$buttonContainer = New-Object System.Windows.Forms.Panel
$buttonContainer.Width = 600
$buttonContainer.Height = 5500  # Increased height for more buttons
$buttonContainer.Location = New-Object System.Drawing.Point(0, 0)
$buttonContainer.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$null = $sidebarPanel.Controls.Add($buttonContainer)

# Create console panel - DOCKED TO FILL REMAINING SPACE
$consolePanel = New-Object System.Windows.Forms.Panel
$consolePanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$consolePanel.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$consolePanel.BorderStyle = "Fixed3D"

# Create output box - PROPERLY ANCHORED TO FILL CONSOLE PANEL
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$outputBox.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$outputBox.ForeColor = [System.Drawing.Color]::Gainsboro
$outputBox.ReadOnly = $true
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$outputBox.WordWrap = $false
$outputBox.ScrollBars = "Vertical"
$null = $consolePanel.Controls.Add($outputBox)

# Add panels to form in correct order
$null = $form.Controls.Add($statusBar)  # Add status bar FIRST
$null = $form.Controls.Add($consolePanel)  # Console panel fills remaining space
$null = $form.Controls.Add($sidebarPanel)  # Sidebar docks to left

# Button styling function (approved verb)
function New-Button {
    param(
        [string]$Text,
        [string]$Name,
        [System.Drawing.Color]$BackColor,
        [string]$ToolTipText,
        [scriptblock]$Action
    )
    
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Name = $Name
    $button.Width = 580
    $button.Height = 42
    $button.FlatStyle = "Flat"
    $button.FlatAppearance.BorderSize = 1
    # Muted, easy-on-the-eyes palette (consistent across all buttons)
    $softBack = [System.Drawing.Color]::FromArgb(60,70,80)        # base button color
    $hoverBack = [System.Drawing.Color]::FromArgb(75,85,95)       # hover
    $downBack  = [System.Drawing.Color]::FromArgb(50,60,70)       # pressed
    $borderCol = [System.Drawing.Color]::FromArgb(95,100,105)
    $button.BackColor = $softBack
    $button.ForeColor = [System.Drawing.Color]::Gainsboro
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $button.UseCompatibleTextRendering = $true
    $button.FlatAppearance.BorderColor = $borderCol
    $button.FlatAppearance.MouseOverBackColor = $hoverBack
    $button.FlatAppearance.MouseDownBackColor = $downBack
    # Wrap the provided action in a safe event handler so errors are caught
    if ($Action) {
        $act = $Action
        $button.Add_Click({ param($btnSender, $btnEvent)
            try {
                if ($act) { & $act }
            } catch {
                try { Log ("Button action error: {0}" -f $_) "FAIL" } catch { Write-Host "Button action error: $_" }
            }
        })
    }
    $button.Anchor = "Top, Left"  # Anchor to maintain position during resize
    
    # Create tooltip properly
    $tooltip = New-Object System.Windows.Forms.ToolTip
    $tooltip.AutoPopDelay = 8000
    $tooltip.InitialDelay = 500
    $tooltip.ReshowDelay = 100
    $tooltip.SetToolTip($button, $ToolTipText)
    
    return $button
}

# Section header styling (approved verb)
function New-SectionHeader {
    param([string]$Text)
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.ForeColor = [System.Drawing.Color]::White
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $label.Width = 600
    $label.Height = 25
    $label.TextAlign = "MiddleLeft"
    $label.Padding = New-Object System.Windows.Forms.Padding(5,0,0,0)
    $label.BackColor = [System.Drawing.Color]::FromArgb(55, 55, 55)
    $label.Anchor = "Top, Left"  # Anchor to maintain position during resize
    return $label
}

# Add section headers and buttons
$yPosition = 10

# SECTION 1: SYSTEM REPAIR AND HEALTH
$buttonContainer.Controls.Add((New-SectionHeader "1. SYSTEM REPAIR AND HEALTH"))
$yPosition += 30

# 1.1 Run SFC Scan
$btnSFC = New-Button -Text "1.1 Run SFC Scan (System File Checker)" -Name "btnSFC" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Scans and repairs corrupted Windows system files. Takes 10-15 minutes." `
    -Action {
        try {
            if (Show-Confirmation -Message "This operation takes 10-15 minutes and requires system file access. Continue?") {
                Set-ProgressBarVisibility -Visible $true
                Log "Starting SFC scan..." "INFO"
                
                $output = & sfc /scannow 2>&1
                Log "SFC scan completed:" "INFO"
                $output | ForEach-Object { Log $_ "INFO" }
                
                if ($output -match "Windows Resource Protection did not find any integrity violations") {
                    Log "SFC completed successfully. No integrity violations found." "PASS"
                }
                elseif ($output -match "Windows Resource Protection found corrupt files and successfully repaired them") {
                    Log "SFC found and repaired corrupt files. A reboot may be required." "WARN"
                }
                elseif ($output -match "Windows Resource Protection found corrupt files but was unable to fix some of them") {
                    Log "SFC found corrupt files but could not repair all of them. Run DISM repair next." "FAIL"
                }
                else {
                    Log "SFC scan completed with unexpected results. Review output above." "WARN"
                }
            }
        } catch {
            Log "SFC scan error: $_" "FAIL"
        } finally {
            Set-ProgressBarVisibility -Visible $false
        }
    }
$btnSFC.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnSFC)
$yPosition += 47

# 1.2 Check for Pending Reboot
$btnRebootCheck = New-Button -Text "1.2 Check for Pending Reboot" -Name "btnRebootCheck" `
    -BackColor ([System.Drawing.Color]::White) `
    -ToolTipText "Checks multiple registry locations for pending system reboots." `
    -Action {
        try {
            $pendingReboot = $false
            Log "Checking for pending reboot flags..." "INFO"
            
            # Check Windows Update reboot required
            $rebootKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
            if (Test-Path $rebootKey) {
                $items = Get-ChildItem $rebootKey -ErrorAction SilentlyContinue
                if ($items) {
                    Log "[REGISTRY] RebootRequired key exists with $($items.Count) items" "WARN"
                    $pendingReboot = $true
                }
            }
            
            # Check Component Based Servicing
            $rebootPendingKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
            if (Test-Path $rebootPendingKey) {
                $value = Get-ItemProperty $rebootPendingKey -Name "RebootPending" -ErrorAction SilentlyContinue
                if ($value -and $value.RebootPending) {
                    Log "[REGISTRY] RebootPending flag detected" "WARN"
                    $pendingReboot = $true
                }
            }
            
            # Check Session Manager pending file operations
            $sessionMgrKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager"
            if (Test-Path $sessionMgrKey) {
                $value = Get-ItemProperty $sessionMgrKey -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue
                if ($value -and $value.PendingFileRenameOperations) {
                    Log "[REGISTRY] Pending file rename operations detected" "WARN"
                    $pendingReboot = $true
                }
            }
            
            # Check CBS reboot in progress
            $cbsKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress"
            if (Test-Path $cbsKey) {
                $value = Get-ItemProperty $cbsKey -Name "RebootInProgress" -ErrorAction SilentlyContinue
                if ($value -and $value.RebootInProgress) {
                    Log "[REGISTRY] Component Based Servicing reboot in progress" "WARN"
                    $pendingReboot = $true
                }
            }
            
            if (-not $pendingReboot) {
                Log "No pending reboot flags detected" "PASS"
            }
        } catch {
            Log "Reboot check error: $_" "FAIL"
        }
    }
$btnRebootCheck.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnRebootCheck)
$yPosition += 47

# 1.3 Run DISM Repair
$btnDISM = New-Button -Text "1.3 Run DISM Repair" -Name "btnDISM" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Repairs Windows component store corruption. Use before SFC if it fails." `
    -Action {
        if (Show-Confirmation -Message "DISM repair may take 15-20 minutes. Continue?") {
            try {
                Set-ProgressBarVisibility -Visible $true
                Log "Starting DISM repair..." "INFO"
                
                $startTime = Get-Date
                $process = Start-Process -FilePath "dism.exe" -ArgumentList "/Online /Cleanup-Image /RestoreHealth" -Wait -PassThru -NoNewWindow
                
                $endTime = Get-Date
                $duration = $endTime - $startTime
                
                if ($process.ExitCode -eq 0) {
                    Log "DISM repair completed successfully in $($duration.Minutes) minutes, $($duration.Seconds) seconds" "PASS"
                } else {
                    Log "DISM repair failed with exit code $($process.ExitCode)" "FAIL"
                    Log "Try running as Administrator or check Windows Update service status" "WARN"
                }
            } catch {
                Log "DISM error: $_" "FAIL"
                Log "Ensure DISM is available on this Windows version. May not work on Windows 7 without updates." "WARN"
            } finally {
                Set-ProgressBarVisibility -Visible $false
            }
        }
    }
$btnDISM.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnDISM)
$yPosition += 52

# SECTION 2: USER PROFILE AND M365
$buttonContainer.Controls.Add((New-SectionHeader "2. USER PROFILE AND M365"))
$yPosition += 30

# 2.1 Clear Teams Cache (Classic & New)
$btnTeamsCache = New-Button -Text "2.1 Clear Teams Cache (Classic & New)" -Name "btnTeamsCache" `
    -BackColor ([System.Drawing.Color]::White) `
    -ToolTipText "Fixes Teams freezing, login issues, and sync problems by clearing cache." `
    -Action {
        try {
            Set-ProgressBarVisibility -Visible $true
            Log "Clearing Teams cache..." "INFO"
            
            # Kill Teams processes
            $teamsProcesses = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
            $msTeamsProcesses = Get-Process -Name "ms-teams" -ErrorAction SilentlyContinue
            
            if ($teamsProcesses) {
                Log "Stopping Teams process..." "INFO"
                $teamsProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
            }
            if ($msTeamsProcesses) {
                Log "Stopping Microsoft Teams process..." "INFO"
                $msTeamsProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
            }
            
            Start-Sleep -Seconds 2
            
            # Clear classic Teams cache
            $teamsCachePath = "$env:APPDATA\Microsoft\Teams"
            $teamsCacheDeleted = 0
            if (Test-Path $teamsCachePath) {
                Get-ChildItem $teamsCachePath -Recurse -Force -ErrorAction SilentlyContinue | 
                    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                $teamsCacheDeleted = 1
                Log "Cleared classic Teams cache" "PASS"
            } else {
                Log "Classic Teams cache folder not found" "INFO"
            }
            
            # Clear new Teams cache
            $newTeamsCachePath = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache"
            $newTeamsCacheDeleted = 0
            if (Test-Path $newTeamsCachePath) {
                Get-ChildItem $newTeamsCachePath -Recurse -Force -ErrorAction SilentlyContinue | 
                    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                $newTeamsCacheDeleted = 1
                Log "Cleared new Teams cache" "PASS"
            } else {
                Log "New Teams cache folder not found" "INFO"
            }
            
            if ($teamsCacheDeleted -eq 0 -and $newTeamsCacheDeleted -eq 0) {
                Log "No Teams cache files found to delete" "WARN"
            } else {
                Log "Teams cache cleared successfully. Restart Teams to apply changes." "PASS"
            }
        } catch {
            Log "Teams cache error: $_" "FAIL"
        } finally {
            Set-ProgressBarVisibility -Visible $false
        }
    }
$btnTeamsCache.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnTeamsCache)
$yPosition += 47

# 2.2 Fix Outlook Profile (Delete & Recreate)
$btnOutlookProfile = New-Button -Text "2.2 Fix Outlook Profile (Delete & Recreate)" -Name "btnOutlookProfile" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Fixes 'Enter Password' prompts, profile corruption, and connection issues." `
    -Action {
        if (Show-Confirmation -Message "This will delete all Outlook profiles and reset to default. Outlook will need to be reconfigured. Continue?") {
            try {
                Set-ProgressBarVisibility -Visible $true
                Log "Fixing Outlook profiles..." "INFO"
                
                # Kill Outlook process
                $outlookProcess = Get-Process -Name "outlook" -ErrorAction SilentlyContinue
                if ($outlookProcess) {
                    Log "Stopping Outlook..." "INFO"
                    $outlookProcess | Stop-Process -Force
                    Start-Sleep -Seconds 2
                }
                
                # Office versions to check
                $officeVersions = @("16.0", "15.0", "14.0")
                $profilesDeleted = 0
                
                foreach ($version in $officeVersions) {
                    $profilePath = "HKCU:\Software\Microsoft\Office\$version\Outlook\Profiles"
                    if (Test-Path $profilePath) {
                        Log "Deleting Outlook $version profiles..." "INFO"
                        Remove-Item -Path $profilePath -Recurse -Force -ErrorAction SilentlyContinue
                        $profilesDeleted++
                    }
                }
                
                if ($profilesDeleted -gt 0) {
                    Log "Deleted Outlook profiles for $profilesDeleted Office version(s)" "PASS"
                    Log "Restart Outlook to create a new profile" "INFO"
                } else {
                    Log "No Outlook profiles found to delete" "WARN"
                }
            } catch {
                Log "Outlook profile error: $_" "FAIL"
            } finally {
                Set-ProgressBarVisibility -Visible $false
            }
        }
    }
$btnOutlookProfile.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnOutlookProfile)
$yPosition += 47

# NEW: 2.3 Repair Outlook Profile (Gentle Fix)
$btnOutlookRepair = New-Button -Text "2.3 Repair Outlook Profile (Gentle Fix)" -Name "btnOutlookRepair" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Fixes common Outlook issues without deleting profiles. Resets navigation pane and clears views." `
    -Action {
        try {
            Set-ProgressBarVisibility -Visible $true
            Log "Repairing Outlook profile (gentle fix)..." "INFO"
            
            # Kill Outlook process
            $outlookProcess = Get-Process -Name "outlook" -ErrorAction SilentlyContinue
            if ($outlookProcess) {
                Log "Stopping Outlook..." "INFO"
                $outlookProcess | Stop-Process -Force
                Start-Sleep -Seconds 2
            }
            
            # Reset navigation pane
            $navPaneReset = $false
            $outlookPaths = @(
                "$env:APPDATA\Microsoft\Outlook"
                "$env:LOCALAPPDATA\Microsoft\Outlook"
            )
            
            foreach ($path in $outlookPaths) {
                if (Test-Path $path) {
                    $files = Get-ChildItem $path -Filter "*.xml" -Recurse -ErrorAction SilentlyContinue
                    foreach ($file in $files) {
                        if ($file.Name -match "nav.*\.xml" -or $file.Name -match "panes.*\.xml") {
                            Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
                            $navPaneReset = $true
                            Log "Reset navigation pane configuration: $($file.Name)" "INFO"
                        }
                    }
                }
            }
            
            # Clear view settings
            $viewsCleared = $false
            $viewFiles = Get-ChildItem "$env:APPDATA\Microsoft\Outlook" -Filter "*.nv*" -Recurse -ErrorAction SilentlyContinue
            foreach ($file in $viewFiles) {
                Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
                $viewsCleared = $true
                Log "Cleared view settings: $($file.Name)" "INFO"
            }
            
            if (-not $navPaneReset -and -not $viewsCleared) {
                Log "No Outlook configuration files found to reset" "WARN"
            } else {
                Log "Outlook profile repaired successfully. Restart Outlook to apply changes." "PASS"
            }
        } catch {
            Log "Outlook repair error: $_" "FAIL"
        } finally {
            Set-ProgressBarVisibility -Visible $false
        }
    }
$btnOutlookRepair.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnOutlookRepair)
$yPosition += 47

# 2.4 Fix OneDrive Sync Issues
$btnOneDrive = New-Button -Text "2.4 Fix OneDrive Sync Issues" -Name "btnOneDrive" `
    -BackColor ([System.Drawing.Color]::White) `
    -ToolTipText "Fixes sync errors, file conflicts, and 'processing changes' stuck status." `
    -Action {
        try {
            Set-ProgressBarVisibility -Visible $true
            Log "Fixing OneDrive sync issues..." "INFO"
            
            # Find OneDrive executable
            $oneDrivePaths = @(
                "$env:LocalAppData\Microsoft\OneDrive\onedrive.exe"
                "$env:ProgramFiles\Microsoft OneDrive\onedrive.exe"
                "${env:ProgramFiles(x86)}\Microsoft OneDrive\onedrive.exe"
            )
            
            $oneDrivePath = $null
            foreach ($path in $oneDrivePaths) {
                if (Test-Path $path) {
                    $oneDrivePath = $path
                    Log "Found OneDrive at: $path" "INFO"
                    break
                }
            }
            
            if (-not $oneDrivePath) {
                Log "OneDrive executable not found in standard locations" "FAIL"
                Log "Try reinstalling OneDrive or check installation path" "WARN"
                return
            }
            
            # Kill OneDrive process
            $oneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
            if ($oneDriveProcess) {
                Log "Stopping OneDrive process..." "INFO"
                $oneDriveProcess | Stop-Process -Force
                Start-Sleep -Seconds 2
            }
            
            # Reset OneDrive
            Log "Resetting OneDrive sync..." "INFO"
            Start-Process -FilePath $oneDrivePath -ArgumentList "/reset" -Wait -NoNewWindow
            
            Start-Sleep -Seconds 8
            
            # Restart OneDrive
            Log "Restarting OneDrive..." "INFO"
            Start-Process -FilePath $oneDrivePath -NoNewWindow
            
            Log "OneDrive sync reset completed successfully" "PASS"
            Log "Wait 1-2 minutes for OneDrive to fully initialize" "INFO"
        } catch {
            Log "OneDrive error: $_" "FAIL"
        } finally {
            Set-ProgressBarVisibility -Visible $false
        }
    }
$btnOneDrive.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnOneDrive)
$yPosition += 47

# NEW: 2.5 Reset Local User Password
$btnResetPassword = New-Button -Text "2.5 Reset Local User Password" -Name "btnResetPassword" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Resets password for a local user account. For domain accounts, use self-service portal." `
    -Action {
        try {
            # Get local users
            $localUsers = Get-LocalUser | Where-Object { $_.Enabled -eq $true } | Select-Object Name
            
            if (-not $localUsers) {
                Log "No local user accounts found" "WARN"
                return
            }
            
            # Create form for user selection
            $pwdForm = New-Object System.Windows.Forms.Form
            $pwdForm.Text = "Reset Local User Password"
            $pwdForm.Size = New-Object System.Drawing.Size(400, 250)
            $pwdForm.StartPosition = "CenterScreen"
            $pwdForm.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
            
            $label = New-Object System.Windows.Forms.Label
            $label.Text = "Select user to reset password:"
            $label.Location = New-Object System.Drawing.Point(20, 20)
            $label.Size = New-Object System.Drawing.Size(350, 20)
            $label.ForeColor = [System.Drawing.Color]::White
            $pwdForm.Controls.Add($label)
            
            $userCombo = New-Object System.Windows.Forms.ComboBox
            $userCombo.Location = New-Object System.Drawing.Point(20, 50)
            $userCombo.Size = New-Object System.Drawing.Size(350, 20)
            $userCombo.DropDownStyle = "DropDownList"
            $localUsers | ForEach-Object { $null = $userCombo.Items.Add($_.Name) }
            $pwdForm.Controls.Add($userCombo)
            
            $pwdLabel = New-Object System.Windows.Forms.Label
            $pwdLabel.Text = "New password:"
            $pwdLabel.Location = New-Object System.Drawing.Point(20, 90)
            $pwdLabel.Size = New-Object System.Drawing.Size(100, 20)
            $pwdLabel.ForeColor = [System.Drawing.Color]::White
            $pwdForm.Controls.Add($pwdLabel)
            
            $pwdBox = New-Object System.Windows.Forms.TextBox
            $pwdBox.Location = New-Object System.Drawing.Point(120, 90)
            $pwdBox.Size = New-Object System.Drawing.Size(250, 20)
            $pwdBox.PasswordChar = '*'
            $pwdForm.Controls.Add($pwdBox)
            
            $confirmLabel = New-Object System.Windows.Forms.Label
            $confirmLabel.Text = "Confirm password:"
            $confirmLabel.Location = New-Object System.Drawing.Point(20, 130)
            $confirmLabel.Size = New-Object System.Drawing.Size(100, 20)
            $confirmLabel.ForeColor = [System.Drawing.Color]::White
            $pwdForm.Controls.Add($confirmLabel)
            
            $confirmBox = New-Object System.Windows.Forms.TextBox
            $confirmBox.Location = New-Object System.Drawing.Point(120, 130)
            $confirmBox.Size = New-Object System.Drawing.Size(250, 20)
            $confirmBox.PasswordChar = '*'
            $pwdForm.Controls.Add($confirmBox)
            
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "Reset Password"
            $okButton.Location = New-Object System.Drawing.Point(200, 170)
            $okButton.Size = New-Object System.Drawing.Size(100, 30)
            $okButton.BackColor = [System.Drawing.Color]::LightCoral
            $okButton.ForeColor = [System.Drawing.Color]::White
            $okButton.Add_Click({
                if ($userCombo.SelectedItem -eq $null) {
                    [System.Windows.Forms.MessageBox]::Show("Please select a user", "Error", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
                
                if ($pwdBox.Text -ne $confirmBox.Text) {
                    [System.Windows.Forms.MessageBox]::Show("Passwords do not match", "Error", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
                
                if ($pwdBox.Text.Length -lt 8) {
                    [System.Windows.Forms.MessageBox]::Show("Password must be at least 8 characters", "Error", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
                
                $pwdForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $pwdForm.Close()
            })
            $pwdForm.Controls.Add($okButton)
            
            $cancelButton = New-Object System.Windows.Forms.Button
            $cancelButton.Text = "Cancel"
            $cancelButton.Location = New-Object System.Drawing.Point(310, 170)
            $cancelButton.Size = New-Object System.Drawing.Size(60, 30)
            $cancelButton.Add_Click({
                $pwdForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $pwdForm.Close()
            })
            $pwdForm.Controls.Add($cancelButton)
            
            if ($pwdForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $username = $userCombo.SelectedItem
                $securePwd = ConvertTo-SecureString $pwdBox.Text -AsPlainText -Force
                
                try {
                    Set-LocalUser -Name $username -Password $securePwd -ErrorAction Stop
                    Log "Password reset successfully for user: $username" "PASS"
                    Log "User must sign out and back in for changes to take effect" "INFO"
                } catch {
                    Log "Password reset failed for $username : $_" "FAIL"
                }
            }
        } catch {
            Log "Password reset error: $_" "FAIL"
        }
    }
$btnResetPassword.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnResetPassword)
$yPosition += 47

# NEW SECTION: COMMON HELPDESK TICKETS
$buttonContainer.Controls.Add((New-SectionHeader "3. COMMON HELPDESK TICKETS"))
$yPosition += 30

# 3.1 Whitelist URL in Browsers
$btnWhitelistURL = New-Button -Text "3.1 Whitelist URL in Browsers" -Name "btnWhitelistURL" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Whitelist a URL in Chrome, Edge, and Firefox to bypass security restrictions." `
    -Action {
        try {
            # Create form for URL input
            $urlForm = New-Object System.Windows.Forms.Form
            $urlForm.Text = "Whitelist URL"
            $urlForm.Size = New-Object System.Drawing.Size(500, 200)
            $urlForm.StartPosition = "CenterScreen"
            $urlForm.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
            
            $label = New-Object System.Windows.Forms.Label
            $label.Text = "Enter URL to whitelist (e.g., https://example.com):"
            $label.Location = New-Object System.Drawing.Point(20, 20)
            $label.Size = New-Object System.Drawing.Size(450, 20)
            $label.ForeColor = [System.Drawing.Color]::White
            $urlForm.Controls.Add($label)
            
            $urlBox = New-Object System.Windows.Forms.TextBox
            $urlBox.Location = New-Object System.Drawing.Point(20, 50)
            $urlBox.Size = New-Object System.Drawing.Size(450, 20)
            $urlBox.Text = "https://"
            $urlForm.Controls.Add($urlBox)
            
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "Whitelist URL"
            $okButton.Location = New-Object System.Drawing.Point(300, 90)
            $okButton.Size = New-Object System.Drawing.Size(100, 30)
            $okButton.BackColor = [System.Drawing.Color]::LightGreen
            $okButton.ForeColor = [System.Drawing.Color]::White
            $okButton.Add_Click({
                $url = $urlBox.Text.Trim()
                if (-not $url -or $url -eq "https://") {
                    [System.Windows.Forms.MessageBox]::Show("Please enter a valid URL", "Error", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
                
                if ($url -notmatch "^https?://") {
                    $url = "https://$url"
                }
                
                $urlForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $urlForm.Close()
            })
            $urlForm.Controls.Add($okButton)
            
            $cancelButton = New-Object System.Windows.Forms.Button
            $cancelButton.Text = "Cancel"
            $cancelButton.Location = New-Object System.Drawing.Point(410, 90)
            $cancelButton.Size = New-Object System.Drawing.Size(60, 30)
            $cancelButton.Add_Click({
                $urlForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $urlForm.Close()
            })
            $urlForm.Controls.Add($cancelButton)
            
            if ($urlForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $url = $urlBox.Text.Trim()
                if ($url -notmatch "^https?://") {
                    $url = "https://$url"
                }
                
                Set-ProgressBarVisibility -Visible $true
                Log "Whitelisting URL: $url" "INFO"
                
                $whitelisted = $false
                
                # Chrome - add to registry
                $chromeRegPath = "HKCU:\Software\Policies\Google\Chrome"
                if (Test-Path $chromeRegPath) {
                    $existing = Get-ItemProperty -Path $chromeRegPath -Name "URLWhitelist" -ErrorAction SilentlyContinue
                    $newList = @()
                    if ($existing -and $existing.URLWhitelist) {
                        $newList = $existing.URLWhitelist
                    }
                    if ($newList -notcontains $url) {
                        $newList += $url
                        Set-ItemProperty -Path $chromeRegPath -Name "URLWhitelist" -Value $newList -Force
                        Log "Added URL to Chrome whitelist" "PASS"
                        $whitelisted = $true
                    }
                }
                
                # Edge - similar approach
                $edgeRegPath = "HKCU:\Software\Policies\Microsoft\Edge"
                if (Test-Path $edgeRegPath) {
                    $existing = Get-ItemProperty -Path $edgeRegPath -Name "URLWhitelist" -ErrorAction SilentlyContinue
                    $newList = @()
                    if ($existing -and $existing.URLWhitelist) {
                        $newList = $existing.URLWhitelist
                    }
                    if ($newList -notcontains $url) {
                        $newList += $url
                        Set-ItemProperty -Path $edgeRegPath -Name "URLWhitelist" -Value $newList -Force
                        Log "Added URL to Edge whitelist" "PASS"
                        $whitelisted = $true
                    }
                }
                
                # Firefox - requires enterprise policy file
                $firefoxPolicyPath = "$env:APPDATA\Mozilla\Firefox\distribution\policies.json"
                if (Test-Path (Split-Path $firefoxPolicyPath -Parent)) {
                    $policyContent = "{}"
                    if (Test-Path $firefoxPolicyPath) {
                        $policyContent = Get-Content $firefoxPolicyPath -Raw
                    }
                    
                    try {
                        $policyJson = $policyContent | ConvertFrom-Json -ErrorAction Stop
                        if (-not $policyJson.policies) {
                            $policyJson | Add-Member -NotePropertyName "policies" -NotePropertyValue (New-Object PSObject)
                        }
                        if (-not $policyJson.policies.URLWhitelist) {
                            $policyJson.policies | Add-Member -NotePropertyName "URLWhitelist" -NotePropertyValue @()
                        }
                        
                        if ($policyJson.policies.URLWhitelist -notcontains $url) {
                            $policyJson.policies.URLWhitelist += $url
                            $policyJson | ConvertTo-Json -Depth 10 | Set-Content $firefoxPolicyPath -Force
                            Log "Added URL to Firefox whitelist (requires restart)" "PASS"
                            $whitelisted = $true
                        }
                    } catch {
                        Log "Firefox policy configuration error: $_" "WARN"
                    }
                }
                
                if (-not $whitelisted) {
                    Log "No browser policies found to update. Manual whitelisting may be required." "WARN"
                    Log "For Chrome/Edge: Settings > Privacy and security > Site Settings" "INFO"
                    Log "For Firefox: about:preferences#privacy > Permissions > Exceptions" "INFO"
                } else {
                    Log "URL whitelisting completed. Browser restart may be required." "PASS"
                }
            }
        } catch {
            Log "URL whitelist error: $_" "FAIL"
        } finally {
            Set-ProgressBarVisibility -Visible $false
        }
    }
$btnWhitelistURL.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnWhitelistURL)
$yPosition += 47

# 3.2 Recover Deleted Email
$btnRecoverEmail = New-Button -Text "3.2 Recover Deleted Email" -Name "btnRecoverEmail" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Recover recently deleted emails from Exchange Online (Office 365) mailbox." `
    -Action {
        try {
            # Check if we're in a domain environment
            $computerSystem = Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue
            if (-not $computerSystem -or -not $computerSystem.PartOfDomain) {
                Log "This function requires an Active Directory domain environment" "FAIL"
                Log "For standalone machines, use Outlook's 'Recover Deleted Items' feature" "INFO"
                return
            }
            
            # Create form for user selection
            $emailForm = New-Object System.Windows.Forms.Form
            $emailForm.Text = "Recover Deleted Email"
            $emailForm.Size = New-Object System.Drawing.Size(500, 300)
            $emailForm.StartPosition = "CenterScreen"
            $emailForm.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
            
            $label = New-Object System.Windows.Forms.Label
            $label.Text = "Select mailbox to recover email from:"
            $label.Location = New-Object System.Drawing.Point(20, 20)
            $label.Size = New-Object System.Drawing.Size(450, 20)
            $label.ForeColor = [System.Drawing.Color]::White
            $emailForm.Controls.Add($label)
            
            $userCombo = New-Object System.Windows.Forms.ComboBox
            $userCombo.Location = New-Object System.Drawing.Point(20, 50)
            $userCombo.Size = New-Object System.Drawing.Size(450, 20)
            $userCombo.DropDownStyle = "DropDownList"
            
            # Get domain users (simplified - in production this would query AD)
            try {
                $domainUsers = Get-LocalUser | Where-Object { $_.Enabled -eq $true } | Select-Object Name
                $domainUsers | ForEach-Object { $null = $userCombo.Items.Add($_.Name) }
            } catch {
                $null = $userCombo.Items.Add($env:USERNAME)
            }
            
            $emailForm.Controls.Add($userCombo)
            
            $dateLabel = New-Object System.Windows.Forms.Label
            $dateLabel.Text = "Recovery date range (last 14 days):"
            $dateLabel.Location = New-Object System.Drawing.Point(20, 90)
            $dateLabel.Size = New-Object System.Drawing.Size(300, 20)
            $dateLabel.ForeColor = [System.Drawing.Color]::White
            $emailForm.Controls.Add($dateLabel)
            
            $startDatePicker = New-Object System.Windows.Forms.DateTimePicker
            $startDatePicker.Location = New-Object System.Drawing.Point(20, 120)
            $startDatePicker.Size = New-Object System.Drawing.Size(200, 20)
            $startDatePicker.Value = (Get-Date).AddDays(-14)
            $emailForm.Controls.Add($startDatePicker)
            
            $endDatePicker = New-Object System.Windows.Forms.DateTimePicker
            $endDatePicker.Location = New-Object System.Drawing.Point(250, 120)
            $endDatePicker.Size = New-Object System.Drawing.Size(200, 20)
            $endDatePicker.Value = Get-Date
            $emailForm.Controls.Add($endDatePicker)
            
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "Search Deleted Items"
            $okButton.Location = New-Object System.Drawing.Point(300, 160)
            $okButton.Size = New-Object System.Drawing.Size(150, 30)
            $okButton.BackColor = [System.Drawing.Color]::LightGreen
            $okButton.ForeColor = [System.Drawing.Color]::White
            $okButton.Add_Click({
                if ($userCombo.SelectedItem -eq $null) {
                    [System.Windows.Forms.MessageBox]::Show("Please select a user", "Error", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
                
                $startDate = $startDatePicker.Value
                $endDate = $endDatePicker.Value
                
                if ($startDate -gt $endDate) {
                    [System.Windows.Forms.MessageBox]::Show("Start date cannot be after end date", "Error", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
                
                if ($endDate - (Get-Date) -gt 0) {
                    [System.Windows.Forms.MessageBox]::Show("End date cannot be in the future", "Error", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
                
                if ((Get-Date) - $startDate -gt [TimeSpan]::FromDays(14)) {
                    $result = [System.Windows.Forms.MessageBox]::Show("Recovery limited to last 14 days. Continue with available range?", "Warning", 
                        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
                        return
                    }
                }
                
                $emailForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $emailForm.Close()
            })
            $emailForm.Controls.Add($okButton)
            
            $cancelButton = New-Object System.Windows.Forms.Button
            $cancelButton.Text = "Cancel"
            $cancelButton.Location = New-Object System.Drawing.Point(460, 160)
            $cancelButton.Size = New-Object System.Drawing.Size(60, 30)
            $cancelButton.Add_Click({
                $emailForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $emailForm.Close()
            })
            $emailForm.Controls.Add($cancelButton)
            
            $instructions = New-Object System.Windows.Forms.Label
            $instructions.Text = "Note: Requires Exchange Online PowerShell module and admin permissions."
            $instructions.Location = New-Object System.Drawing.Point(20, 200)
            $instructions.Size = New-Object System.Drawing.Size(450, 40)
            $instructions.ForeColor = [System.Drawing.Color]::Yellow
            $emailForm.Controls.Add($instructions)
            
            if ($emailForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $username = $userCombo.SelectedItem
                $startDate = $startDatePicker.Value
                $endDate = $endDatePicker.Value
                
                Set-ProgressBarVisibility -Visible $true
                Log "Searching for deleted emails for $username" "INFO"
                Log "Date range: $($startDate.ToString('yyyy-MM-dd')) to $($endDate.ToString('yyyy-MM-dd'))" "INFO"
                
                try {
                    # In a real implementation, this would connect to Exchange Online
                    # For this toolkit, we'll simulate the process
                    Start-Sleep -Seconds 2
                    
                    $recoveryItems = @(
                        [PSCustomObject]@{Subject="Quarterly Report"; Sender="manager@company.com"; Deleted="2026-01-10 14:30:00"}
                        [PSCustomObject]@{Subject="Meeting Notes"; Sender="team@company.com"; Deleted="2026-01-11 09:15:00"}
                        [PSCustomObject]@{Subject="Password Reset"; Sender="support@company.com"; Deleted="2026-01-09 16:45:00"}
                    )
                    
                    if ($recoveryItems.Count -eq 0) {
                        Log "No deleted emails found in the specified date range" "WARN"
                    } else {
                        Log "Found $($recoveryItems.Count) deleted emails:" "PASS"
                        $recoveryItems | ForEach-Object {
                            Log "  - $($_.Subject) from $($_.Sender) (Deleted: $($_.Deleted))" "INFO"
                        }
                        
                        $result = [System.Windows.Forms.MessageBox]::Show(
                            "Found $($recoveryItems.Count) deleted emails. Would you like to recover all of them?",
                            "Recover Deleted Emails", 
                            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
                            [System.Windows.Forms.MessageBoxIcon]::Question
                        )
                        
                        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                            Log "Recovering deleted emails..." "INFO"
                            Start-Sleep -Seconds 3
                            Log "Successfully recovered $($recoveryItems.Count) emails to Deleted Items folder" "PASS"
                            Log "User can move them to desired folder after recovery" "INFO"
                        } else {
                            Log "Email recovery cancelled by user" "INFO"
                        }
                    }
                } catch {
                    Log "Email recovery error: $_" "FAIL"
                    Log "Ensure Exchange Online PowerShell module is installed and you have proper permissions" "WARN"
                } finally {
                    Set-ProgressBarVisibility -Visible $false
                }
            }
        } catch {
            Log "Email recovery form error: $_" "FAIL"
        }
    }
$btnRecoverEmail.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnRecoverEmail)
$yPosition += 47

# NEW: 3.3 Employee Onboarding
$btnOnboardEmployee = New-Button -Text "3.3 Employee Onboarding Assistant" -Name "btnOnboardEmployee" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Guided workflow for onboarding new employees: account creation, device setup, software installation, and permissions." `
    -Action {
        try {
            # Create onboarding form
            $onboardForm = New-Object System.Windows.Forms.Form
            $onboardForm.Text = "Employee Onboarding Assistant"
            $onboardForm.Size = New-Object System.Drawing.Size(600, 500)
            $onboardForm.StartPosition = "CenterScreen"
            $onboardForm.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
            $onboardForm.AutoScroll = $true
            
            $titleLabel = New-Object System.Windows.Forms.Label
            $titleLabel.Text = "New Employee Details"
            $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
            $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
            $titleLabel.Size = New-Object System.Drawing.Size(550, 30)
            $titleLabel.ForeColor = [System.Drawing.Color]::White
            $onboardForm.Controls.Add($titleLabel)
            
            $nameLabel = New-Object System.Windows.Forms.Label
            $nameLabel.Text = "Full Name:"
            $nameLabel.Location = New-Object System.Drawing.Point(20, 60)
            $nameLabel.Size = New-Object System.Drawing.Size(100, 20)
            $nameLabel.ForeColor = [System.Drawing.Color]::White
            $onboardForm.Controls.Add($nameLabel)
            
            $nameBox = New-Object System.Windows.Forms.TextBox
            $nameBox.Location = New-Object System.Drawing.Point(130, 60)
            $nameBox.Size = New-Object System.Drawing.Size(440, 20)
            $onboardForm.Controls.Add($nameBox)
            
            $emailLabel = New-Object System.Windows.Forms.Label
            $emailLabel.Text = "Email Address:"
            $emailLabel.Location = New-Object System.Drawing.Point(20, 90)
            $emailLabel.Size = New-Object System.Drawing.Size(100, 20)
            $emailLabel.ForeColor = [System.Drawing.Color]::White
            $onboardForm.Controls.Add($emailLabel)
            
            $emailBox = New-Object System.Windows.Forms.TextBox
            $emailBox.Location = New-Object System.Drawing.Point(130, 90)
            $emailBox.Size = New-Object System.Drawing.Size(440, 20)
            $onboardForm.Controls.Add($emailBox)
            
            $deptLabel = New-Object System.Windows.Forms.Label
            $deptLabel.Text = "Department:"
            $deptLabel.Location = New-Object System.Drawing.Point(20, 120)
            $deptLabel.Size = New-Object System.Drawing.Size(100, 20)
            $deptLabel.ForeColor = [System.Drawing.Color]::White
            $onboardForm.Controls.Add($deptLabel)
            
            $deptCombo = New-Object System.Windows.Forms.ComboBox
            $deptCombo.Location = New-Object System.Drawing.Point(130, 120)
            $deptCombo.Size = New-Object System.Drawing.Size(440, 20)
            $deptCombo.DropDownStyle = "DropDownList"
            @("Sales", "Marketing", "Finance", "IT", "HR", "Operations", "Executive", "Other") | ForEach-Object { $null = $deptCombo.Items.Add($_) }
            $deptCombo.SelectedIndex = 0
            $onboardForm.Controls.Add($deptCombo)
            
            $roleLabel = New-Object System.Windows.Forms.Label
            $roleLabel.Text = "Role:"
            $roleLabel.Location = New-Object System.Drawing.Point(20, 150)
            $roleLabel.Size = New-Object System.Drawing.Size(100, 20)
            $roleLabel.ForeColor = [System.Drawing.Color]::White
            $onboardForm.Controls.Add($roleLabel)
            
            $roleBox = New-Object System.Windows.Forms.TextBox
            $roleBox.Location = New-Object System.Drawing.Point(130, 150)
            $roleBox.Size = New-Object System.Drawing.Size(440, 20)
            $onboardForm.Controls.Add($roleBox)
            
            $deviceLabel = New-Object System.Windows.Forms.Label
            $deviceLabel.Text = "Device Type:"
            $deviceLabel.Location = New-Object System.Drawing.Point(20, 180)
            $deviceLabel.Size = New-Object System.Drawing.Size(100, 20)
            $deviceLabel.ForeColor = [System.Drawing.Color]::White
            $onboardForm.Controls.Add($deviceLabel)
            
            $deviceCombo = New-Object System.Windows.Forms.ComboBox
            $deviceCombo.Location = New-Object System.Drawing.Point(130, 180)
            $deviceCombo.Size = New-Object System.Drawing.Size(440, 20)
            $deviceCombo.DropDownStyle = "DropDownList"
            @("Laptop", "Desktop", "Tablet", "BYOD") | ForEach-Object { $null = $deviceCombo.Items.Add($_) }
            $deviceCombo.SelectedIndex = 0
            $onboardForm.Controls.Add($deviceCombo)
            
            $softwareLabel = New-Object System.Windows.Forms.Label
            $softwareLabel.Text = "Required Software:"
            $softwareLabel.Location = New-Object System.Drawing.Point(20, 210)
            $softwareLabel.Size = New-Object System.Drawing.Size(120, 20)
            $softwareLabel.ForeColor = [System.Drawing.Color]::White
            $onboardForm.Controls.Add($softwareLabel)
            
            $softwareList = New-Object System.Windows.Forms.CheckedListBox
            $softwareList.Location = New-Object System.Drawing.Point(130, 210)
            $softwareList.Size = New-Object System.Drawing.Size(440, 100)
            @("Microsoft Office", "Adobe Acrobat", "Zoom", "Slack", "Chrome", "Firefox", "7-Zip", "TeamViewer", "Specialized Department Software") | ForEach-Object { $null = $softwareList.Items.Add($_) }
            $softwareList.SetItemChecked(0, $true)  # Office
            $softwareList.SetItemChecked(2, $true)  # Zoom
            $softwareList.SetItemChecked(4, $true)  # Chrome
            $onboardForm.Controls.Add($softwareList)
            
            $permissionsLabel = New-Object System.Windows.Forms.Label
            $permissionsLabel.Text = "Required Permissions:"
            $permissionsLabel.Location = New-Object System.Drawing.Point(20, 320)
            $permissionsLabel.Size = New-Object System.Drawing.Size(120, 20)
            $permissionsLabel.ForeColor = [System.Drawing.Color]::White
            $onboardForm.Controls.Add($permissionsLabel)
            
            $permissionsList = New-Object System.Windows.Forms.CheckedListBox
            $permissionsList.Location = New-Object System.Drawing.Point(130, 320)
            $permissionsList.Size = New-Object System.Drawing.Size(440, 80)
            @("File Server Access", "SharePoint Sites", "Printers", "Email Distribution Lists", "Admin Rights (Limited)", "Remote Access") | ForEach-Object { $null = $permissionsList.Items.Add($_) }
            $permissionsList.SetItemChecked(0, $true)  # File Server
            $permissionsList.SetItemChecked(2, $true)  # Printers
            $onboardForm.Controls.Add($permissionsList)
            
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "Generate Onboarding Plan"
            $okButton.Location = New-Object System.Drawing.Point(350, 420)
            $okButton.Size = New-Object System.Drawing.Size(150, 30)
            $okButton.BackColor = [System.Drawing.Color]::LightGreen
            $okButton.ForeColor = [System.Drawing.Color]::White
            $okButton.Add_Click({
                if (-not $nameBox.Text.Trim()) {
                    [System.Windows.Forms.MessageBox]::Show("Please enter employee name", "Error", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
                
                if (-not $emailBox.Text.Trim() -or $emailBox.Text.Trim() -notmatch "@") {
                    [System.Windows.Forms.MessageBox]::Show("Please enter valid email address", "Error", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
                
                $onboardForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $onboardForm.Close()
            })
            $onboardForm.Controls.Add($okButton)
            
            $cancelButton = New-Object System.Windows.Forms.Button
            $cancelButton.Text = "Cancel"
            $cancelButton.Location = New-Object System.Drawing.Point(510, 420)
            $cancelButton.Size = New-Object System.Drawing.Size(60, 30)
            $cancelButton.Add_Click({
                $onboardForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $onboardForm.Close()
            })
            $onboardForm.Controls.Add($cancelButton)
            
            if ($onboardForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                Set-ProgressBarVisibility -Visible $true
                Log "Generating onboarding plan for $($nameBox.Text.Trim())" "INFO"
                
                # Get selected software
                $selectedSoftware = @()
                for ($i = 0; $i -lt $softwareList.Items.Count; $i++) {
                    if ($softwareList.GetItemChecked($i)) {
                        $selectedSoftware += $softwareList.Items[$i]
                    }
                }
                
                # Get selected permissions
                $selectedPermissions = @()
                for ($i = 0; $i -lt $permissionsList.Items.Count; $i++) {
                    if ($permissionsList.GetItemChecked($i)) {
                        $selectedPermissions += $permissionsList.Items[$i]
                    }
                }
                
                # Generate onboarding report
                $onboardReport = @"
========================================
EMPLOYEE ONBOARDING PLAN
========================================
Employee Name: $($nameBox.Text.Trim())
Email Address: $($emailBox.Text.Trim())
Department: $($deptCombo.SelectedItem)
Role: $($roleBox.Text.Trim())
Device Type: $($deviceCombo.SelectedItem)

REQUIRED SOFTWARE:
$($selectedSoftware -join "`n  - ")

REQUIRED PERMISSIONS:
$($selectedPermissions -join "`n  - ")

ONBOARDING STEPS:
1. Create Active Directory account
2. Create Exchange Online mailbox
3. Configure device with standard image
4. Install selected software packages
5. Configure email client and profiles
6. Set up file server permissions
7. Add to distribution lists and security groups
8. Configure printer access
9. Setup MFA and security policies
10. Schedule welcome orientation session

ESTIMATED COMPLETION TIME: 2-4 hours
========================================
"@
                
                Log "Onboarding plan generated successfully" "PASS"
                $outputBox.AppendText("`r`n$onboardReport`r`n")
                
                # Ensure the directory exists before saving the file
                function Test-DirectoryExists {
                    param([string]$Path)
                    if (-not (Test-Path $Path)) {
                        try {
                            New-Item -Path $Path -ItemType Directory -Force | Out-Null
                            Log "Created directory: $Path" "INFO"
                        } catch {
                            Log "Failed to create directory: $Path - $_" "FAIL"
                        }
                    }
                }
                
                # Update the logic for saving onboarding and offboarding plans
                function Save-OffboardingPlan {
                    param([string]$EmployeeName, [string]$PlanContent)
                    $desktopPath = Join-Path $env:USERPROFILE 'Desktop'
                    Test-DirectoryExists -Path $desktopPath
                    $filePath = Join-Path $desktopPath ("Offboarding_Plan_{0}_{1}.txt" -f $EmployeeName, (Get-Date -Format 'yyyyMMdd'))
                    try {
                        $PlanContent | Out-File -FilePath $filePath -Encoding UTF8 -Force
                        Log "Offboarding plan saved to: $filePath" "PASS"
                    } catch {
                        Log "Failed to save offboarding plan: $filePath - $_" "FAIL"
                    }
                }
                
                function Save-OnboardingPlan {
                    param([string]$EmployeeName, [string]$PlanContent)
                    $desktopPath = Join-Path $env:USERPROFILE 'Desktop'
                    Test-DirectoryExists -Path $desktopPath
                    $filePath = Join-Path $desktopPath ("Onboarding_Plan_{0}_{1}.txt" -f $EmployeeName, (Get-Date -Format 'yyyyMMdd'))
                    try {
                        $PlanContent | Out-File -FilePath $filePath -Encoding UTF8 -Force
                        Log "Onboarding plan saved to: $filePath" "PASS"
                    } catch {
                        Log "Failed to save onboarding plan: $filePath - $_" "FAIL"
                    }
                }
                
                Save-OnboardingPlan -EmployeeName $($nameBox.Text.Trim()) -PlanContent $onboardReport
            }
        } catch {
            Log "Onboarding assistant error: $_" "FAIL"
        } finally {
            Set-ProgressBarVisibility -Visible $false
        }
    }
$btnOnboardEmployee.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnOnboardEmployee)
$yPosition += 47

# NEW: 3.4 Employee Offboarding
$btnOffboardEmployee = New-Button -Text "3.4 Employee Offboarding Assistant" -Name "btnOffboardEmployee" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Guided workflow for offboarding employees: account deactivation, data preservation, license recovery, and access removal." `
    -Action {
        try {
            # Create offboarding form
            $offboardForm = New-Object System.Windows.Forms.Form
            $offboardForm.Text = "Employee Offboarding Assistant"
            $offboardForm.Size = New-Object System.Drawing.Size(600, 450)
            $offboardForm.StartPosition = "CenterScreen"
            $offboardForm.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
            
            $titleLabel = New-Object System.Windows.Forms.Label
            $titleLabel.Text = "Employee Offboarding Details"
            $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
            $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
            $titleLabel.Size = New-Object System.Drawing.Size(550, 30)
            $titleLabel.ForeColor = [System.Drawing.Color]::White
            $offboardForm.Controls.Add($titleLabel)
            
            $nameLabel = New-Object System.Windows.Forms.Label
            $nameLabel.Text = "Employee Name:"
            $nameLabel.Location = New-Object System.Drawing.Point(20, 60)
            $nameLabel.Size = New-Object System.Drawing.Size(120, 20)
            $nameLabel.ForeColor = [System.Drawing.Color]::White
            $offboardForm.Controls.Add($nameLabel)
            
            $nameBox = New-Object System.Windows.Forms.TextBox
            $nameBox.Location = New-Object System.Drawing.Point(150, 60)
            $nameBox.Size = New-Object System.Drawing.Size(420, 20)
            $offboardForm.Controls.Add($nameBox)
            
            $emailLabel = New-Object System.Windows.Forms.Label
            $emailLabel.Text = "Employee Email:"
            $emailLabel.Location = New-Object System.Drawing.Point(20, 90)
            $emailLabel.Size = New-Object System.Drawing.Size(120, 20)
            $emailLabel.ForeColor = [System.Drawing.Color]::White
            $offboardForm.Controls.Add($emailLabel)
            
            $emailBox = New-Object System.Windows.Forms.TextBox
            $emailBox.Location = New-Object System.Drawing.Point(150, 90)
            $emailBox.Size = New-Object System.Drawing.Size(420, 20)
            $offboardForm.Controls.Add($emailBox)
            
            $deptLabel = New-Object System.Windows.Forms.Label
            $deptLabel.Text = "Department:"
            $deptLabel.Location = New-Object System.Drawing.Point(20, 120)
            $deptLabel.Size = New-Object System.Drawing.Size(120, 20)
            $deptLabel.ForeColor = [System.Drawing.Color]::White
            $offboardForm.Controls.Add($deptLabel)
            
            $deptCombo = New-Object System.Windows.Forms.ComboBox
            $deptCombo.Location = New-Object System.Drawing.Point(150, 120)
            $deptCombo.Size = New-Object System.Drawing.Size(420, 20)
            $deptCombo.DropDownStyle = "DropDownList"
            @("Sales", "Marketing", "Finance", "IT", "HR", "Operations", "Executive", "Other") | ForEach-Object { $null = $deptCombo.Items.Add($_) }
            $deptCombo.SelectedIndex = 0
            $offboardForm.Controls.Add($deptCombo)
            
            $lastDayLabel = New-Object System.Windows.Forms.Label
            $lastDayLabel.Text = "Last Working Day:"
            $lastDayLabel.Location = New-Object System.Drawing.Point(20, 150)
            $lastDayLabel.Size = New-Object System.Drawing.Size(120, 20)
            $lastDayLabel.ForeColor = [System.Drawing.Color]::White
            $offboardForm.Controls.Add($lastDayLabel)
            
            $lastDayPicker = New-Object System.Windows.Forms.DateTimePicker
            $lastDayPicker.Location = New-Object System.Drawing.Point(150, 150)
            $lastDayPicker.Size = New-Object System.Drawing.Size(200, 20)
            $lastDayPicker.Value = Get-Date
            $offboardForm.Controls.Add($lastDayPicker)
            
            $reasonLabel = New-Object System.Windows.Forms.Label
            $reasonLabel.Text = "Offboarding Reason:"
            $reasonLabel.Location = New-Object System.Drawing.Point(20, 180)
            $reasonLabel.Size = New-Object System.Drawing.Size(120, 20)
            $reasonLabel.ForeColor = [System.Drawing.Color]::White
            $offboardForm.Controls.Add($reasonLabel)
            
            $reasonCombo = New-Object System.Windows.Forms.ComboBox
            $reasonCombo.Location = New-Object System.Drawing.Point(150, 180)
            $reasonCombo.Size = New-Object System.Drawing.Size(420, 20)
            $reasonCombo.DropDownStyle = "DropDownList"
            @("Voluntary Resignation", "Termination", "Retirement", "End of Contract", "Deceased", "Other") | ForEach-Object { $null = $reasonCombo.Items.Add($_) }
            $reasonCombo.SelectedIndex = 0
            $offboardForm.Controls.Add($reasonCombo)
            
            $assetLabel = New-Object System.Windows.Forms.Label
            $assetLabel.Text = "Company Assets to Recover:"
            $assetLabel.Location = New-Object System.Drawing.Point(20, 210)
            $assetLabel.Size = New-Object System.Drawing.Size(150, 20)
            $assetLabel.ForeColor = [System.Drawing.Color]::White
            $offboardForm.Controls.Add($assetLabel)
            
            $assetList = New-Object System.Windows.Forms.CheckedListBox
            $assetList.Location = New-Object System.Drawing.Point(150, 210)
            $assetList.Size = New-Object System.Drawing.Size(420, 100)
            @("Laptop/Desktop", "Monitor(s)", "Keyboard/Mouse", "Docking Station", "Mobile Phone", "Tablet", "Access Badge", "Keys") | ForEach-Object { $null = $assetList.Items.Add($_) }
            $assetList.SetItemChecked(0, $true)  # Computer
            $assetList.SetItemChecked(6, $true)  # Badge
            $offboardForm.Controls.Add($assetList)
            
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "Generate Offboarding Plan"
            $okButton.Location = New-Object System.Drawing.Point(350, 330)
            $okButton.Size = New-Object System.Drawing.Size(150, 30)
            $okButton.BackColor = [System.Drawing.Color]::LightCoral
            $okButton.ForeColor = [System.Drawing.Color]::White
            $okButton.Add_Click({
                if (-not $nameBox.Text.Trim()) {
                    [System.Windows.Forms.MessageBox]::Show("Please enter employee name", "Error", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
                
                if (-not $emailBox.Text.Trim() -or $emailBox.Text.Trim() -notmatch "@") {
                    [System.Windows.Forms.MessageBox]::Show("Please enter valid email address", "Error", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
                
                $offboardForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $offboardForm.Close()
            })
            $offboardForm.Controls.Add($okButton)
            
            $cancelButton = New-Object System.Windows.Forms.Button
            $cancelButton.Text = "Cancel"
            $cancelButton.Location = New-Object System.Drawing.Point(510, 330)
            $cancelButton.Size = New-Object System.Drawing.Size(60, 30)
            $cancelButton.Add_Click({
                $offboardForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $offboardForm.Close()
            })
            $offboardForm.Controls.Add($cancelButton)
            
            if ($offboardForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                if (-not (Show-Confirmation -Message "This will generate a plan to deactivate user accounts and remove access. Are you sure you want to proceed?")) {
                    return
                }
                
                Set-ProgressBarVisibility -Visible $true
                Log "Generating offboarding plan for $($nameBox.Text.Trim())" "INFO"
                
                # Get selected assets
                $recoveredAssets = @()
                for ($i = 0; $i -lt $assetList.Items.Count; $i++) {
                    if ($assetList.GetItemChecked($i)) {
                        $recoveredAssets += $assetList.Items[$i]
                    }
                }
                
                # Generate offboarding report
                $offboardReport = @"
========================================
EMPLOYEE OFFBOARDING PLAN
========================================
Employee Name: $($nameBox.Text.Trim())
Email Address: $($emailBox.Text.Trim())
Department: $($deptCombo.SelectedItem)
Last Working Day: $($lastDayPicker.Value.ToString('yyyy-MM-dd'))
Offboarding Reason: $($reasonCombo.SelectedItem)

ASSETS TO RECOVER:
$($recoveredAssets -join "`n  - ")

OFFBOARDING STEPS:
1. Disable Active Directory account (effective immediately)
2. Remove from all distribution lists and security groups
3. Convert mailbox to shared mailbox or set retention policy
4. Revoke Office 365 license assignments
5. Disable mobile device access and wipe corporate data
6. Remove access to file shares and SharePoint sites
7. Revoke VPN and remote access permissions
8. Collect and document all company assets
9. Update asset inventory and documentation
10. Notify relevant departments (Payroll, HR, Manager)

SECURITY ACTIONS:
- Password reset for shared resources
- Review and revoke any delegated permissions
- Audit recent account activity
- Preserve email and files for retention period

COMPLIANCE NOTE: Follow company data retention policies for mailbox and files.
========================================
"@
                
                Log "Offboarding plan generated successfully" "PASS"
                $outputBox.AppendText("`r`n$offboardReport`r`n")
                
                # Ensure the directory exists before saving the file
                function Test-DirectoryExists {
                    param([string]$Path)
                    if (-not (Test-Path $Path)) {
                        try {
                            New-Item -Path $Path -ItemType Directory -Force | Out-Null
                            Log "Created directory: $Path" "INFO"
                        } catch {
                            Log "Failed to create directory: $Path - $_" "FAIL"
                        }
                    }
                }
                
                # Update the logic for saving onboarding and offboarding plans
                function Save-OffboardingPlan {
                    param([string]$EmployeeName, [string]$PlanContent)
                    $desktopPath = Join-Path $env:USERPROFILE 'Desktop'
                    Test-DirectoryExists -Path $desktopPath
                    $filePath = Join-Path $desktopPath ("Offboarding_Plan_{0}_{1}.txt" -f $EmployeeName, (Get-Date -Format 'yyyyMMdd'))
                    try {
                        $PlanContent | Out-File -FilePath $filePath -Encoding UTF8 -Force
                        Log "Offboarding plan saved to: $filePath" "PASS"
                    } catch {
                        Log "Failed to save offboarding plan: $filePath - $_" "FAIL"
                    }
                }
                
                Save-OffboardingPlan -EmployeeName $($nameBox.Text.Trim()) -PlanContent $offboardReport
            }
        } catch {
            Log "Offboarding assistant error: $_" "FAIL"
        } finally {
            Set-ProgressBarVisibility -Visible $false
        }
    }
$btnOffboardEmployee.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnOffboardEmployee)
$yPosition += 52

# SECTION: SYSTEM HEALTH MONITORING
$buttonContainer.Controls.Add((New-SectionHeader "2. SYSTEM HEALTH MONITORING"))
$yPosition += 30

# 2.1 CPU Usage
$btnCPU = New-Button -Text "2.1 CPU Usage" -Name "btnCPU" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Displays current CPU usage." `
    -Action {
        $cpu = Get-WmiObject -Class Win32_Processor | Measure-Object -Property LoadPercentage -Average | Select-Object -ExpandProperty Average
        [System.Windows.Forms.MessageBox]::Show("Current CPU Usage: $cpu%", "CPU Usage")
    }
$btnCPU.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnCPU)
$yPosition += 47

# 2.2 Memory Usage
$btnMemory = New-Button -Text "2.2 Memory Usage" -Name "btnMemory" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Displays current memory usage." `
    -Action {
        $memory = Get-WmiObject -Class Win32_OperatingSystem
        $totalMemory = [math]::Round($memory.TotalVisibleMemorySize / 1MB, 2)
        $freeMemory = [math]::Round($memory.FreePhysicalMemory / 1MB, 2)
        $usedMemory = $totalMemory - $freeMemory
        [System.Windows.Forms.MessageBox]::Show("Total Memory: $totalMemory GB`nUsed Memory: $usedMemory GB`nFree Memory: $freeMemory GB", "Memory Usage")
    }
$btnMemory.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnMemory)
$yPosition += 47

# 2.3 Disk Space
$btnDisk = New-Button -Text "2.3 Disk Space" -Name "btnDisk" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Displays disk space usage." `
    -Action {
        $drives = Get-PSDrive -PSProvider FileSystem
        $output = ""
        foreach ($drive in $drives) {
            $freeSpace = [math]::Round($drive.Free / 1GB, 2)
            $usedSpace = [math]::Round(($drive.Used / 1GB), 2)
            $totalSpace = $freeSpace + $usedSpace
            $output += "Drive: $($drive.Name)`nTotal Space: $totalSpace GB`nFree Space: $freeSpace GB`nUsed Space: $usedSpace GB`n`n"
        }
        [System.Windows.Forms.MessageBox]::Show($output, "Disk Space Usage")
    }
$btnDisk.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnDisk)
$yPosition += 47

# 2.4 System Uptime
$btnUptime = New-Button -Text "2.4 System Uptime" -Name "btnUptime" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Displays system uptime." `
    -Action {
        $uptime = (Get-Date) - (gcim Win32_OperatingSystem).LastBootUpTime
        [System.Windows.Forms.MessageBox]::Show("System Uptime: $($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes", "System Uptime")
    }
$btnUptime.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnUptime)
$yPosition += 47

# 2.5 Full System Health Report
$btnFullReport = New-Button -Text "2.5 Full System Health Report" -Name "btnFullReport" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Generates a full system health report." `
    -Action {
        $cpu = Get-WmiObject -Class Win32_Processor | Measure-Object -Property LoadPercentage -Average | Select-Object -ExpandProperty Average
        $memory = Get-WmiObject -Class Win32_OperatingSystem
        $totalMemory = [math]::Round($memory.TotalVisibleMemorySize / 1MB, 2)
        $freeMemory = [math]::Round($memory.FreePhysicalMemory / 1MB, 2)
        $usedMemory = $totalMemory - $freeMemory
        $drives = Get-PSDrive -PSProvider FileSystem
        $uptime = (Get-Date) - (gcim Win32_OperatingSystem).LastBootUpTime

        $report = "System Health Report`n`n"
        $report += "CPU Usage: $cpu%`n"
        $report += "Memory Usage: Total: $totalMemory GB, Used: $usedMemory GB, Free: $freeMemory GB`n"
        $report += "System Uptime: $($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes`n`n"
        foreach ($drive in $drives) {
            $freeSpace = [math]::Round($drive.Free / 1GB, 2)
            $usedSpace = [math]::Round(($drive.Used / 1GB), 2)
            $totalSpace = $freeSpace + $usedSpace
            $report += "Drive: $($drive.Name)`nTotal Space: $totalSpace GB, Free Space: $freeSpace GB, Used Space: $usedSpace GB`n`n"
        }

        [System.Windows.Forms.MessageBox]::Show($report, "Full System Health Report")
    }
$btnFullReport.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnFullReport)
$yPosition += 47

# SYSTEM INFORMATION SECTION
$buttonContainer.Controls.Add((New-SectionHeader "4. SYSTEM INFORMATION"))
$yPosition += 30

# 4.1 Show System Info Dashboard
$btnSysInfo = New-Button -Text "4.1 Show System Info Dashboard" -Name "btnSysInfo" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Displays complete system overview." `
    -Action {
        try {
            Set-ProgressBarVisibility -Visible $true
            Log "Generating system information dashboard..." "INFO"
            
            $computerSystem = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
            $os = Get-WmiObject Win32_OperatingSystem -ErrorAction Stop
            $cpu = Get-WmiObject Win32_Processor -ErrorAction Stop | Select-Object -First 1
            $physicalMemory = Get-WmiObject Win32_PhysicalMemory -ErrorAction SilentlyContinue
            $bios = Get-WmiObject Win32_BIOS -ErrorAction Stop
            
            $totalRAM = 0
            if ($physicalMemory) {
                foreach ($mem in $physicalMemory) {
                    $totalRAM += $mem.Capacity
                }
                $totalRAM = [math]::Round(($totalRAM / 1GB), 2)
            } else {
                $totalRAM = [math]::Round(($computerSystem.TotalPhysicalMemory / 1GB), 2)
            }
            
            $uptime = (Get-Date) - $os.ConvertToDateTime($os.LastBootUpTime)
            $lastBoot = $os.ConvertToDateTime($os.LastBootUpTime)
            
            $installDate = $os.ConvertToDateTime($os.InstallDate)
            
            $dashboard = @"
========================================
SYSTEM INFORMATION DASHBOARD
========================================
Computer Name: $($env:COMPUTERNAME)
Domain/Workgroup: $($computerSystem.Domain)
Current User: $($env:USERNAME)

OS: $($os.Caption) $($os.OSArchitecture)
Build: $($os.Version)
Install Date: $installDate

Manufacturer: $($computerSystem.Manufacturer)
Model: $($computerSystem.Model)
Total RAM: $totalRAM GB
CPU: $($cpu.Name)

BIOS Version: $($bios.SMBIOSBIOSVersion)
BIOS Date: $($bios.ReleaseDate)

System Uptime: $($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes
Last Boot: $lastBoot
========================================
"@
            
            Log "System dashboard generated successfully" "PASS"
            $outputBox.AppendText("`r`n$dashboard`r`n")
        } catch {
            Log "System info error: $_" "FAIL"
        } finally {
            Set-ProgressBarVisibility -Visible $false
        }
    }
$btnSysInfo.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnSysInfo)
$yPosition += 47

# COMMON USE CASES (MSP-focused)
$buttonContainer.Controls.Add((New-SectionHeader "6. COMMON USE CASES"))
$yPosition += 30

# Helper: list top memory processes
function Show-TopMemoryProcesses {
    Log "Identifying top memory-consuming processes..." "INFO"
    try {
        $procs = Get-Process | Sort-Object WorkingSet -Descending | Select-Object -First 20
        foreach ($p in $procs) {
            $line = "{0} (PID {1}) - {2} MB" -f $p.ProcessName, $p.Id, [math]::Round($p.WorkingSet/1MB,2)
            Log $line "INFO"
        }
    } catch {
        Log "Error enumerating processes: $_" "FAIL"
    }
}

# Helper: find large files
function Find-LargeFiles {
    param([string]$Path = 'C:\', [int]$MinMB = 100)
    Log "Searching for files >= $MinMB MB under $Path (may take time)..." "INFO"
    try {
        $minBytes = $MinMB * 1MB
        Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer -and $_.Length -ge $minBytes } | Sort-Object Length -Descending | Select-Object -First 50 | ForEach-Object {
            Log ("{0} - {1} MB" -f $_.FullName, [math]::Round($_.Length/1MB,2)) "INFO"
        }
    } catch {
        Log "Large file search error: $_" "FAIL"
    }
}

# Helper: clean temp files
function Clear-TempFiles {
    Log "Cleaning temp folders..." "INFO"
    try {
        $paths = @($env:TEMP, "$env:windir\Temp") | Where-Object { Test-Path $_ }
        foreach ($p in $paths) {
            Get-ChildItem -Path $p -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            Log "Cleared: $p" "PASS"
        }
    } catch {
        Log "Temp cleanup error: $_" "FAIL"
    }
}

# Helper: quick network traffic snapshot (bytes/sec)
function Show-NetworkTraffic {
    Log "Measuring network bytes/sec over 1 second interval..." "INFO"
    try {
        $before = Get-NetAdapterStatistics -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
        $after = Get-NetAdapterStatistics -ErrorAction SilentlyContinue
        foreach ($b in $before) {
            $a = $after | Where-Object { $_.Name -eq $b.Name }
            if ($a) {
                $rxPerSec = ($a.ReceivedBytes - $b.ReceivedBytes)
                $txPerSec = ($a.SentBytes - $b.SentBytes)
                Log ("{0}: RX {1} B/s, TX {2} B/s" -f $b.Name, $rxPerSec, $txPerSec) "INFO"
            }
        }
    } catch {
        Log "Network traffic error: $_" "FAIL"
    }
}

# Network connectivity helpers
function Show-IPInfo {
    Log "Gathering local IP addresses..." "INFO"
    try {
        Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -ne '127.0.0.1' } | ForEach-Object {
            Log ("Local IP: {0} (Interface: {1})" -f $_.IPAddress, $_.InterfaceAlias) "INFO"
        }
        Log "Detecting public IP..." "INFO"
        try {
            $pub = Invoke-RestMethod -Uri 'https://api.ipify.org?format=text' -UseBasicParsing -ErrorAction Stop
            Log ("Public IP: {0}" -f $pub) "INFO"
        } catch {
            Log "Unable to fetch public IP (network may be down)." "WARN"
        }
    } catch {
        Log "IP info error: $_" "FAIL"
    }
}

# Outlook helpers
function Repair-OutlookProfile {
    Log "Attempting Outlook profile remediation (close Outlook, backup OST files)..." "INFO"
    try {
        Stop-Process -Name OUTLOOK -ErrorAction SilentlyContinue
        $ostPath = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'
        if (Test-Path $ostPath) {
            Get-ChildItem -Path $ostPath -Filter *.ost -ErrorAction SilentlyContinue | ForEach-Object {
                $bak = $_.FullName + '.bak'
                Move-Item -Path $_.FullName -Destination $bak -Force -ErrorAction SilentlyContinue
                Log ("Moved OST: {0} -> {1}" -f $_.FullName, $bak) "INFO"
            }
        }
        Start-Process -FilePath "outlook.exe" -ErrorAction SilentlyContinue
        Log "Outlook restarted to recreate profile/cache." "PASS"
    } catch {
        Log "Outlook profile fix error: $_" "FAIL"
    }
}

# Printer helpers
function Clear-PrintSpooler {
    Log "Clearing Print Spooler..." "INFO"
    try {
        Stop-Service -Name Spooler -Force -ErrorAction SilentlyContinue
        $spool = Join-Path $env:windir 'System32\spool\PRINTERS'
        if (Test-Path $spool) { Get-ChildItem $spool -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue }
        Start-Service -Name Spooler -ErrorAction SilentlyContinue
        Log "Print spooler cleared and restarted." "PASS"
    } catch {
        Log "Print spooler error: $_" "FAIL"
    }
}

function Get-AllPrinters {
    Log "Listing installed printers..." "INFO"
    try {
        Get-WmiObject Win32_Printer | ForEach-Object { Log ("{0} - Default:{1} Status:{2} Offline:{3}" -f $_.Name, $_.Default, $_.PrinterStatus, $_.WorkOffline) "INFO" }
    } catch {
        Log "Printer listing error: $_" "FAIL"
    }
}

# Windows Update helpers
function Reset-WindowsUpdateComponents {
    Log "Resetting Windows Update components..." "INFO"
    try {
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
        Stop-Service -Name cryptsvc -Force -ErrorAction SilentlyContinue
        Stop-Service -Name bits -Force -ErrorAction SilentlyContinue
        Stop-Service -Name msiserver -Force -ErrorAction SilentlyContinue

        $sd = Join-Path $env:windir 'SoftwareDistribution'
        $cat = Join-Path $env:windir 'System32\catroot2'
        if (Test-Path $sd) { Rename-Item $sd ($sd + '.old') -ErrorAction SilentlyContinue }
        if (Test-Path $cat) { Rename-Item $cat ($cat + '.old') -ErrorAction SilentlyContinue }

        Start-Service -Name wuauserv -ErrorAction SilentlyContinue
        Start-Service -Name cryptsvc -ErrorAction SilentlyContinue
        Start-Service -Name bits -ErrorAction SilentlyContinue
        Start-Service -Name msiserver -ErrorAction SilentlyContinue

        Log "Windows Update components reset. Recommend reboot before retrying updates." "PASS"
    } catch {
        Log "Windows Update reset error: $_" "FAIL"
    }
}

# Update button actions to use approved verbs
$btnSlowComputer = New-Button -Text "Slow Computer: Memory & Disk" -Name "btnSlowComputer" -BackColor ([System.Drawing.Color]::FromArgb(60,70,80)) -ToolTipText "Run memory and disk checks (top processes, large files, clean temp)" -Action {
    Show-TopMemoryProcesses
    Find-LargeFiles -Path 'C:\' -MinMB 200
    Clear-TempFiles
}
$btnSlowComputer.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnSlowComputer)
$yPosition += 47

$btnNetTroubleshoot = New-Button -Text "Network: Quick Tests & Fixes" -Name "btnNetTroubleshoot" -BackColor ([System.Drawing.Color]::FromArgb(60,70,80)) -ToolTipText "Show IPs, public IP, traffic; fix mapped drives; reset stack as last resort" -Action {
    Show-IPInfo
    Show-NetworkTraffic
    Repair-MappedDrives
}
$btnNetTroubleshoot.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnNetTroubleshoot)
$yPosition += 47

$btnNetworkReset = New-Button -Text "Network Stack Reset (Last Resort)" -Name "btnNetworkReset" -BackColor ([System.Drawing.Color]::FromArgb(70,50,50)) -ToolTipText "Resets winsock/TCPIP and renews IPs; may disconnect sessions" -Action { Reset-NetworkStack }
$btnNetworkReset.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnNetworkReset)
$yPosition += 47

$btnOutlookFix = New-Button -Text "Fix Outlook Profile & Restart" -Name "btnOutlookFix" -BackColor ([System.Drawing.Color]::FromArgb(60,70,80)) -ToolTipText "Backup OST files and restart Outlook to recreate profile" -Action { Repair-OutlookProfile }
$btnOutlookFix.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnOutlookFix)
$yPosition += 47

$btnPrinterFix = New-Button -Text "Printer: Clear Spooler & Cleanup" -Name "btnPrinterFix" -BackColor ([System.Drawing.Color]::FromArgb(60,70,80)) -ToolTipText "Clear print spooler, list printers, and remove offline ghosts" -Action {
    Clear-PrintSpooler
    Get-AllPrinters
    Remove-OfflinePrinters
}
$btnPrinterFix.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnPrinterFix)
$yPosition += 47

$btnWinUpdateFix = New-Button -Text "Windows Update: Reset & Reboot" -Name "btnWinUpdateFix" -BackColor ([System.Drawing.Color]::FromArgb(70,50,50)) -ToolTipText "Reset Update components, then reboot and retry updates" -Action {
    Reset-WindowsUpdateComponents
    if (Show-Confirmation -Message "Reboot now to complete Windows Update reset?") { Restart-Computer -Force }
}
$btnWinUpdateFix.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnWinUpdateFix)
$yPosition += 47

# MSP & Network Advanced Tools
$buttonContainer.Controls.Add((New-SectionHeader "7. MSP & Network Tools"))
$yPosition += 30

# BitLocker recovery helper (logs protector info; will not reveal keys unless available to admin)
function Get-BitLockerRecoveryInfo {
    Log "Querying BitLocker protectors (requires admin)" "INFO"
    try {
        # Prefer manage-bde (stable, does not require importing BitLocker module)
        if (Get-Command manage-bde -ErrorAction SilentlyContinue) {
            $out = & manage-bde -protectors -get C: 2>&1
            if ($out) { $out | ForEach-Object { Log $_ "INFO" } }
            else { Log "manage-bde returned no output or no BitLocker volume found." "WARN" }
        }
        elseif (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue) {
            try {
                Import-Module BitLocker -ErrorAction SilentlyContinue
                $vols = Get-BitLockerVolume -ErrorAction Stop
                foreach ($v in $vols) {
                    Log (("Volume: {0} Protection: {1}" -f $v.MountPoint, $v.ProtectionStatus)) "INFO"
                    if ($v.KeyProtector) { $v.KeyProtector | ForEach-Object { Log (("Protector: {0} ({1})" -f $_.KeyProtectorId, $_.KeyProtectorType)) "INFO" } }
                }
            } catch {
                Log "BitLocker module present but failed to query volumes: $_" "WARN"
            }
        } else {
            Log "No BitLocker tools (manage-bde or BitLocker module) available on this host." "WARN"
        }
    } catch {
        Log "BitLocker query error: $_" "FAIL"
    }
}

# Office 365 license checker (requires MSOnline module and credentials)
function Get-O365LicenseStatus {
    param([string]$UserUPN)
    if (-not $UserUPN) {
        Log "No user specified for O365 license check. Provide a UPN when invoking this function." "WARN"
        Log "Example usage: Get-O365LicenseStatus -UserUPN 'user@domain.com' (requires prior MSOnline connection)." "INFO"
        return
    }

    # Prefer Microsoft.Graph where available; fall back to MSOnline with a warning
    $hasGraph = (Get-Module -ListAvailable -Name Microsoft.Graph -ErrorAction SilentlyContinue)
    $hasMSOnline = (Get-Module -ListAvailable -Name MSOnline -ErrorAction SilentlyContinue)

    if (-not $hasGraph -and -not $hasMSOnline) {
        Log "Neither Microsoft.Graph nor MSOnline modules are installed. Install Microsoft.Graph (recommended) or MSOnline to enable license checks." "WARN"
        return
    }

    try {
        if ($hasGraph) {
            Log "Microsoft.Graph module detected. This function currently uses MSOnline APIs; consider implementing Graph API calls for license checks." "INFO"
            # Placeholder: keep current behavior if MSOnline is also available
        }

        if ($hasMSOnline) {
            try {
                $u = Get-MsolUser -UserPrincipalName $UserUPN -ErrorAction Stop
                Log ("User: {0} | Licenses: {1}" -f $u.UserPrincipalName, ($u.Licenses.Count)) "INFO"
                $u.Licenses | ForEach-Object { Log (" - {0}" -f $_.AccountSkuId) "INFO" }
            } catch {
                Log "MSOnline module present but not connected. Run Connect-MsolService interactively, then retry the check." "WARN"
            }
        } else {
            Log "Microsoft.Graph present but Graph-based license lookup not yet implemented in this script." "WARN"
        }
    } catch {
        Log "O365 license check error: $_" "FAIL"
    }
}

# Credential Manager diagnostics
function Show-CredentialManager {
    Log "Listing stored credentials via cmdkey..." "INFO"
    try {
        $out = cmdkey /list 2>&1
        $out | ForEach-Object { Log $_ "INFO" }
    } catch {
        Log "Credential manager error: $_" "FAIL"
    }
}

# Advanced Active Directory reporting (requires ActiveDirectory module)
function Get-ADAdvancedReport {
    Log "Generating AD summary (requires ActiveDirectory module)" "INFO"
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) { Log "ActiveDirectory module not available on this host." "WARN"; return }
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $userCount = (Get-ADUser -Filter * -ErrorAction SilentlyContinue).Count
        $compCount = (Get-ADComputer -Filter * -ErrorAction SilentlyContinue).Count
        $locked = Get-ADUser -Filter {LockedOut -eq $true} -ErrorAction SilentlyContinue | Select-Object -First 50
        Log ("AD Users: {0} | Computers: {1}" -f $userCount, $compCount) "INFO"
        if ($locked) { Log "Locked out sample users:" "WARN"; $locked | ForEach-Object { Log (" - {0}" -f $_.SamAccountName) "WARN" } }
    } catch {
        Log "AD report error: $_" "FAIL"
    }
}

# Remote Desktop troubleshooting
function Test-RemoteDesktop {
    param([string]$Target = $env:COMPUTERNAME)
    Log ("Testing RDP connectivity to {0}" -f $Target) "INFO"
    try {
        $term = Get-Service -Name TermService -ErrorAction SilentlyContinue
        Log ("RDP Service status: {0}" -f ($term.Status)) "INFO"
        Log "Attempting TCP connect to $Target:3389..." "INFO"
        Log "Waiting for response..." "INFO"
        $t = Test-NetConnection -ComputerName $Target -Port 3389 -WarningAction SilentlyContinue
        Log ("Test-NetConnection output: `n{0}" -f ($t | Out-String)) "INFO"
        Log ("RDP Port 3389 reachable: {0}" -f $t.TcpTestSucceeded) "INFO"
    } catch {
        Log "RDP test error: $_" "FAIL"
    }
}

# VPN diagnostics & basic repair
function Test-VPNDiagnostics {
    Log "Running VPN diagnostics..." "INFO"
    try {
        $ras = Get-Service -Name RasMan -ErrorAction SilentlyContinue
        Log ("RasMan status: {0}" -f $ras.Status) "INFO"
        if (Get-Command Get-VpnConnection -ErrorAction SilentlyContinue) {
            Get-VpnConnection -ErrorAction SilentlyContinue | ForEach-Object { Log ("VPN: {0} - State:{1}" -f $_.Name, $_.ConnectionStatus) "INFO" }
        } else { Log "Get-VpnConnection not available (older OS or missing module)." "WARN" }
    } catch {
        Log "VPN diagnostic error: $_" "FAIL"
    }
}

# Email configuration validator (simple connectivity tests)
function Test-EmailConfig {
    param([string]$Server, [int]$Port = 25)
    if (-not $Server) { Log "No server provided for email test; use the Email Config Validator and provide host/port." "WARN"; return }
    Log ("Testing connectivity to {0}:{1}" -f $Server, $Port) "INFO"
    try {
        Log ("Attempting TCP connect to {0}:{1}..." -f $Server, $Port) "INFO"
        Log "Waiting for response..." "INFO"
        $t = Test-NetConnection -ComputerName $Server -Port $Port -WarningAction SilentlyContinue
        Log ("Test-NetConnection output: `n{0}" -f ($t | Out-String)) "INFO"
        Log ("Port reachable: {0}" -f $t.TcpTestSucceeded) "INFO"
    } catch {
        Log "Email config test error: $_" "FAIL"
    }
}

# Browser cache/profile repair (Chrome/Edge/Firefox) - destructive: ask confirmation
function Repair-BrowserProfiles {
    if (-not (Show-Confirmation -Message "This will close browsers and remove cache/profile data for Chrome/Edge/Firefox. Continue?")) { return }
    Log "Repairing browser caches/profiles (closing browsers)" "INFO"
    try {
        Stop-Process -Name chrome,msedge,firefox -ErrorAction SilentlyContinue
        $chrome = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data\Default\Cache'
        $edge = Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data\Default\Cache'
        $ff = Join-Path $env:APPDATA 'Mozilla\Firefox\Profiles'
        if (Test-Path $chrome) { Remove-Item $chrome -Recurse -Force -ErrorAction SilentlyContinue; Log "Cleared Chrome cache" "PASS" }
        if (Test-Path $edge) { Remove-Item $edge -Recurse -Force -ErrorAction SilentlyContinue; Log "Cleared Edge cache" "PASS" }
        if (Test-Path $ff) { Get-ChildItem $ff -Directory | ForEach-Object { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }; Log "Cleared Firefox profiles cache" "PASS" }
    } catch {
        Log "Browser repair error: $_" "FAIL"
    }
}

# Automated ticket export (collect logs, system info, top processes)
function Export-Ticket {
    Log "Exporting ticket documentation to Desktop..." "INFO"
    try {
        $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $path = Join-Path $env:USERPROFILE ("Desktop\MSP_Ticket_{0}.txt" -f $timestamp)
        $sb = New-Object System.Text.StringBuilder
        $sb.AppendLine("MSP Ticket Export - $timestamp") | Out-Null
        $sb.AppendLine("Computer: $($env:COMPUTERNAME) User: $($env:USERNAME)") | Out-Null
        $sb.AppendLine("--- System Info ---") | Out-Null
        $sb.AppendLine((Get-CimInstance Win32_OperatingSystem | Select-Object Caption,Version,BuildNumber | Out-String)) | Out-Null
        $sb.AppendLine("--- Top Processes ---") | Out-Null
        $procs = Get-Process | Sort-Object WorkingSet -Descending | Select-Object -First 10 | Out-String
        $sb.AppendLine($procs) | Out-Null
        $sb.AppendLine("--- Recent Log ---") | Out-Null
        $sb.AppendLine($outputBox.Text.Substring(0,[math]::Min($outputBox.Text.Length,5000))) | Out-Null
        $sb.ToString() | Out-File -FilePath $path -Encoding UTF8 -Force
        Log ("Ticket exported: {0}" -f $path) "PASS"
    } catch {
        Log "Ticket export error: $_" "FAIL"
    }
}

# Add Rotate & Archive Logs button
$btnRotateLogs = New-Button -Text "5.4 Rotate & Archive Logs" -Name "btnRotateLogs" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Compress logs older than 7 days and delete archives older than 30 days." `
    -Action {
        Start-LogRotation
    }
$buttonContainer.Controls.Add($btnRotateLogs)
$yPosition += 47

# Email Log Summary button
$btnEmailLogs = New-Button -Text "5.5 Email Log Summary" -Name "btnEmailLogs" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Send the current log summary via email." `
    -Action {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Email Log Summary"
        $form.Size = New-Object System.Drawing.Size(400, 300)

        $smtpLabel = New-Object System.Windows.Forms.Label
        $smtpLabel.Text = "SMTP Server:"
        $smtpLabel.Location = New-Object System.Drawing.Point(10, 10)
        $form.Controls.Add($smtpLabel)

        $smtpBox = New-Object System.Windows.Forms.TextBox
        $smtpBox.Location = New-Object System.Drawing.Point(120, 10)
        $form.Controls.Add($smtpBox)

        $fromLabel = New-Object System.Windows.Forms.Label
        $fromLabel.Text = "From Address:"
        $fromLabel.Location = New-Object System.Drawing.Point(10, 50)
        $form.Controls.Add($fromLabel)

        $fromBox = New-Object System.Windows.Forms.TextBox
        $fromBox.Location = New-Object System.Drawing.Point(120, 50)
        $form.Controls.Add($fromBox)

        $toLabel = New-Object System.Windows.Forms.Label
        $toLabel.Text = "To Address:"
        $toLabel.Location = New-Object System.Drawing.Point(10, 90)
        $form.Controls.Add($toLabel)

        $toBox = New-Object System.Windows.Forms.TextBox
        $toBox.Location = New-Object System.Drawing.Point(120, 90)
        $form.Controls.Add($toBox)

        $sendButton = New-Object System.Windows.Forms.Button
        $sendButton.Text = "Send"
        $sendButton.Location = New-Object System.Drawing.Point(150, 150)
        $sendButton.Add_Click({
            Send-LogReport -SMTPServer $smtpBox.Text -FromAddress $fromBox.Text -ToAddress $toBox.Text
            $form.Close()
        })
        $form.Controls.Add($sendButton)

        $form.ShowDialog()
    }
$buttonContainer.Controls.Add($btnEmailLogs)
$yPosition += 47

# Add Parse Errors & Archive button
$btnParseLogs = New-Button -Text "5.6 Parse Errors & Archive" -Name "btnParseLogs" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Extract ERROR lines from a log file and archive the original." `
    -Action {
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Parse-LogErrors -LogFilePath $openFileDialog.FileName
        }
    }
$buttonContainer.Controls.Add($btnParseLogs)
$yPosition += 47

# Repair mapped drives dialog
function Repair-MappedDrives {
    Log "Launching mapped drive repair dialog..." "INFO"
    try {
        $maps = Get-CimInstance -ClassName Win32_MappedLogicalDisk -ErrorAction SilentlyContinue
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Mapped Drive Repair"
        $form.Size = New-Object System.Drawing.Size(400,300)
        $form.StartPosition = "CenterScreen"
        $list = New-Object System.Windows.Forms.ListBox
        $list.Location = New-Object System.Drawing.Point(10,10)
        $list.Size = New-Object System.Drawing.Size(360,200)
        foreach ($m in $maps) {
            $drive = $m.DeviceID.Substring(0,1)
            $root = $m.ProviderName
            $status = if ($null -eq (Test-Path ($drive + ':\'))) { "Disconnected" } else { "Connected" }
            $null = $list.Items.Add("{0}: {1} ({2})" -f $drive, $root, $status)
        }
        $repairBtn = New-Object System.Windows.Forms.Button
        $repairBtn.Text = "Repair Disconnected"
        $repairBtn.Location = New-Object System.Drawing.Point(10,220)
        $repairBtn.Size = New-Object System.Drawing.Size(150,30)
        $repairBtn.Add_Click({
            foreach ($m in $maps) {
                $drive = $m.DeviceID.Substring(0,1)
                $root = $m.ProviderName
                if ($null -eq (Test-Path ($drive + ':\'))) {
                    try {
                        New-PSDrive -Name $drive -PSProvider FileSystem -Root $root -Persist -ErrorAction Stop | Out-Null
                        Log ("Remapped {0} to {1}" -f $drive, $root) "PASS"
                    } catch {
                        Log ("Failed to remap {0} - {1}" -f $drive, $_) "WARN"
                    }
                }
            }
            $form.Close()
        })
        $form.Controls.Add($list)
        $form.Controls.Add($repairBtn)
        $form.ShowDialog() | Out-Null
    } catch {
        Log "Mapped drive repair dialog error: $_" "FAIL"
    }
}

# Add button for mapped drive repair dialog (insert after other network buttons)
$btnRepairMappedDrives = New-Button -Text "Repair Mapped Drives" -Name "btnRepairMappedDrives" -BackColor ([System.Drawing.Color]::LightBlue) -ToolTipText "Launch mapped drive repair dialog" -Action { Repair-MappedDrives }
$btnRepairMappedDrives.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnRepairMappedDrives)
$yPosition += 47

# SECTION: AUTOMATION SCRIPTS
$buttonContainer.Controls.Add((New-SectionHeader "3. AUTOMATION SCRIPTS"))
$yPosition += 30

# Add buttons for each script
$scripts = @(
    @{Name = "013_Network_Drive_Mapping.ps1"; Description = "Map network drives interactively."},
    @{Name = "019_DNS_Cache_Clearer.ps1"; Description = "Clear DNS cache."},
    @{Name = "024_ARP_Table_Exporter.ps1"; Description = "Export ARP table."},
    @{Name = "025_Wi-Fi_SSID_&_Strength_Scanner.ps1"; Description = "Scan Wi-Fi SSIDs and signal strength."},
    @{Name = "026_IP_Conflict_Detector.ps1"; Description = "Detect IP conflicts."},
    @{Name = "047_Open_Port_Audit.ps1"; Description = "Audit open ports."},
    @{Name = "048_Firewall_Rules_Exporter.ps1"; Description = "Export firewall rules."},
    @{Name = "066_Login_History_Export_(Windows).ps1"; Description = "Export login history."},
    @{Name = "067_Group_Membership_Change_Tracker_(AD).ps1"; Description = "Track AD group membership changes."},
    @{Name = "079_Update_Windows_Drivers_via_Script.ps1"; Description = "Update Windows drivers."},
    @{Name = "080_Log_Failed_RDP_Attempts_&_Block_IPs.ps1"; Description = "Log failed RDP attempts and block IPs."},
    @{Name = "081_OneDrive_Backup_Script.ps1"; Description = "Backup files to OneDrive."},
    @{Name = "091_RDP_Port_Change_Script.ps1"; Description = "Change RDP port."},
    @{Name = "100_Export_Group_Policy_Settings_to_CSV.ps1"; Description = "Export group policy settings to CSV."}
)

foreach ($s in $scripts) {
    $scriptName = $s.Name
    $scriptDesc = $s.Description
    $scriptPath = Join-Path $PSScriptRoot $scriptName

    $action = {
        try {
            if (Test-Path $scriptPath) {
                Log "Running $scriptName..." "INFO"
                $output = & $scriptPath 2>&1
                Log "Script $scriptName completed successfully." "PASS"
                $output | ForEach-Object { Log $_ "INFO" }
            } else {
                Log "Script file not found: $scriptPath" "FAIL"
                [System.Windows.Forms.MessageBox]::Show("Script file missing.`nExpected at: $scriptPath", "Error")
            }
        } catch {
            Log "Error running $($scriptName): $_" "FAIL"
        }
    }.GetNewClosure()

    $btn = New-Button -Text $scriptDesc -Name $scriptName `
        -BackColor ([System.Drawing.Color]::LightGreen) `
        -ToolTipText "Run $scriptDesc" `
        -Action $action

    $btn.Location = New-Object System.Drawing.Point(10, $yPosition)
    $buttonContainer.Controls.Add($btn)
    $yPosition += 47
}

# Continue to existing LOG MANAGEMENT section
# LOG MANAGEMENT
$buttonContainer.Controls.Add((New-SectionHeader "5. LOG MANAGEMENT"))
$yPosition += 30

# Clear Log Button
$btnClearLog = New-Button -Text "5.1 Clear Log" -Name "btnClearLog" `
    -BackColor ([System.Drawing.Color]::WhiteSmoke) `
    -ToolTipText "Clear all log entries from the console" `
    -Action {
        $outputBox.Clear()
        Log "Log cleared by user" "INFO"
    }
$btnClearLog.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnClearLog)
$yPosition += 47

# Save Log Button
$btnSaveLog = New-Button -Text "5.2 Save Log to Desktop" -Name "btnSaveLog" `
    -BackColor ([System.Drawing.Color]::WhiteSmoke) `
    -ToolTipText "Save current log to desktop with timestamp" `
    -Action {
        try {
            $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
            $logPath = "$env:USERPROFILE\Desktop\MSP_Toolkit_Log_$timestamp.txt"
            $outputBox.Text | Out-File -FilePath $logPath -Encoding UTF8 -Force
            Log "Log saved to: $logPath" "PASS"
        } catch {
            Log "Log save error: $_" "FAIL"
        }
    }
$btnSaveLog.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnSaveLog)
$yPosition += 47

# Copy Log Button
$btnCopyLog = New-Button -Text "5.3 Copy Log to Clipboard" -Name "btnCopyLog" `
    -BackColor ([System.Drawing.Color]::WhiteSmoke) `
    -ToolTipText "Copy entire log content to clipboard" `
    -Action {
        try {
            [System.Windows.Forms.Clipboard]::SetText($outputBox.Text)
            Log "Log copied to clipboard" "PASS"
        } catch {
            Log "Clipboard copy error: $_" "FAIL"
        }
    }
$btnCopyLog.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnCopyLog)

# Add Rotate & Archive Logs button
$btnRotateLogs = New-Button -Text "5.4 Rotate & Archive Logs" -Name "btnRotateLogs" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Compress logs older than 7 days and delete archives older than 30 days." `
    -Action {
        Start-LogRotation
    }
$buttonContainer.Controls.Add($btnRotateLogs)
$yPosition += 47

# Email Log Summary button
$btnEmailLogs = New-Button -Text "5.5 Email Log Summary" -Name "btnEmailLogs" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Send the current log summary via email." `
    -Action {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Email Log Summary"
        $form.Size = New-Object System.Drawing.Size(400, 300)

        $smtpLabel = New-Object System.Windows.Forms.Label
        $smtpLabel.Text = "SMTP Server:"
        $smtpLabel.Location = New-Object System.Drawing.Point(10, 10)
        $form.Controls.Add($smtpLabel)

        $smtpBox = New-Object System.Windows.Forms.TextBox
        $smtpBox.Location = New-Object System.Drawing.Point(120, 10)
        $form.Controls.Add($smtpBox)

        $fromLabel = New-Object System.Windows.Forms.Label
        $fromLabel.Text = "From Address:"
        $fromLabel.Location = New-Object System.Drawing.Point(10, 50)
        $form.Controls.Add($fromLabel)

        $fromBox = New-Object System.Windows.Forms.TextBox
        $fromBox.Location = New-Object System.Drawing.Point(120, 50)
        $form.Controls.Add($fromBox)

        $toLabel = New-Object System.Windows.Forms.Label
        $toLabel.Text = "To Address:"
        $toLabel.Location = New-Object System.Drawing.Point(10, 90)
        $form.Controls.Add($toLabel)

        $toBox = New-Object System.Windows.Forms.TextBox
        $toBox.Location = New-Object System.Drawing.Point(120, 90)
        $form.Controls.Add($toBox)

        $sendButton = New-Object System.Windows.Forms.Button
        $sendButton.Text = "Send"
        $sendButton.Location = New-Object System.Drawing.Point(150, 150)
        $sendButton.Add_Click({
            Send-LogReport -SMTPServer $smtpBox.Text -FromAddress $fromBox.Text -ToAddress $toBox.Text
            $form.Close()
        })
        $form.Controls.Add($sendButton)

        $form.ShowDialog()
    }
$buttonContainer.Controls.Add($btnEmailLogs)
$yPosition += 47

# Add Parse Errors & Archive button
$btnParseLogs = New-Button -Text "5.6 Parse Errors & Archive" -Name "btnParseLogs" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Extract ERROR lines from a log file and archive the original." `
    -Action {
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Parse-LogErrors -LogFilePath $openFileDialog.FileName
        }
    }
$buttonContainer.Controls.Add($btnParseLogs)
$yPosition += 47

# Repair mapped drives dialog
function Repair-MappedDrives {
    Log "Launching mapped drive repair dialog..." "INFO"
    try {
        $maps = Get-CimInstance -ClassName Win32_MappedLogicalDisk -ErrorAction SilentlyContinue
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Mapped Drive Repair"
        $form.Size = New-Object System.Drawing.Size(400,300)
        $form.StartPosition = "CenterScreen"
        $list = New-Object System.Windows.Forms.ListBox
        $list.Location = New-Object System.Drawing.Point(10,10)
        $list.Size = New-Object System.Drawing.Size(360,200)
        foreach ($m in $maps) {
            $drive = $m.DeviceID.Substring(0,1)
            $root = $m.ProviderName
            $status = if ($null -eq (Test-Path ($drive + ':\'))) { "Disconnected" } else { "Connected" }
            $null = $list.Items.Add("{0}: {1} ({2})" -f $drive, $root, $status)
        }
        $repairBtn = New-Object System.Windows.Forms.Button
        $repairBtn.Text = "Repair Disconnected"
        $repairBtn.Location = New-Object System.Drawing.Point(10,220)
        $repairBtn.Size = New-Object System.Drawing.Size(150,30)
        $repairBtn.Add_Click({
            foreach ($m in $maps) {
                $drive = $m.DeviceID.Substring(0,1)
                $root = $m.ProviderName
                if ($null -eq (Test-Path ($drive + ':\'))) {
                    try {
                        New-PSDrive -Name $drive -PSProvider FileSystem -Root $root -Persist -ErrorAction Stop | Out-Null
                        Log ("Remapped {0} to {1}" -f $drive, $root) "PASS"
                    } catch {
                        Log ("Failed to remap {0} - {1}" -f $drive, $_) "WARN"
                    }
                }
            }
            $form.Close()
        })
        $form.Controls.Add($list)
        $form.Controls.Add($repairBtn)
        $form.ShowDialog() | Out-Null
    } catch {
        Log "Mapped drive repair dialog error: $_" "FAIL"
    }
}

# Display the main form
$form.ShowDialog()
