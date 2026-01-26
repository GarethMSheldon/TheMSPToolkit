#Requires -RunAsAdministrator
<#
.SYNOPSIS
MSP Technician Toolkit - Professional Edition
.DESCRIPTION
Comprehensive troubleshooting GUI for MSP technicians with 60+ functions
covering system repair, security, networking, printers, storage, updates,
performance, domain management, diagnostics, and common helpdesk tickets.
.NOTES
Author: MSP Solutions Team
Version: 3.1.3
Date: January 26, 2026
Requires: PowerShell 3.0+, Administrator privileges
Compatible: Windows 7 SP1 through Windows 11 / Server 2022
#>

# Auto-elevation block
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$($myinvocation.mycommand.definition)`""
    Start-Process powershell -Verb RunAs -ArgumentList $arguments
    exit
}

# Enable DPI awareness for better scaling on high-resolution displays
try {
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
"@ -ErrorAction Stop

    [DPIAwareness]::SetDPIAware()
} catch {
    Write-Host "Warning: DPI awareness registration failed: $_"
}

# Load required assemblies
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
} catch {
    Write-Host "Warning: Failed loading System.Windows.Forms: $_"
}

try {
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
} catch {
    Write-Host "Warning: Failed loading System.Drawing: $_"
}

# Core logging function with color coding
function Write-ToolkitLog {
    param(
        [string]$Message,
        [string]$Status = "INFO"
    )
    try {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $color = "White"
        if($Status -eq "PASS"){$color="Lime"}
        if($Status -eq "FAIL"){$color="OrangeRed"}
        if($Status -eq "WARN"){$color="Yellow"}
        
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
}

# Alias for backwards compatibility
New-Alias -Name Log -Value Write-ToolkitLog -Force -ErrorAction SilentlyContinue

# Wrapper to execute actions with consistent start/done logging and error handling
function Invoke-ToolkitAction {
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    Log "Starting: $Name" "INFO"
    $bwReturned = $null
    try {
        Set-ProgressBarVisibility -Visible $true

        if ($null -eq $Action) {
            throw "No action provided"
        }

        $atype = $Action.GetType().FullName
        Log "Action type: $atype" "INFO"

        if ($Action -is [ScriptBlock]) {
            $bwReturned = & $Action
        } elseif ($Action -is [string]) {
            $bwReturned = Invoke-Expression $Action
        } elseif ($Action -is [System.Management.Automation.CommandInfo]) {
            $bwReturned = & $Action.Name
        } else {
            # Try invocation safely
            try { $bwReturned = & $Action } catch { throw "Action is not invokable: $atype" }
        }

        # If the action started a BackgroundWorker, let its completion handler clear progress
        if ($bwReturned -is [System.ComponentModel.BackgroundWorker]) {
            Log "Background task launched for $Name" "INFO"
            return
        }

        Log "Done: $Name" "INFO"
    } catch {
        Log "$Name error: $_" "FAIL"
    } finally {
        if (-not ($bwReturned -is [System.ComponentModel.BackgroundWorker])) {
            Set-ProgressBarVisibility -Visible $false
        }
    }
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
    try {
        if ($script:progressBar -and $script:progressBar -is [System.Windows.Forms.ToolStripProgressBar]) {
            $script:progressBar.Visible = $Visible
        }
    } catch {}
    try { $statusBar.Refresh() } catch {}
    try { [System.Windows.Forms.Application]::DoEvents() } catch {}
}

# Track background workers and external processes for safe shutdown
$script:BackgroundWorkers = @()
$script:ExternalProcesses = @()

# Lightweight locks for thread-safe access to shared trackers
$script:BackgroundWorkersLock = New-Object System.Object
$script:ExternalProcessesLock = New-Object System.Object

# Helper functions to manipulate shared trackers under a lock
function Add-BackgroundWorker { param($bw)
    [System.Threading.Monitor]::Enter($script:BackgroundWorkersLock)
    try { $script:BackgroundWorkers += $bw } finally { [System.Threading.Monitor]::Exit($script:BackgroundWorkersLock) }
}

function Remove-BackgroundWorker { param($bw)
    [System.Threading.Monitor]::Enter($script:BackgroundWorkersLock)
    try {
        # Unregister any registered event subscribers for this worker
        try {
            if ($script:BackgroundWorkerEvents.ContainsKey($bw)) {
                $subs = $script:BackgroundWorkerEvents[$bw]
                foreach ($sid in $subs.SourceIds) {
                    try { Unregister-Event -SourceIdentifier $sid -ErrorAction SilentlyContinue } catch {}
                    try { Get-EventSubscriber -SourceIdentifier $sid | Unregister-Event -ErrorAction SilentlyContinue } catch {}
                }
                try { $script:BackgroundWorkerEvents.Remove($bw) } catch {}
            }
        } catch {}

        $script:BackgroundWorkers = $script:BackgroundWorkers | Where-Object { $_ -ne $bw }
    } finally { [System.Threading.Monitor]::Exit($script:BackgroundWorkersLock) }
}

function Add-ExternalProcess { param($proc)
    [System.Threading.Monitor]::Enter($script:ExternalProcessesLock)
    try { $script:ExternalProcesses += $proc } finally { [System.Threading.Monitor]::Exit($script:ExternalProcessesLock) }
}

function Remove-ExternalProcess { param($proc)
    [System.Threading.Monitor]::Enter($script:ExternalProcessesLock)
    try { $script:ExternalProcesses = $script:ExternalProcesses | Where-Object { $_ -ne $proc } } finally { [System.Threading.Monitor]::Exit($script:ExternalProcessesLock) }
}

# Event/action registries for BackgroundWorker fallback handling
$script:BackgroundWorkerEvents = @{}
$script:BackgroundWorkerActions = @{}

# Start a background task using BackgroundWorker and keep track of it
function Start-BackgroundTask {
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    try {
        if (-not $Action) { throw 'No action provided to Start-BackgroundTask' }

        $actionBlock = $Action -as [scriptblock]
        if (-not $actionBlock) { throw 'Start-BackgroundTask requires -Action as a scriptblock' }

        if ($null -ne $form -and $form -is [System.Windows.Forms.Form] -and $form.InvokeRequired) {
            try { $form.Invoke([Action]{ Set-ProgressBarVisibility -Visible $true }) } catch { Set-ProgressBarVisibility -Visible $true }
        } else { Set-ProgressBarVisibility -Visible $true }

        $bw = New-Object System.ComponentModel.BackgroundWorker
        $bw.WorkerReportsProgress = $false
        $bw.WorkerSupportsCancellation = $true

        $attached = $false
        try {
            $bw.DoWork.Add({ param($bwSender,$bwArgs)
                try {
                    & $actionBlock
                } catch {
                    $bwArgs.Result = $_
                    throw $_
                }
            })

            $bw.RunWorkerCompleted.Add({ param($bwSender,$bwArgs)
                try {
                    if ($null -ne $form -and $form -is [System.Windows.Forms.Form] -and $form.InvokeRequired) {
                        try { $form.Invoke([Action]{ Set-ProgressBarVisibility -Visible $false }) } catch { Set-ProgressBarVisibility -Visible $false }
                    } else { Set-ProgressBarVisibility -Visible $false }
                    if ($bwArgs.Error) {
                        $msg = if ($bwArgs.Error.Exception) { $bwArgs.Error.Exception.Message } else { $bwArgs.Error.ToString() }
                        Log (("{0} error: {1}" -f $Name, $msg)) "FAIL"
                    } else {
                        Log ("Background task completed: {0}" -f $Name) "INFO"
                    }
                } catch { Log (("Background completion handler error for {0}: {1}" -f $Name, $_)) "WARN" }
            })

            $attached = $true
        } catch {
            # Fallback: Register-ObjectEvent subscribers when direct .Add fails
            try {
                $bwId = [guid]::NewGuid().ToString()
                $script:BackgroundWorkerActions[$bwId] = $actionBlock

                $sidDo = "MSP_BW_DoWork_$bwId"
                $sidComp = "MSP_BW_Completed_$bwId"

                Register-ObjectEvent -InputObject $bw -EventName DoWork -SourceIdentifier $sidDo -MessageData $bwId -Action {
                    try {
                        $id = $Event.MessageData
                        & $script:BackgroundWorkerActions[$id]
                    } catch {
                        try { $Event.SourceEventArgs.Result = $_ } catch {}
                        throw $_
                    }
                } | Out-Null

                Register-ObjectEvent -InputObject $bw -EventName RunWorkerCompleted -SourceIdentifier $sidComp -MessageData $bwId -Action {
                    try {
                        $id = $Event.MessageData
                        if ($Event.SourceEventArgs.Error) {
                            $msg = $Event.SourceEventArgs.Error.Message
                            Log (("{0} error: {1}" -f $Event.SourceEventArgs.Result, $msg)) "FAIL"
                        } else {
                            Log (("Background task completed: {0}" -f $id)) "INFO"
                        }
                    } catch { Log (("Background completion handler error for fallback {0}: {1}" -f $id, $_)) "WARN" }
                } | Out-Null

                # Record the source identifiers for cleanup
                $script:BackgroundWorkerEvents[$bw] = @{ SourceIds = @($sidDo,$sidComp); Id = $bwId }
                $attached = $true
            } catch {
                throw ("Failed to attach event handlers (direct and fallback): {0}" -f $_)
            }
        }

        try {
            # Start worker and register only after successful start
            $started = $bw.RunWorkerAsync()
            Add-BackgroundWorker $bw
            return $bw
        } catch {
            throw ("Failed to start BackgroundWorker: {0}" -f $_)
        }
    } catch {
        try { Log (("Failed to start background task {0}: {1}" -f $Name, $_)) "FAIL" } catch { Write-Host ("Failed to start background task {0}: {1}" -f $Name, $_) }
        return $null
    }
}

# Safe external command invoker: returns hashtable with ExitCode and Output (array of lines)
function Invoke-ExternalCommand {
    param(
        [Parameter(Mandatory=$true)][string]$Command,
        [string[]]$Args = @()
    )

    try {
        # Resolve command using Get-Command if necessary
        $cmdObj = Get-Command $Command -ErrorAction SilentlyContinue
        if ($cmdObj) { $CommandPath = $cmdObj.Source } else { $CommandPath = $Command }

        if (-not (Test-Path $CommandPath) -and -not $cmdObj) {
            Log "External command not found: $CommandPath" "FAIL"
            return @{ ExitCode = 127; Output = @(); Error = "NotFound" }
        }

        $argList = $Args | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } } | Join-String ' '

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $CommandPath
        $psi.Arguments = $argList
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        Add-ExternalProcess $proc
        try {
            $started = $proc.Start()
            $stdOut = $proc.StandardOutput.ReadToEnd()
            $stdErr = $proc.StandardError.ReadToEnd()
            $proc.WaitForExit()
            $exit = $proc.ExitCode

            $outLines = @()
            if ($stdOut) { $outLines += ($stdOut -split "`r?`n") }
            if ($stdErr) { $outLines += ($stdErr -split "`r?`n") }

            return @{ ExitCode = $exit; Output = $outLines }
        } finally {
            try { Remove-ExternalProcess $proc } catch {}
            try { $proc.Dispose() } catch {}
        }
    } catch {
        Log ("Invoke-ExternalCommand failed for {0}: {1}" -f $Command, $_) "FAIL"
        return @{ ExitCode = 1; Output = @(); Error = $_ }
    }
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "MSP Technician Toolkit - Professional Edition"
$form.Size = New-Object System.Drawing.Size(1300, 950)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.TopMost = $false
$form.ShowIcon = $false
 # Improve high-DPI rendering
 $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi

# Add resize event handler
$form.add_Resize({
    $sidebarPanel.Height = $form.ClientSize.Height - $statusBar.Height - 20
    $consolePanel.Height = $form.ClientSize.Height - $statusBar.Height - 20
    $outputBox.Width = $consolePanel.ClientSize.Width - 20
    $outputBox.Height = $consolePanel.ClientSize.Height - 20
    $buttonContainer.Height = [Math]::Max(5500, $sidebarPanel.ClientSize.Height + $sidebarPanel.VerticalScroll.Value + 100)
})

# Create status bar
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
    $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $domainStatus.Text = "Domain: $($computerSystem.Domain)"
} catch {
    $domainStatus.Text = "Domain: Error detecting domain"
}

$statusBar.Items.AddRange(@($adminStatus, $computerName, $currentUserName, $domainStatus))

# Progress bar
$progressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$progressBar.MarqueeAnimationSpeed = 30
$progressBar.Width = 200
$progressBar.Visible = $false
$null = $statusBar.Items.Add($progressBar)
try { $script:progressBar = $progressBar } catch {}

# Create sidebar panel
$sidebarPanel = New-Object System.Windows.Forms.Panel
$sidebarPanel.Width = 620
$sidebarPanel.Dock = [System.Windows.Forms.DockStyle]::Left
$sidebarPanel.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$sidebarPanel.AutoScroll = $true
$sidebarPanel.BorderStyle = "None"
try { $sidebarPanel.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi } catch {}

# Create scrollable container
$buttonContainer = New-Object System.Windows.Forms.Panel
$buttonContainer.Width = 600
$buttonContainer.Height = 600
$buttonContainer.Location = New-Object System.Drawing.Point(0, 0)
$buttonContainer.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$null = $sidebarPanel.Controls.Add($buttonContainer)
try { $buttonContainer.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi } catch {}

# Create console panel
$consolePanel = New-Object System.Windows.Forms.Panel
$consolePanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$consolePanel.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$consolePanel.BorderStyle = "Fixed3D"
try { $consolePanel.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi } catch {}

# Create output box
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$outputBox.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$outputBox.ForeColor = [System.Drawing.Color]::Gainsboro
$outputBox.ReadOnly = $true
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$outputBox.WordWrap = $false
$outputBox.ScrollBars = "Vertical"
$null = $consolePanel.Controls.Add($outputBox)

# Add panels to form
$null = $form.Controls.Add($statusBar)
$null = $form.Controls.Add($consolePanel)
$null = $form.Controls.Add($sidebarPanel)

# Button creation function
function New-ToolkitButton {
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
    $softBack = [System.Drawing.Color]::FromArgb(60,70,80)
    $hoverBack = [System.Drawing.Color]::FromArgb(75,85,95)
    $downBack  = [System.Drawing.Color]::FromArgb(50,60,70)
    $borderCol = [System.Drawing.Color]::FromArgb(95,100,105)
    $button.BackColor = $softBack
    $button.ForeColor = [System.Drawing.Color]::Gainsboro
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $button.UseCompatibleTextRendering = $true
    $button.FlatAppearance.BorderColor = $borderCol
    $button.FlatAppearance.MouseOverBackColor = $hoverBack
    $button.FlatAppearance.MouseDownBackColor = $downBack
    # Store text and action in Tag to avoid closure/event handler binding issues
    $button.Tag = @{ Name = $Text; Action = $Action }
    $button.Add_Click({ param($btnSender,$btnArgs)
        $info = $btnSender.Tag
        $name = if ($info -and $info.Name) { $info.Name } else { $btnSender.Name }
        $act  = if ($info -and $info.Action) { $info.Action } else { $null }
        Invoke-ToolkitAction -Name $name -Action $act
    })
    $button.Anchor = "Top, Left"
    
    $tooltip = New-Object System.Windows.Forms.ToolTip
    $tooltip.AutoPopDelay = 8000
    $tooltip.InitialDelay = 500
    $tooltip.ReshowDelay = 100
    $tooltip.SetToolTip($button, $ToolTipText)
    
    return $button
}

# Section header function
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
    $label.Anchor = "Top, Left"
    return $label
}

# Initialize position
$yPosition = 10

# SECTION 1: SYSTEM REPAIR
$buttonContainer.Controls.Add((New-SectionHeader "1. SYSTEM REPAIR AND HEALTH"))
$yPosition += 30

# 1.1 SFC Scan
$btnSFC = New-ToolkitButton -Text "1.1 Run SFC Scan (System File Checker)" -Name "btnSFC" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Scans and repairs corrupted Windows system files. Takes 10-15 minutes." `
    -Action {
        if (Show-Confirmation -Message "This operation takes 10-15 minutes and requires system file access. Continue?") {
            Start-BackgroundTask -Name 'SFC Scan' -Action {
                try {
                    Log "Starting SFC scan..." "INFO"
                    $sfcCmd = Get-Command sfc -ErrorAction SilentlyContinue
                    if (-not $sfcCmd) { Log "sfc.exe not available on this system" "FAIL"; return }

                    $result = Invoke-ExternalCommand -Command $sfcCmd.Source -Args '/scannow'
                    if ($result.Output) { $result.Output | ForEach-Object { Log $_ "INFO" } }

                    if ($result.Output -match "Windows Resource Protection did not find any integrity violations") {
                        Log "SFC completed successfully. No integrity violations found." "PASS"
                    }
                    elseif ($result.Output -match "Windows Resource Protection found corrupt files and successfully repaired them") {
                        Log "SFC found and repaired corrupt files. A reboot may be required." "WARN"
                    }
                    elseif ($result.Output -match "Windows Resource Protection found corrupt files but was unable to fix some of them") {
                        Log "SFC found corrupt files but could not repair all of them. Run DISM repair next." "FAIL"
                    }
                    else {
                        Log "SFC scan completed with unexpected results. Review output above." "WARN"
                    }
                } catch {
                    Log "SFC scan error: $_" "FAIL"
                }
            }
        }
    }
$btnSFC.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnSFC)
$yPosition += 47

# 1.2 Reboot Check
$btnRebootCheck = New-ToolkitButton -Text "1.2 Check for Pending Reboot" -Name "btnRebootCheck" `
    -BackColor ([System.Drawing.Color]::White) `
    -ToolTipText "Checks multiple registry locations for pending system reboots." `
    -Action {
        try {
            $pendingReboot = $false
            Log "Checking for pending reboot flags..." "INFO"
            
            $rebootKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
            if (Test-Path $rebootKey) {
                $items = Get-ChildItem $rebootKey -ErrorAction SilentlyContinue
                if ($items) {
                    Log "[REGISTRY] RebootRequired key exists with $($items.Count) items" "WARN"
                    $pendingReboot = $true
                }
            }
            
            $rebootPendingKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
            if (Test-Path $rebootPendingKey) {
                $value = Get-ItemProperty $rebootPendingKey -Name "RebootPending" -ErrorAction SilentlyContinue
                if ($value -and $value.RebootPending) {
                    Log "[REGISTRY] RebootPending flag detected" "WARN"
                    $pendingReboot = $true
                }
            }
            
            $sessionMgrKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager"
            if (Test-Path $sessionMgrKey) {
                $value = Get-ItemProperty $sessionMgrKey -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue
                if ($value -and $value.PendingFileRenameOperations) {
                    Log "[REGISTRY] Pending file rename operations detected" "WARN"
                    $pendingReboot = $true
                }
            }
            
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

# 1.3 DISM Repair
$btnDISM = New-ToolkitButton -Text "1.3 Run DISM Repair" -Name "btnDISM" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Repairs Windows component store corruption. Use before SFC if it fails." `
    -Action {
        if (Show-Confirmation -Message "DISM repair may take 15-20 minutes. Continue?") {
            Start-BackgroundTask -Name 'DISM Repair' -Action {
                try {
                    Log "Starting DISM repair..." "INFO"
                    $dismCmd = Get-Command dism.exe -ErrorAction SilentlyContinue
                    if (-not $dismCmd) { Log "DISM not available on this system" "FAIL"; return }

                    $startTime = Get-Date
                    $res = Invoke-ExternalCommand -Command $dismCmd.Source -Args @('/Online','/Cleanup-Image','/RestoreHealth')
                    $endTime = Get-Date
                    $duration = $endTime - $startTime

                    if ($res.Output) { $res.Output | ForEach-Object { Log $_ "INFO" } }

                    if ($res.ExitCode -eq 0) {
                        Log "DISM repair completed successfully in $($duration.Minutes) minutes, $($duration.Seconds) seconds" "PASS"
                    } else {
                        Log "DISM repair failed with exit code $($res.ExitCode)" "FAIL"
                        Log "Try running as Administrator or check Windows Update service status" "WARN"
                    }
                } catch {
                    Log "DISM error: $_" "FAIL"
                    Log "Ensure DISM is available on this Windows version. May not work on Windows 7 without updates." "WARN"
                }
            }
        }
    }
$btnDISM.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnDISM)
$yPosition += 52

# SECTION 2: USER PROFILE
$buttonContainer.Controls.Add((New-SectionHeader "2. USER PROFILE AND M365"))
$yPosition += 30

$btnTeamsCache = New-ToolkitButton -Text "2.1 Clear Teams Cache" -Name "btnTeamsCache" `
    -BackColor ([System.Drawing.Color]::White) `
    -ToolTipText "Fixes Teams freezing, login issues, and sync problems" `
    -Action {
        Start-BackgroundTask -Name 'Clear Teams Cache' -Action {
            try {
                Log "Clearing Teams cache..." "INFO"
                $teamsProcesses = Get-Process -Name "Teams","ms-teams" -ErrorAction SilentlyContinue
                if ($teamsProcesses) {
                    Log "Stopping Teams process..." "INFO"
                    $teamsProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 2
                }

                $teamsCachePath = "$env:APPDATA\Microsoft\Teams"
                if (Test-Path $teamsCachePath) {
                    Get-ChildItem $teamsCachePath -Recurse -Force -ErrorAction SilentlyContinue |
                        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                    Log "Cleared classic Teams cache" "PASS"
                }

                $newTeamsCachePath = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache"
                if (Test-Path $newTeamsCachePath) {
                    Get-ChildItem $newTeamsCachePath -Recurse -Force -ErrorAction SilentlyContinue |
                        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                    Log "Cleared new Teams cache" "PASS"
                }

                Log "Teams cache cleared. Restart Teams to apply changes." "PASS"
            } catch {
                Log "Teams cache error: $_" "FAIL"
            }
        }
    }
$btnTeamsCache.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnTeamsCache)
$yPosition += 47

# 2.2 Outlook Profile Fix
$btnOutlookFix = New-ToolkitButton -Text "2.2 Fix Outlook Profile" -Name "btnOutlookFix" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Fixes 'Enter Password' prompts and profile corruption" `
    -Action {
        if (Show-Confirmation -Message "This will reset Outlook profiles. Continue?") {
            Start-BackgroundTask -Name 'Fix Outlook Profile' -Action {
                try {
                    Log "Fixing Outlook profiles..." "INFO"

                    $outlookProcess = Get-Process -Name "outlook" -ErrorAction SilentlyContinue
                    if ($outlookProcess) {
                        $outlookProcess | Stop-Process -Force
                        Start-Sleep -Seconds 2
                    }

                    $officeVersions = @("16.0", "15.0", "14.0")
                    $profilesDeleted = 0

                    foreach ($version in $officeVersions) {
                        $profilePath = "HKCU:\Software\Microsoft\Office\$version\Outlook\Profiles"
                        if (Test-Path $profilePath) {
                            Remove-Item -Path $profilePath -Recurse -Force -ErrorAction SilentlyContinue
                            $profilesDeleted++
                        }
                    }

                    if ($profilesDeleted -gt 0) {
                        Log "Deleted Outlook profiles. Restart Outlook to create new profile." "PASS"
                    } else {
                        Log "No Outlook profiles found" "WARN"
                    }
                } catch {
                    Log "Outlook profile error: $_" "FAIL"
                }
            }
        }
    }
$btnOutlookFix.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnOutlookFix)
$yPosition += 47

# 2.3 OneDrive Sync Fix
$btnOneDrive = New-ToolkitButton -Text "2.3 Fix OneDrive Sync" -Name "btnOneDrive" `
    -BackColor ([System.Drawing.Color]::White) `
    -ToolTipText "Resets OneDrive sync engine for stuck files" `
    -Action {
        Start-BackgroundTask -Name 'Fix OneDrive Sync' -Action {
            try {
                Log "Fixing OneDrive sync..." "INFO"

                $oneDrivePaths = @(
                    "$env:LocalAppData\Microsoft\OneDrive\onedrive.exe",
                    "$env:ProgramFiles\Microsoft OneDrive\onedrive.exe",
                    "${env:ProgramFiles(x86)}\Microsoft OneDrive\onedrive.exe"
                )

                $oneDrivePath = $oneDrivePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

                if (-not $oneDrivePath) {
                    Log "OneDrive executable not found" "FAIL"
                    return
                }

                $oneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
                if ($oneDriveProcess) {
                    $oneDriveProcess | Stop-Process -Force
                    Start-Sleep -Seconds 2
                }

                Start-Process -FilePath $oneDrivePath -ArgumentList "/reset" -Wait -NoNewWindow
                Start-Sleep -Seconds 8
                Start-Process -FilePath $oneDrivePath -NoNewWindow

                Log "OneDrive sync reset completed" "PASS"
            } catch {
                Log "OneDrive error: $_" "FAIL"
            }
        }
    }
$btnOneDrive.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnOneDrive)
$yPosition += 47

# 2.4 Mapped Drive Diagnostics
$btnMappedDrives = New-ToolkitButton -Text "2.4 Test Mapped Drives" -Name "btnMappedDrives" `
    -BackColor ([System.Drawing.Color]::White) `
    -ToolTipText "Tests connectivity to all mapped network drives" `
    -Action { Test-MappedDrives }
$btnMappedDrives.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnMappedDrives)
$yPosition += 52

# SECTION 3: ADVANCED NETWORKING
$buttonContainer.Controls.Add((New-SectionHeader "3. ADVANCED NETWORKING"))
$yPosition += 30

# 3.1 Network Traffic Monitor
$btnNetTraffic = New-ToolkitButton -Text "3.1 Network Traffic Monitor" -Name "btnNetTraffic" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Identifies bandwidth-consuming processes" `
    -Action { Show-NetworkTraffic }
$btnNetTraffic.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnNetTraffic)
$yPosition += 47

# 3.2 WiFi Password Recovery
$btnWiFiPass = New-ToolkitButton -Text "3.2 WiFi Password Recovery" -Name "btnWiFiPass" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Extracts saved wireless network credentials" `
    -Action { Get-WiFiPasswords }
$btnWiFiPass.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnWiFiPass)
$yPosition += 47

# 3.3 Network Stack Reset
$btnNetReset = New-ToolkitButton -Text "3.3 Network Stack Reset" -Name "btnNetReset" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Nuclear option for persistent connectivity issues" `
    -Action { Reset-NetworkStack }
$btnNetReset.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnNetReset)
$yPosition += 47

# 3.4 IP Address Display
$btnIPDisplay = New-ToolkitButton -Text "3.4 Show IP Addresses" -Name "btnIPDisplay" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Shows both local and public IP addresses" `
    -Action { Show-IPAddresses }
$btnIPDisplay.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnIPDisplay)
$yPosition += 52

# SECTION 4: PRINTER TROUBLESHOOTING
$buttonContainer.Controls.Add((New-SectionHeader "4. PRINTER TROUBLESHOOTING"))
$yPosition += 30

# 4.1 Clear Print Spooler
$btnPrintSpooler = New-ToolkitButton -Text "4.1 Clear Print Spooler" -Name "btnPrintSpooler" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Clears stuck print jobs and restarts spooler service" `
    -Action { Clear-PrintSpooler }
$btnPrintSpooler.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnPrintSpooler)
$yPosition += 47

# 4.2 Printer Inventory
$btnPrinterInventory = New-ToolkitButton -Text "4.2 Printer Inventory" -Name "btnPrinterInventory" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Lists all installed printers with status" `
    -Action { Show-PrinterInventory }
$btnPrinterInventory.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnPrinterInventory)
$yPosition += 47

# 4.3 Remove Ghost Printers
$btnGhostPrinters = New-ToolkitButton -Text "4.3 Remove Ghost Printers" -Name "btnGhostPrinters" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Automatically removes offline/error printers" `
    -Action { Remove-GhostPrinters }
$btnGhostPrinters.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnGhostPrinters)
$yPosition += 52

# SECTION 5: DISK AND STORAGE
$buttonContainer.Controls.Add((New-SectionHeader "5. DISK AND STORAGE HEALTH"))
$yPosition += 30

# 5.1 CHKDSK Scan
$btnCHKDSK = New-ToolkitButton -Text "5.1 Run CHKDSK Scan" -Name "btnCHKDSK" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Non-destructive disk integrity check" `
    -Action { Start-CHKDSKScan }
$btnCHKDSK.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnCHKDSK)
$yPosition += 47

# 5.2 Large File Finder
$btnLargeFiles = New-ToolkitButton -Text "5.2 Find Large Files (500MB+)" -Name "btnLargeFiles" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Locates files consuming significant disk space" `
    -Action { Find-LargeFiles -Path 'C:\' -MinMB 500 }
$btnLargeFiles.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnLargeFiles)
$yPosition += 47

# 5.3 SMART Status
$btnSMART = New-ToolkitButton -Text "5.3 Check SMART Status" -Name "btnSMART" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Early warning system for failing drives" `
    -Action { Get-SMARTStatus }
$btnSMART.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnSMART)
$yPosition += 47

# 5.4 Clean Temp Files
$btnCleanTemp = New-ToolkitButton -Text "5.4 Clean Temp Files" -Name "btnCleanTemp" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Automated cleanup of Windows and user temp folders" `
    -Action { Clear-TempFiles }
$btnCleanTemp.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnCleanTemp)
$yPosition += 47

# 5.5 Disk Health Dashboard
$btnDiskHealth = New-ToolkitButton -Text "5.5 Disk Health Dashboard" -Name "btnDiskHealth" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Physical disk health monitoring" `
    -Action { Show-DiskHealth }
$btnDiskHealth.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnDiskHealth)
$yPosition += 52

# SECTION 6: WINDOWS UPDATE
$buttonContainer.Controls.Add((New-SectionHeader "6. WINDOWS UPDATE FIXES"))
$yPosition += 30

# 6.1 Reset Update Components
$btnResetUpdate = New-ToolkitButton -Text "6.1 Reset Update Components" -Name "btnResetUpdate" `
    -BackColor ([System.Drawing.Color]::LightCoral) `
    -ToolTipText "Fixes 80% of Windows Update failures" `
    -Action { Reset-WindowsUpdate }
$btnResetUpdate.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnResetUpdate)
$yPosition += 47

# 6.2 Update History
$btnUpdateHistory = New-ToolkitButton -Text "6.2 View Update History" -Name "btnUpdateHistory" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Shows last 5 installed updates with dates" `
    -Action { Show-UpdateHistory }
$btnUpdateHistory.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnUpdateHistory)
$yPosition += 52

# SECTION 7: PERFORMANCE & DIAGNOSTICS
$buttonContainer.Controls.Add((New-SectionHeader "7. PERFORMANCE AND DIAGNOSTICS"))
$yPosition += 30

# 7.1 Memory Usage Report
$btnMemoryReport = New-ToolkitButton -Text "7.1 Memory Usage Report" -Name "btnMemoryReport" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Identifies RAM-hogging processes" `
    -Action { Show-TopMemoryProcesses }
$btnMemoryReport.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnMemoryReport)
$yPosition += 47

# 7.2 Blue Screen Detection
$btnBSOD = New-ToolkitButton -Text "7.2 Detect Blue Screens" -Name "btnBSOD" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Scans event logs for recent BSOD errors" `
    -Action { Get-BSODEvents }
$btnBSOD.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnBSOD)
$yPosition += 47

# 7.3 Battery Health Report
$btnBattery = New-ToolkitButton -Text "7.3 Battery Health Report" -Name "btnBattery" `
    -BackColor ([System.Drawing.Color]::LightBlue) `
    -ToolTipText "Generates detailed battery degradation analysis (laptops)" `
    -Action { New-BatteryReport }
$btnBattery.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnBattery)
$yPosition += 52

# SECTION 8: DOMAIN & AD
$buttonContainer.Controls.Add((New-SectionHeader "8. DOMAIN AND ACTIVE DIRECTORY"))
$yPosition += 30

# 8.1 Domain Connection Test
$btnDomainTest = New-ToolkitButton -Text "8.1 Test Domain Connection" -Name "btnDomainTest" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Verifies secure channel to domain controller" `
    -Action { Test-DomainConnection }
$btnDomainTest.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnDomainTest)
$yPosition += 47

# 8.2 Group Policy Update
$btnGPUpdate = New-ToolkitButton -Text "8.2 Force Group Policy Update" -Name "btnGPUpdate" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Forces immediate policy refresh" `
    -Action { Update-GroupPolicy }
$btnGPUpdate.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnGPUpdate)
$yPosition += 47

# 8.3 Last Logon Tracker
$btnLastLogon = New-ToolkitButton -Text "8.3 Show Last Logon" -Name "btnLastLogon" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Shows last user login time and username" `
    -Action { Show-LastLogon }
$btnLastLogon.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnLastLogon)
$yPosition += 52

# SECTION 9: SYSTEM INFORMATION
$buttonContainer.Controls.Add((New-SectionHeader "9. SYSTEM INFORMATION"))
$yPosition += 30

# 9.1 System Dashboard
$btnSystemDash = New-ToolkitButton -Text "9.1 System Dashboard" -Name "btnSystemDash" `
    -BackColor ([System.Drawing.Color]::LightGreen) `
    -ToolTipText "Complete system overview - name, domain, OS, hardware, uptime" `
    -Action {
        try {
                Set-ProgressBarVisibility -Visible $true
                Log "Generating system dashboard..." "INFO"
                
                $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
                $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
                $cpu = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
                $physicalMemory = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue
                $bios = Get-CimInstance Win32_BIOS -ErrorAction Stop
            
            $totalRAM = 0
            if ($physicalMemory) {
                foreach ($mem in $physicalMemory) {
                    $totalRAM += $mem.Capacity
                }
                $totalRAM = [math]::Round(($totalRAM / 1GB), 2)
            } else {
                $totalRAM = [math]::Round(($computerSystem.TotalPhysicalMemory / 1GB), 2)
            }
            
            $lastBoot = $null
            $installDate = $null
            try {
                $lastBoot = [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
            } catch {
                try { $lastBoot = Get-Date ($os.LastBootUpTime) } catch { $lastBoot = $null }
            }

            try {
                $installDate = [Management.ManagementDateTimeConverter]::ToDateTime($os.InstallDate)
            } catch {
                try { $installDate = Get-Date ($os.InstallDate) } catch { $installDate = $null }
            }

            if ($lastBoot) { $uptime = (Get-Date) - $lastBoot } else { $uptime = $null }
            
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
            if ($outputBox.InvokeRequired) {
                $outputBox.Invoke([Action]{ $outputBox.AppendText($dashboard) })
            } else {
                $outputBox.AppendText($dashboard)
            }
        } catch {
            Log "System dashboard error: $_" "FAIL"
        } finally {
            Set-ProgressBarVisibility -Visible $false
        }
    }
$btnSystemDash.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnSystemDash)
$yPosition += 52

# Helper functions for common tasks
function Show-TopMemoryProcesses {
    Log "Identifying top memory-consuming processes..." "INFO"
    try {
        $procs = Get-Process | Sort-Object WorkingSet -Descending | Select-Object -First 20
        foreach ($p in $procs) {
            $memMB = [math]::Round($p.WorkingSet/1MB,2)
            Log "$($p.ProcessName) (PID $($p.Id)) - $memMB MB" "INFO"
        }
    } catch {
        Log "Error enumerating processes: $_" "FAIL"
    }
}

function Find-LargeFiles {
    param([string]$Path = 'C:\', [int]$MinMB = 100)
    Log "Searching for files >= $MinMB MB under $Path (may take time)..." "INFO"
    try {
        $minBytes = $MinMB * 1MB
        Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue | 
            Where-Object { -not $_.PSIsContainer -and $_.Length -ge $minBytes } | 
            Sort-Object Length -Descending | 
            Select-Object -First 50 | 
            ForEach-Object {
                $sizeMB = [math]::Round($_.Length/1MB,2)
                Log "$($_.FullName) - $sizeMB MB" "INFO"
            }
    } catch {
        Log "Large file search error: $_" "FAIL"
    }
}

function Clear-TempFiles {
    Log "Cleaning temp folders..." "INFO"
    try {
        $paths = @($env:TEMP, "$env:windir\Temp") | Where-Object { Test-Path $_ }
        foreach ($p in $paths) {
            Get-ChildItem -Path $p -Recurse -Force -ErrorAction SilentlyContinue | 
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            Log "Cleared: $p" "PASS"
        }
    } catch {
        Log "Temp cleanup error: $_" "FAIL"
    }
}

function Start-CHKDSKScan {
    Log "Scheduling CHKDSK scan (non-destructive)..." "INFO"
    Start-BackgroundTask -Name 'CHKDSK' -Action {
        try {
            $cmd = Get-Command chkdsk -ErrorAction SilentlyContinue
            if (-not $cmd) { Log "chkdsk not available on this system" "FAIL"; return }
            $result = Invoke-ExternalCommand -Command $cmd.Source -Args 'C:'
            if ($result.Output) { $result.Output | ForEach-Object { Log $_ "INFO" } }
            Log "CHKDSK scan completed. Review output above for errors." "PASS"
        } catch {
            Log "CHKDSK error: $_" "FAIL"
        }
    }
}

function Test-MappedDrives {
    Log "Testing mapped drive connectivity..." "INFO"
    try {
        $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -like "\\*" }
        if ($drives) {
            foreach ($drive in $drives) {
                if (Test-Path $drive.Root) {
                    Log "$($drive.Name): $($drive.Root) - Connected" "PASS"
                } else {
                    Log "$($drive.Name): $($drive.Root) - Disconnected" "FAIL"
                }
            }
        } else {
            Log "No mapped drives found" "INFO"
        }
    } catch {
        Log "Mapped drive test error: $_" "FAIL"
    }
}

function Show-NetworkTraffic {
    Log "Monitoring network traffic (5 second sample)..." "INFO"
    try {
        $before = Get-NetAdapterStatistics -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 5
        $after = Get-NetAdapterStatistics -ErrorAction SilentlyContinue
        
        foreach ($b in $before) {
            $a = $after | Where-Object { $_.Name -eq $b.Name }
            if ($a) {
                $rxKBps = [math]::Round(($a.ReceivedBytes - $b.ReceivedBytes) / 5 / 1KB, 2)
                $txKBps = [math]::Round(($a.SentBytes - $b.SentBytes) / 5 / 1KB, 2)
                Log "$($b.Name): Download $rxKBps KB/s, Upload $txKBps KB/s" "INFO"
            }
        }
    } catch {
        Log "Network traffic monitor error: $_" "FAIL"
    }
}

function Get-WiFiPasswords {
    Log "Recovering saved WiFi passwords..." "INFO"
    try {
        $profiles = (netsh wlan show profiles) | Select-String "All User Profile" | ForEach-Object { ($_ -split ":")[-1].Trim() }
        foreach ($wifiProfile in $profiles) {
            $key = (netsh wlan show profile name="$wifiProfile" key=clear) | Select-String "Key Content"
            if ($key) {
                $password = ($key -split ":")[-1].Trim()
                Log "$wifiProfile : $password" "PASS"
            } else {
                Log "$wifiProfile : No password stored" "INFO"
            }
        }
    } catch {
        Log "WiFi password recovery error: $_" "FAIL"
    }
}

function Reset-NetworkStack {
    if (-not (Show-Confirmation -Message "This will reset network stack and may disconnect active sessions. Continue?")) { return }
    Log "Resetting network stack..." "INFO"
    Start-BackgroundTask -Name 'Network Reset' -Action {
        try {
            $netsh = Get-Command netsh -ErrorAction SilentlyContinue
            if ($netsh) { Invoke-ExternalCommand -Command $netsh.Source -Args @('winsock','reset') | Out-Null }
            if ($netsh) { Invoke-ExternalCommand -Command $netsh.Source -Args @('int','ip','reset') | Out-Null }

            $ipcmd = Get-Command ipconfig -ErrorAction SilentlyContinue
            if ($ipcmd) { Invoke-ExternalCommand -Command $ipcmd.Source -Args '/flushdns' | Out-Null }
            if ($ipcmd) { Invoke-ExternalCommand -Command $ipcmd.Source -Args '/release' | Out-Null }
            if ($ipcmd) { Invoke-ExternalCommand -Command $ipcmd.Source -Args '/renew' | Out-Null }
            Log "Network stack reset complete. Reboot recommended." "PASS"
        } catch {
            Log "Network stack reset error: $_" "FAIL"
        }
    }
}

function Show-IPAddresses {
    Log "Gathering IP address information..." "INFO"
    try {
        # Local IPs
        Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | 
            Where-Object { $_.IPAddress -ne '127.0.0.1' } | 
            ForEach-Object {
                Log "Local IP: $($_.IPAddress) (Interface: $($_.InterfaceAlias))" "INFO"
            }
        
        # Public IP
        try {
            $publicIP = Invoke-RestMethod -Uri 'https://api.ipify.org?format=text' -UseBasicParsing -ErrorAction Stop -TimeoutSec 5
            Log "Public IP: $publicIP" "PASS"
        } catch {
            Log "Unable to retrieve public IP (network may be down)" "WARN"
        }
    } catch {
        Log "IP address retrieval error: $_" "FAIL"
    }
}

function Clear-PrintSpooler {
    Log "Clearing print spooler..." "INFO"
    try {
        Stop-Service -Name Spooler -Force -ErrorAction SilentlyContinue
        $spoolPath = Join-Path $env:windir 'System32\spool\PRINTERS'
        if (Test-Path $spoolPath) {
            Get-ChildItem $spoolPath -Recurse -Force -ErrorAction SilentlyContinue | 
                Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
        }
        Start-Service -Name Spooler -ErrorAction SilentlyContinue
        Log "Print spooler cleared and restarted" "PASS"
    } catch {
        Log "Print spooler error: $_" "FAIL"
    }
}

function Show-PrinterInventory {
    Log "Listing installed printers..." "INFO"
    try {
        Get-CimInstance Win32_Printer | ForEach-Object {
            $status = switch ($_.PrinterStatus) {
                1 { "Other" }
                2 { "Unknown" }
                3 { "Idle" }
                4 { "Printing" }
                5 { "Warmup" }
                6 { "Stopped Printing" }
                7 { "Offline" }
                8 { "Paused" }
                9 { "Error" }
                10 { "Busy" }
                11 { "Not Available" }
                12 { "Waiting" }
                13 { "Processing" }
                14 { "Initializing" }
                15 { "Warming Up" }
                default { "Status: $($_.PrinterStatus)" }
            }
            $offline = if ($_.WorkOffline) { "OFFLINE" } else { "Online" }
            Log "$($_.Name) - $status - $offline - Default: $($_.Default)" "INFO"
        }
    } catch {
        Log "Printer inventory error: $_" "FAIL"
    }
}

function Remove-GhostPrinters {
    Log "Removing offline/ghost printers..." "INFO"
    try {
        $printers = Get-CimInstance Win32_Printer | Where-Object { $_.WorkOffline -eq $true -and $_.Default -ne $true }
        if ($printers) {
            foreach ($printer in $printers) {
                try {
                    Log "Removing: $($printer.Name)" "INFO"
                    $printer.Delete() | Out-Null
                    Log "Removed $($printer.Name)" "PASS"
                } catch {
                    Log "Failed to remove $($printer.Name): $_" "WARN"
                }
            }
        } else {
            Log "No ghost printers found" "INFO"
        }
    } catch {
        Log "Ghost printer removal error: $_" "FAIL"
    }
}

function Get-SMARTStatus {
    Log "Checking disk SMART status..." "INFO"
    try {
        $disks = Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction SilentlyContinue
        if ($disks) {
            foreach ($disk in $disks) {
                $status = if ($disk.PredictFailure) { "WARNING - Failure predicted!" } else { "Healthy" }
                Log "Disk $($disk.InstanceName): $status" $(if ($disk.PredictFailure) { "WARN" } else { "PASS" })
            }
        } else {
            Log "SMART data not available (may require specific drivers)" "WARN"
        }
    } catch {
        Log "SMART status check error: $_" "FAIL"
    }
}

function Show-DiskHealth {
    Log "Generating disk health dashboard..." "INFO"
    try {
        $disks = Get-PhysicalDisk -ErrorAction SilentlyContinue
        if ($disks) {
            foreach ($disk in $disks) {
                $sizeGB = [math]::Round($disk.Size / 1GB, 2)
                Log "Disk $($disk.DeviceId): $($disk.FriendlyName) - $sizeGB GB - Health: $($disk.HealthStatus) - Op: $($disk.OperationalStatus)" "INFO"
            }
        }
        
        $volumes = Get-Volume -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter }
        foreach ($vol in $volumes) {
            $sizeGB = [math]::Round($vol.Size / 1GB, 2)
            $freeGB = [math]::Round($vol.SizeRemaining / 1GB, 2)
            $percentFree = [math]::Round(($vol.SizeRemaining / $vol.Size) * 100, 1)
            Log "Volume $($vol.DriveLetter): $freeGB GB free of $sizeGB GB ($percentFree% free) - Health: $($vol.HealthStatus)" "INFO"
        }
    } catch {
        Log "Disk health dashboard error: $_" "FAIL"
    }
}

function Reset-WindowsUpdate {
    if (-not (Show-Confirmation -Message "This will reset Windows Update components. Continue?")) { return }
    Log "Resetting Windows Update components..." "INFO"
    try {
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
        Stop-Service -Name cryptsvc -Force -ErrorAction SilentlyContinue
        Stop-Service -Name bits -Force -ErrorAction SilentlyContinue
        Stop-Service -Name msiserver -Force -ErrorAction SilentlyContinue
        
        $softDist = Join-Path $env:windir 'SoftwareDistribution'
        $catroot = Join-Path $env:windir 'System32\catroot2'
        
        if (Test-Path $softDist) { 
            Rename-Item $softDist "$softDist.old" -ErrorAction SilentlyContinue 
        }
        if (Test-Path $catroot) { 
            Rename-Item $catroot "$catroot.old" -ErrorAction SilentlyContinue 
        }
        
        Start-Service -Name wuauserv -ErrorAction SilentlyContinue
        Start-Service -Name cryptsvc -ErrorAction SilentlyContinue
        Start-Service -Name bits -ErrorAction SilentlyContinue
        Start-Service -Name msiserver -ErrorAction SilentlyContinue
        
        Log "Windows Update components reset. Reboot recommended before retrying updates." "PASS"
    } catch {
        Log "Windows Update reset error: $_" "FAIL"
    }
}

function Show-UpdateHistory {
    Log "Retrieving Windows Update history..." "INFO"
    try {
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $historyCount = $searcher.GetTotalHistoryCount()
        $updates = $searcher.QueryHistory(0, [Math]::Min(5, $historyCount))
        
        foreach ($update in $updates) {
            $date = $update.Date.ToString("yyyy-MM-dd HH:mm")
            $result = switch ($update.ResultCode) {
                0 { "Not Started" }
                1 { "In Progress" }
                2 { "Succeeded" }
                3 { "Succeeded with Errors" }
                4 { "Failed" }
                5 { "Aborted" }
                default { "Unknown" }
            }
            Log "$date - $($update.Title) - $result" "INFO"
        }
    } catch {
        Log "Update history error: $_" "FAIL"
    }
}

function Get-BSODEvents {
    Log "Scanning for recent Blue Screen events..." "INFO"
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='System'; ID=1001,1003; ProviderName='BugCheck'} -MaxEvents 10 -ErrorAction SilentlyContinue
        if ($events) {
            foreach ($evt in $events) {
                $msgSnippet = $null
                try { $msgSnippet = $evt.Message.Substring(0, [Math]::Min(200, $evt.Message.Length)) } catch { $msgSnippet = $evt.Message }
                Log "$($evt.TimeCreated) - BSOD detected: $msgSnippet" "WARN"
            }
        } else {
            Log "No recent Blue Screen events found" "PASS"
        }
    } catch {
        Log "No BSOD events found or error accessing event log" "INFO"
    }
}

function New-BatteryReport {
    Log "Scheduling battery health report..." "INFO"
    Start-BackgroundTask -Name 'Battery Report' -Action {
        try {
            $reportPath = "$env:USERPROFILE\Desktop\battery-report.html"
            $p = Get-Command powercfg -ErrorAction SilentlyContinue
            if (-not $p) { Log "powercfg not available on this system" "FAIL"; return }

            $res = Invoke-ExternalCommand -Command $p.Source -Args @('/batteryreport','/output',$reportPath)
            if ($res.ExitCode -eq 0 -and (Test-Path $reportPath)) {
                Log "Battery report saved to: $reportPath" "PASS"
                Start-Process $reportPath
            } else {
                Log "Battery report generation failed (may not be a laptop or permission issue)" "WARN"
            }
        } catch {
            Log "Battery report error: $_" "FAIL"
        }
    }
}

function Test-DomainConnection {
    Log "Testing domain connection..." "INFO"
    try {
        $computerSystem = Get-CimInstance Win32_ComputerSystem
        if ($computerSystem.PartOfDomain) {
            $result = Test-ComputerSecureChannel -ErrorAction Stop
            if ($result) {
                Log "Secure channel to domain controller is healthy" "PASS"
            } else {
                Log "Secure channel to domain controller is broken" "FAIL"
                Log "Run 'Test-ComputerSecureChannel -Repair' to fix" "INFO"
            }
        } else {
            Log "Computer is not domain-joined" "INFO"
        }
    } catch {
        Log "Domain connection test error: $_" "FAIL"
    }
}

function Update-GroupPolicy {
    Log "Forcing Group Policy update..." "INFO"
    try {
        $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($computerSystem -and -not $computerSystem.PartOfDomain) {
            Log "Computer is not domain-joined; gpupdate may have limited effect" "WARN"
        }

        # Prefer Invoke-GPUpdate when available (part of GroupPolicy module)
        if (Get-Command Invoke-GPUpdate -ErrorAction SilentlyContinue) {
            try {
                Log "Using Invoke-GPUpdate (GroupPolicy module)" "INFO"
                Invoke-GPUpdate -RandomDelayInMinutes 0 -Force -ErrorAction Stop -Confirm:$false
                Log "Group Policy update requested via Invoke-GPUpdate" "PASS"
                return
            } catch {
                Log "Invoke-GPUpdate failed: $_" "WARN"
                # Fall through to gpupdate.exe
            }
        }

        $gpCmd = Get-Command gpupdate.exe -ErrorAction SilentlyContinue
        $gpPath = if ($gpCmd) { $gpCmd.Source } else { Join-Path $env:windir 'System32\gpupdate.exe' }

        if (-not (Test-Path $gpPath)) {
            $sysnative = Join-Path $env:windir 'Sysnative\gpupdate.exe'
            $wow64 = Join-Path $env:windir 'SysWOW64\gpupdate.exe'
            if (Test-Path $sysnative) { $gpPath = $sysnative }
            elseif (Test-Path $wow64) { $gpPath = $wow64 }
        }

        if (-not (Test-Path $gpPath)) {
            Log "gpupdate.exe not found at $gpPath" "FAIL"
            Log "Ensure Group Policy tools are installed on this system." "INFO"
            return
        }

        Log "Running: $gpPath /force" "INFO"
        $output = & $gpPath /force 2>&1
        $exitCode = $LASTEXITCODE

        if ($output) { $output | ForEach-Object { Log $_ "INFO" } }

        if ($exitCode -eq 0) {
            Log "Group Policy update completed successfully (exit code 0)" "PASS"
        } else {
            Log "Group Policy update exited with code $exitCode" "FAIL"
            Log "If failures persist, check network, DNS, domain controller reachability, and credentials." "INFO"
        }
    } catch {
        Log "Group Policy update error: $_" "FAIL"
    }
}

function Show-LastLogon {
    Log "Retrieving last logon information..." "INFO"
    try {
        $lastLogon = $null
        $profiles = Get-CimInstance -ClassName Win32_NetworkLoginProfile -ErrorAction SilentlyContinue
        if ($profiles) {
            $profilesSafe = foreach ($p in $profiles) {
                $dt = $null
                try {
                    if ($p.PSObject.Properties['LastLogon']) {
                        $raw = $p.LastLogon
                        if ($raw) {
                            try { $dt = [Management.ManagementDateTimeConverter]::ToDateTime($raw) } catch { $dt = $null }
                        }
                    }
                } catch {
                    $dt = $null
                }
                [PSCustomObject]@{ Profile = $p; LastLogonDT = $dt }
            }

            $best = $profilesSafe | Where-Object { $_.LastLogonDT -ne $null } | Sort-Object -Property LastLogonDT -Descending | Select-Object -First 1
            if ($best) { $lastLogon = $best.Profile } else { $lastLogon = $null }
        }
        
        if ($lastLogon -and $lastLogon.LastLogon) {
            $logonTime = $null
            try {
                $logonTime = [Management.ManagementDateTimeConverter]::ToDateTime($lastLogon.LastLogon)
            } catch {
                try {
                    $logonTime = Get-Date $lastLogon.LastLogon
                } catch {
                    $logonTime = $lastLogon.LastLogon
                    Log "Last logon value invalid/unparseable: $($lastLogon.LastLogon)" "WARN"
                }
            }

            Log "Last User: $($lastLogon.Name)" "INFO"
            Log ("Last Logon: {0}" -f $logonTime) "INFO"
        } else {
            Log "Last logon information not available" "WARN"
        }
    } catch {
        Log "Last logon retrieval error: $_" "FAIL"
    }
}

# LOG MANAGEMENT
$buttonContainer.Controls.Add((New-SectionHeader "10. LOG MANAGEMENT"))
$yPosition += 30

$btnClearLog = New-ToolkitButton -Text "10.1 Clear Log" -Name "btnClearLog" `
    -BackColor ([System.Drawing.Color]::WhiteSmoke) `
    -ToolTipText "Clear all log entries from the console" `
    -Action {
        $outputBox.Clear()
        Log "Log cleared by user" "INFO"
    }
$btnClearLog.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnClearLog)
$yPosition += 47

$btnSaveLog = New-ToolkitButton -Text "10.2 Save Log to Desktop" -Name "btnSaveLog" `
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

$btnCopyLog = New-ToolkitButton -Text "10.3 Copy Log to Clipboard" -Name "btnCopyLog" `
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
# Adjust button container height dynamically based on content
$yPosition += 47

# Ensure container height fits content
$buttonContainer.Height = $yPosition + 50
# Startup sequence
Log "MSP Technician Toolkit - Professional Edition" "PASS"
Log "Logged in as: $env:USERNAME on $env:COMPUTERNAME" "INFO"
Log "Running with Administrator privileges" "PASS"

try {
    $domain = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Domain
    Log "Domain/Workgroup: $domain" "INFO"
} catch {
    Log "Could not determine domain status: $_" "WARN"
}

Log "PowerShell Version: $($PSVersionTable.PSVersion)" "INFO"

try {
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    Log "OS Version: $($osInfo.Caption) (Build $($osInfo.BuildNumber))" "INFO"
} catch {
    try {
        $osVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
        Log "OS Version: $osVersion (Fallback detection)" "INFO"
    } catch {
        Log "Could not determine OS version" "WARN"
    }
}

Log "Toolkit loaded successfully" "PASS"
Log "Ready for troubleshooting operations." "PASS"
Log " " "INFO"

# Safe shutdown: stop background workers, kill external processes, dispose controls
$form.Add_FormClosing({
    param($frmSender,$frmArgs)
    try {
        Log "Shutting down toolkit and stopping background tasks..." "INFO"

        $workersCopy = $null
        [System.Threading.Monitor]::Enter($script:BackgroundWorkersLock)
        try { $workersCopy = @($script:BackgroundWorkers) } finally { [System.Threading.Monitor]::Exit($script:BackgroundWorkersLock) }
        foreach ($bw in $workersCopy) {
            try {
                if ($bw -and $bw.IsBusy -and $bw.WorkerSupportsCancellation) { $bw.CancelAsync() }
            } catch {}
            try { $bw.Dispose() } catch {}
        }

        $procsCopy = $null
        [System.Threading.Monitor]::Enter($script:ExternalProcessesLock)
        try { $procsCopy = @($script:ExternalProcesses) } finally { [System.Threading.Monitor]::Exit($script:ExternalProcessesLock) }
        foreach ($proc in $procsCopy) {
            try {
                if ($proc -and -not $proc.HasExited) { $proc.Kill() ; $proc.WaitForExit(2000) }
            } catch {}
            try { $proc.Dispose() } catch {}
        }

                            # Dispose top-level controls
        try { $outputBox.Dispose() } catch {}
        try { $statusBar.Dispose() } catch {}
        try { $sidebarPanel.Dispose() } catch {}
        try { $buttonContainer.Dispose() } catch {}
        try { $consolePanel.Dispose() } catch {}
                            try { [GC]::Collect(); [GC]::WaitForPendingFinalizers() } catch {}
    } catch {
        Log "Error during shutdown: $_" "WARN"
    }
})

# Show form
$form.Add_Shown({$form.Activate()})
try {
    [void]$form.ShowDialog()
    } catch {
        Write-Host "UI runtime error: $_"
        exit 1
    }