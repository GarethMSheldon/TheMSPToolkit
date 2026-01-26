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
Version: 3.1.2
Date: January 23, 2026
Requires: PowerShell 3.0+, Administrator privileges
Compatible: Windows 7 SP1 through Windows 11 / Server 2022
#>

# Auto-elevation block
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$($myinvocation.mycommand.definition)`""
    Start-Process powershell -Verb RunAs -ArgumentList $arguments
    exit
}
# Include release notes for this snapshot (from change_3.1.2.ps1)
if (Test-Path (Join-Path $PSScriptRoot 'change_3.1.2.ps1')) { . (Join-Path $PSScriptRoot 'change_3.1.2.ps1') }

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

# Core logging function with color coding
function Log([string]$msg, [string]$status="INFO") {
    try {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $color = "White"
        if($status -eq "PASS"){$color="Lime"}
        if($status -eq "FAIL"){$color="OrangeRed"}
        if($status -eq "WARN"){$color="Yellow"}
        
        # Thread-safe UI update
        if ($outputBox.InvokeRequired) {
            $outputBox.Invoke([Action]{ 
                $outputBox.SelectionColor = [System.Drawing.Color]::$color
                $outputBox.AppendText("[$timestamp][$status] $msg`r`n")
                $outputBox.ScrollToCaret()
            })
        } else {
            $outputBox.SelectionColor = [System.Drawing.Color]::$color
            $outputBox.AppendText("[$timestamp][$status] $msg`r`n")
            $outputBox.ScrollToCaret()
        }
        [System.Windows.Forms.Application]::DoEvents()
    } catch {
        try {
            Write-Host "[$timestamp][$status] $msg" -ForegroundColor $color
        } catch {
            Write-Host "[$timestamp][$status] $msg"
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
    $button.Add_Click($Action)
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
        if (Show-Confirmation -Message "This operation takes 10-15 minutes and requires system file access. Continue?") {
            try {
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
            } catch {
                Log "SFC scan error: $_" "FAIL"
            } finally {
                Set-ProgressBarVisibility -Visible $false
            }
        }
    }
$btnSFC.Location = New-Object System.Drawing.Point(10, $yPosition)
$buttonContainer.Controls.Add($btnSFC)
$yPosition += 47

# (rest of file preserved)
