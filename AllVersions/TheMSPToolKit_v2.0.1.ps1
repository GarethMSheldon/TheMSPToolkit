#Requires -RunAsAdministrator
<#
.SYNOPSIS
MSP Technician Toolkit — v2.0.1 snapshot
.DESCRIPTION
Toolkit snapshot for v2.0.1 (post-release fixes and additions).
.NOTES
Author: MSP Solutions Team
Version: 2.0.1
Date: 2025-05-10
Requires: PowerShell 3.0+, Administrator privileges
Compatible: Windows 7 SP1 through Windows 11 / Server 2022
#>

# Include release notes for this snapshot (from change_2.0.1.ps1)
if (Test-Path (Join-Path $PSScriptRoot 'change_2.0.1.ps1')) { . (Join-Path $PSScriptRoot 'change_2.0.1.ps1') }

# Require Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
	Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
	Exit
}

# Load UI Assemblies
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Main Form Setup ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "MSP Technician IDE - v2.0.1 (Running as Admin)"
$form.Size = New-Object System.Drawing.Size(1150, 850)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

# --- Global Output Box ---
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Dock = "Right"
$outputBox.Width = 600
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$outputBox.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$outputBox.ForeColor = [System.Drawing.Color]::Gainsboro
$form.Controls.Add($outputBox)

function Log([string]$msg, [string]$status="INFO") {
	$color = "White"
	if($status -eq "PASS"){$color="Lime"}
	if($status -eq "FAIL"){$color="OrangeRed"}
	if($status -eq "WARN"){$color="Yellow"}
	$outputBox.SelectionColor = [System.Drawing.Color]::$color
	$outputBox.AppendText("[$status] $msg`r`n")
	$outputBox.ScrollToCaret()
}

# --- Sidebar Panel ---
$sidebar = New-Object System.Windows.Forms.FlowLayoutPanel
$sidebar.Dock = "Left"
$sidebar.Width = 520
$sidebar.FlowDirection = "TopDown"
$sidebar.AutoScroll = $true
$sidebar.Padding = New-Object System.Windows.Forms.Padding(10)
$sidebar.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 235)
$form.Controls.Add($sidebar)

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

# --- SECTION: OS REPAIR ---
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

# (trimmed: remaining UI mirrors v2.0 features)

$form.ShowDialog()
