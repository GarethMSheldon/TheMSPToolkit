<#
.NOTES
Author: MSP Solutions Team
Version: 3.0
Date: November 15, 2025
Requires: PowerShell 3.0+, Administrator privileges
#>

# Include release notes for this snapshot (from change_3.0.0.ps1)
if (Test-Path (Join-Path $PSScriptRoot 'change_3.0.0.ps1')) { . (Join-Path $PSScriptRoot 'change_3.0.0.ps1') }
# Require Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# Load UI Assemblies
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Main Form Setup ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "MSP Technician IDE - Advanced Edition (Running as Admin)"
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

# (rest of file preserved)
